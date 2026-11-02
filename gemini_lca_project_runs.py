#!/usr/bin/env python3
"""
Gemini embodied-carbon (A-C) range estimation, one building per call, N repeated runs per project.

For every row (project) in buildings_metadata_LLM.xlsx the script calls Gemini `--runs` times
(default 10) with the same prompt and writes one file per project:

    <results-root>/<run-name>/project-<project_index>_runs.csv

Prompt caching (explicit)
-------------------------
Everything that is identical for all 2,920 calls is put into ONE Gemini context cache before
the loop starts:

    system instruction
    1. "Forget the information from previous sessions ..."
    2. background papers
    3. CLF practice guide
    4. data glossary
    5. "Reference the given feature information above. Estimate ... the building below:\n\n"

Each call then sends only the building-specific part (building fields + answer-format
instruction) and points at the cache. The cache is refreshed while the job runs and deleted at
the end. Set USE_CACHE = False to fall back to Gemini's automatic (implicit) caching.

Install:  pip install google-genai pandas openpyxl PyPDF2 python-dotenv
Key:      GEMINI_API_KEY in the environment or a .env file

Examples
  python gemini_lca_project_runs.py --projects 4             # one project, 10 runs
  python gemini_lca_project_runs.py --projects 1-10          # first 10 projects
  python gemini_lca_project_runs.py --runs 4                 # all projects, 4 runs each
  python gemini_lca_project_runs.py                          # all projects, 10 runs each
  python gemini_lca_project_runs.py --run-name folder-name   # name the folder or continue running

Model, caching, workers, resume folder etc. are set in the "settings" blocks below.
"""

from __future__ import annotations

import argparse
import csv
import os
import random
import re
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime
from pathlib import Path

import pandas as pd
import PyPDF2
from dotenv import load_dotenv

load_dotenv()

# --------------------------------------------------------------------------------------
# Model settings - edit here (command-line flags can still override them)
# --------------------------------------------------------------------------------------
GEMINI_MODEL = "gemini-3.8-flash"
TEMPERATURE = 1.0
THINKING_LEVEL = "medium"          # "low" | "medium" | "high" | None (model default)
MAX_OUTPUT_TOKENS = 16384          # includes thinking tokens
CACHE_TTL_S = 3600                 # cache lifetime; refreshed automatically during the run

# --------------------------------------------------------------------------------------
# Run settings - edit here
# --------------------------------------------------------------------------------------
RUN_NAME = None          # None -> "yyyymmdd_HH_<model>"; set to an existing folder name to resume it
WORKERS = 64              # projects processed in parallel (8-16 for a full run)
USE_CACHE = True         # explicit context cache for the shared prompt prefix
PDF_MODE = "text"        # "text" (PyPDF2 extraction) or "native" (upload PDFs to Gemini)
RETRY_FAILED = False     # True -> also redo runs with parse_ok=False
DELAY_S = 0.0            # pause between calls per worker
ONTOLOGY_MODE = "none"   # label written to the CSV column
GEOMETRY_MODE = "none"   # label written to the CSV column
DRY_RUN = False          # True -> no API calls, fake responses (offline test)

# --------------------------------------------------------------------------------------
# Paths - all relative to the folder this script is in (CEE226-reference)
# --------------------------------------------------------------------------------------
SCRIPT_DIR = Path(__file__).resolve().parent

BASE_DIR = SCRIPT_DIR
GLOSSARY_PATH = BASE_DIR / "references" / "data_glossary_LLM.xlsx"
BACKGROUND_PATHS = [BASE_DIR / "references" / f"{i}.pdf" for i in range(1, 5)]   # 1.pdf to 4.pdf, as in gemini_lca_batch.py
PROCESS_PATH = BASE_DIR / "references" / "250524_CLF-LCA-Practice-Guide.txt"
BENCHMARK_PATH = BASE_DIR / "references" / "eci_benchmarks_A-C.md"

PROJECT_DIR = SCRIPT_DIR
METADATA_PATH = PROJECT_DIR / "references" / "buildings_metadata_LLM.xlsx"
RESULTS_ROOT = PROJECT_DIR / "results"

MAX_PDF_CHARS = 100_000
SYSTEM_PROMPT = "You are a building LCA expert."

STOP = threading.Event()   # set by Ctrl+C so workers finish their current call and stop

CSV_COLUMNS = [
    "project", "run_idx", "low", "high", "mid", "parse_ok", "failure_type",
    "ontology_mode", "geometry_mode", "finish_reason", "output_tokens", "thought_tokens",
    "cached_tokens", "prompt_tokens", "latency_s", "model", "temperature",
    "thinking_setting", "timestamp",
]


# --------------------------------------------------------------------------------------
# Readers
# --------------------------------------------------------------------------------------
def read_excel_text(path: Path) -> str:
    try:
        return pd.read_excel(path).to_string(index=False)
    except Exception as e:
        print(f"Warning: could not read {path}: {e}")
        return ""


def read_pdf(path: Path, max_chars: int = MAX_PDF_CHARS) -> str:
    if not path.exists():
        print(f"Warning: PDF not found: {path}")
        return ""
    try:
        with open(path, "rb") as f:
            text = "\n".join((p.extract_text() or "") for p in PyPDF2.PdfReader(f).pages)
        return text[:max_chars] + "...[truncated]" if len(text) > max_chars else text
    except Exception as e:
        print(f"Warning: could not read {path}: {e}")
        return ""


def read_txt(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except Exception as e:
        print(f"Warning: could not read {path}: {e}")
        return ""


# --------------------------------------------------------------------------------------
# Prompt: cached prefix + per-building suffix
# --------------------------------------------------------------------------------------
def build_cached_prefix(background_text: str, process_text: str, glossary_text: str,
                        benchmark_text: str) -> list[str]:
    """Identical for every call -> goes into the context cache. Same text as gemini_lca_batch.py."""
    return [
        "Forget the information from previous sessions and calculate everything from the beginning.",

        f"data_glossary_LLM contains feature names and descriptions for the data:\n\n{glossary_text}",

        "These journal papers contain background information about the early building LCA "
        f"process:\n\n{background_text}",

        "This document contains a detailed practice guide on life cycle assessment of buildings "
        f"focused on implementation. Use this guideline for LCA estimation:\n\n{process_text}",

        # "The following benchmark tables give published embodied carbon intensity (ECI) "
        # f"distributions. Use them as described in the notes below the tables:\n\n{benchmark_text}",

        "Reference the given feature information above. Estimate the embodied carbon value of the "
        "building below:\n\n",
    ]


def build_building_prompt(building_text: str) -> str:
    """Different for every project -> sent with each call."""
    return (
        f"{building_text}\n\n"
        "Provide a plausible range in units of kgCO2e/m2 (a range, not a single value). For floor "
        "area, you need to use bldg_cfa. If cfa is not provided, use bldg_renovated_cfa + "
        "bldg_renovated_gfa. Include life cycle stages A-C. Briefly state your method and key "
        "assumptions.\n\n"
        "End your answer with exactly one final line in this format and nothing after it:\n"
        "RANGE_kgCO2e_per_m2: <low> - <high>\n"
        "(plain numbers only, no units or other text on that line)"
    )


def format_building(row: pd.Series) -> str:
    """One 'field: value' line per non-empty column. project_index is the ID, not a feature."""
    lines = []
    for col, val in row.items():
        if col == "project_index" or pd.isna(val):
            continue
        if isinstance(val, float) and val.is_integer():
            val = int(val)
        lines.append(f"{col}: {val}")
    return "\n".join(lines)


# --------------------------------------------------------------------------------------
# Response parsing
# --------------------------------------------------------------------------------------
_RANGE_RE = re.compile(
    r"RANGE_kgCO2e_per_m2\s*:\s*([0-9][0-9,]*\.?[0-9]*)\s*(?:-|\u2013|\u2014|to)\s*([0-9][0-9,]*\.?[0-9]*)",
    re.IGNORECASE,
)


def parse_range(text: str) -> tuple[float | None, float | None, str]:
    """Return (low, high, failure_type). Uses the LAST matching line; failure_type '' on success."""
    clean = text.replace("*", "").replace("`", "")
    matches = _RANGE_RE.findall(clean)
    if not matches:
        return None, None, "no_range"
    try:
        low, high = (float(x.replace(",", "").rstrip(".")) for x in matches[-1])
    except ValueError:
        return None, None, "no_range"
    if low <= 0 or high < low:
        return low, high, "invalid_range"
    return low, high, ""


# --------------------------------------------------------------------------------------
# Backends
# --------------------------------------------------------------------------------------
def with_retries(fn, max_attempts: int = 5, base_delay: float = 5.0):
    for attempt in range(1, max_attempts + 1):
        try:
            return fn()
        except Exception as e:
            msg = str(e)
            transient = any(c in msg for c in ("429", "500", "502", "503", "504",
                                               "RESOURCE_EXHAUSTED", "UNAVAILABLE",
                                               "DEADLINE_EXCEEDED"))
            if not transient or attempt == max_attempts:
                raise
            delay = base_delay * 2 ** (attempt - 1) + random.uniform(0, 2)
            print(f"  transient error ({msg[:100]}); retry {attempt}/{max_attempts - 1} in {delay:.0f}s")
            time.sleep(delay)


class GeminiBackend:
    def __init__(self, model, temperature, max_output_tokens, thinking_level,
                 use_cache=True, cache_ttl_s=CACHE_TTL_S):
        from google import genai
        from google.genai import types

        key = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
        if not key:
            raise RuntimeError("Set GEMINI_API_KEY (or GOOGLE_API_KEY) in your environment/.env")
        self.types = types
        self.client = genai.Client(api_key=key)
        self.model = model
        self.temperature = temperature
        self.max_output_tokens = max_output_tokens
        self.thinking_level = thinking_level
        self.use_cache = use_cache
        self.cache_ttl_s = cache_ttl_s

        self.prefix_parts: list = []
        self.cache_name: str | None = None
        self._cache_refreshed_at = 0.0
        self._lock = threading.Lock()

    # ---------- inputs ----------
    def upload_pdfs(self, paths):
        handles = []
        for p in paths:
            if not p.exists():
                print(f"Warning: PDF not found: {p}")
                continue
            print(f"Uploading {p.name} ...")
            f = self.client.files.upload(file=str(p))
            while getattr(f.state, "name", str(f.state)) == "PROCESSING":
                time.sleep(2)
                f = self.client.files.get(name=f.name)
            handles.append(f)
        return handles

    def _to_part(self, item):
        if isinstance(item, str):
            return self.types.Part.from_text(text=item)
        return self.types.Part.from_uri(file_uri=item.uri, mime_type=item.mime_type)  # uploaded file

    # ---------- explicit context cache ----------
    def prepare(self, prefix_parts: list):
        """Store the shared prefix and, if enabled, put it in an explicit context cache."""
        self.prefix_parts = prefix_parts
        if self.use_cache:
            self._create_cache()

    def _create_cache(self):
        t = self.types
        try:
            cache = self.client.caches.create(
                model=self.model,
                config=t.CreateCachedContentConfig(
                    display_name=f"lca-prefix-{datetime.now():%Y%m%d-%H%M%S}",
                    system_instruction=SYSTEM_PROMPT,
                    contents=[t.Content(role="user",
                                        parts=[self._to_part(p) for p in self.prefix_parts])],
                    ttl=f"{self.cache_ttl_s}s",
                ),
            )
            self.cache_name = cache.name
            self._cache_refreshed_at = time.time()
            n = getattr(cache.usage_metadata, "total_token_count", None)
            print(f"Context cache created: {cache.name} ({n} tokens, TTL {self.cache_ttl_s}s)")
        except Exception as e:
            self.cache_name = None
            print(f"Warning: could not create context cache ({str(e)[:200]}).\n"
                  "         Continuing without it (implicit caching may still apply).")

    def _refresh_cache_if_needed(self):
        """Extend the TTL when half of it has passed, so the cache outlives long runs."""
        if not self.cache_name:
            return
        with self._lock:
            if time.time() - self._cache_refreshed_at < self.cache_ttl_s / 2:
                return
            try:
                self.client.caches.update(
                    name=self.cache_name,
                    config=self.types.UpdateCachedContentConfig(ttl=f"{self.cache_ttl_s}s"))
                self._cache_refreshed_at = time.time()
            except Exception as e:
                print(f"  cache refresh failed ({str(e)[:120]}); recreating")
                self._create_cache()

    def close(self):
        """Delete the cache so it stops accruing storage charges."""
        if self.cache_name:
            try:
                self.client.caches.delete(name=self.cache_name)
                print(f"Context cache deleted: {self.cache_name}")
            except Exception as e:
                print(f"Warning: could not delete cache {self.cache_name}: {e}")

    # ---------- generation ----------
    def _config(self):
        t = self.types
        cfg = dict(max_output_tokens=self.max_output_tokens, temperature=self.temperature)
        if self.thinking_level:
            cfg["thinking_config"] = t.ThinkingConfig(thinking_level=self.thinking_level)
        if self.cache_name:
            cfg["cached_content"] = self.cache_name   # system instruction lives in the cache
        else:
            cfg["system_instruction"] = SYSTEM_PROMPT
        return t.GenerateContentConfig(**cfg)

    def _contents(self, building_prompt: str):
        if self.cache_name:
            return [building_prompt]                  # prefix is already in the cache
        return [*(self._to_part(p) for p in self.prefix_parts), self._to_part(building_prompt)]

    def ask(self, building_prompt: str) -> dict:
        self._refresh_cache_if_needed()

        def _call():
            t0 = time.perf_counter()
            try:
                resp = self.client.models.generate_content(
                    model=self.model, contents=self._contents(building_prompt), config=self._config())
            except Exception as e:
                msg = str(e).lower()
                if self.cache_name and "cache" in msg and ("not found" in msg or "expired" in msg
                                                          or "permission" in msg):
                    print("  cache missing/expired; recreating and retrying")
                    with self._lock:
                        self._create_cache()
                    resp = self.client.models.generate_content(
                        model=self.model, contents=self._contents(building_prompt), config=self._config())
                else:
                    raise
            latency = time.perf_counter() - t0
            u = resp.usage_metadata
            fr = resp.candidates[0].finish_reason if resp.candidates else None
            try:
                text = resp.text or ""
            except Exception:  # blocked / empty candidates
                text = ""
            return {
                "text": text,
                "finish_reason": f"FinishReason.{getattr(fr, 'name', fr)}",
                "output_tokens": getattr(u, "candidates_token_count", None) or 0,
                "thought_tokens": getattr(u, "thoughts_token_count", None) or 0,
                "cached_tokens": getattr(u, "cached_content_token_count", None) or 0,
                "prompt_tokens": getattr(u, "prompt_token_count", None) or 0,
                "latency_s": round(latency, 2),
            }
        return with_retries(_call)


class DryRunBackend:
    """No API calls: fabricates responses so the whole pipeline/CSV format can be tested offline."""

    def upload_pdfs(self, paths):
        return []

    def prepare(self, prefix_parts):
        print("(dry run) skipping context cache")

    def close(self):
        pass

    def ask(self, building_prompt: str) -> dict:
        lo = random.randrange(250, 450, 5)
        return {
            "text": f"(dry run) Method summary.\n\n**RANGE_kgCO2e_per_m2: {lo} - {lo + 90}**",
            "finish_reason": "FinishReason.STOP",
            "output_tokens": random.randint(500, 900),
            "thought_tokens": random.randint(1000, 2000),
            "cached_tokens": 8270, "prompt_tokens": 8967,
            "latency_s": round(random.uniform(7, 14), 2),
        }


# --------------------------------------------------------------------------------------
# Per-project runner
# --------------------------------------------------------------------------------------
def write_csv(path: Path, rows: list[dict]):
    rows = sorted(rows, key=lambda r: int(r["run_idx"]))
    with open(path, "w", newline="", encoding="utf-8") as f:  # default \r\n, as in the example
        w = csv.DictWriter(f, fieldnames=CSV_COLUMNS)
        w.writeheader()
        w.writerows(rows)


def load_existing(path: Path) -> list[dict]:
    if not path.exists():
        return []
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


def run_project(project_idx: int, building_text: str, backend, args, out_dir: Path) -> tuple[str, int, int]:
    name = f"project-{project_idx}"
    csv_path = out_dir / f"{name}_runs.csv"
    raw_dir = out_dir / "raw"
    raw_dir.mkdir(parents=True, exist_ok=True)

    rows = load_existing(csv_path)
    done_ok = {int(r["run_idx"]) for r in rows if str(r["parse_ok"]) == "True"}
    done_any = {int(r["run_idx"]) for r in rows}
    todo = [i for i in range(args.runs) if i not in (done_ok if args.retry_failed else done_any)]
    if not todo:
        return name, 0, 0
    # keep every existing run that is not being redone (including run_idx >= --runs)
    rows = [r for r in rows if int(r["run_idx"]) not in todo]

    building_prompt = build_building_prompt(building_text)
    thinking = f"level={args.thinking_level}" if args.thinking_level else "default"
    n_fail = 0
    for i in todo:
        if STOP.is_set():
            break
        try:
            r = backend.ask(building_prompt)
            low, high, ftype = parse_range(r["text"])
            if ftype and "MAX_TOKENS" in r["finish_reason"]:
                ftype = "truncated"
            (raw_dir / f"{name}_run{i}.md").write_text(r["text"], encoding="utf-8")
        except Exception as e:
            print(f"  [{name} run {i}] API error: {str(e)[:150]}")
            r = {"finish_reason": "", "output_tokens": "", "thought_tokens": "", "cached_tokens": "",
                 "prompt_tokens": "", "latency_s": ""}
            low = high = None
            ftype = "api_error"

        ok = ftype == ""
        n_fail += (not ok)
        rows.append({
            "project": name, "run_idx": i,
            "low": low if low is not None else "", "high": high if high is not None else "",
            "mid": (low + high) / 2 if ok else "",
            "parse_ok": ok, "failure_type": ftype,
            "ontology_mode": args.ontology_mode, "geometry_mode": args.geometry_mode,
            "finish_reason": r["finish_reason"], "output_tokens": r["output_tokens"],
            "thought_tokens": r["thought_tokens"], "cached_tokens": r["cached_tokens"],
            "prompt_tokens": r["prompt_tokens"], "latency_s": r["latency_s"],
            "model": args.model, "temperature": args.temperature, "thinking_setting": thinking,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        })
        write_csv(csv_path, rows)
        if args.delay:
            time.sleep(args.delay)
    return name, len(todo), n_fail


# --------------------------------------------------------------------------------------
# Main
# --------------------------------------------------------------------------------------
def parse_project_selection(spec: str, available: list[int]) -> list[int]:
    if spec == "all":
        return available
    chosen = set()
    for chunk in spec.split(","):
        chunk = chunk.strip()
        if "-" in chunk:
            a, b = chunk.split("-")
            chosen.update(range(int(a), int(b) + 1))
        elif chunk:
            chosen.add(int(chunk))
    missing = sorted(chosen - set(available))
    if missing:
        print(f"Warning: project_index not in metadata, skipped: {missing}")
    return sorted(chosen & set(available))


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--projects", default="all", help='"all", "4", "1-20", or "4,10,25-30"')
    ap.add_argument("--runs", type=int, default=10, help="Runs per project (default 10)")
    ap.add_argument("--run-name", default=RUN_NAME,
                    help="Results folder name (default: yyyymmdd_HH_<model>); use an existing one to resume")
    args = ap.parse_args()

    # Everything else comes from the settings at the top of the file
    args.model, args.temperature, args.thinking_level = GEMINI_MODEL, TEMPERATURE, THINKING_LEVEL
    args.retry_failed, args.delay = RETRY_FAILED, DELAY_S
    args.ontology_mode, args.geometry_mode = ONTOLOGY_MODE, GEOMETRY_MODE

    run_name = args.run_name or f"{datetime.now():%Y%m%d_%H}_{GEMINI_MODEL.replace('.', '_')}"
    out_dir = RESULTS_ROOT / run_name
    out_dir.mkdir(parents=True, exist_ok=True)
    print(f"Results folder: {out_dir}")

    meta = pd.read_excel(METADATA_PATH)
    projects = parse_project_selection(args.projects, [int(x) for x in meta["project_index"]])
    meta = meta.set_index("project_index")
    print(f"{len(projects)} project(s) x {args.runs} runs = {len(projects) * args.runs} calls, "
          f"model={GEMINI_MODEL}")

    backend = (DryRunBackend() if DRY_RUN else
               GeminiBackend(GEMINI_MODEL, TEMPERATURE, MAX_OUTPUT_TOKENS, THINKING_LEVEL,
                             use_cache=USE_CACHE))

    # Shared prefix, built once
    glossary_text = read_excel_text(GLOSSARY_PATH)
    process_text = read_txt(PROCESS_PATH)
    if PDF_MODE == "native":
        raise SystemExit('PDF_MODE = "native" is not supported with this prompt; use "text".')
    background_text = "\n".join(read_pdf(p) for p in BACKGROUND_PATHS)
    benchmark_text = read_txt(BENCHMARK_PATH) if BENCHMARK_PATH.exists() else ""
    prefix = build_cached_prefix(background_text, process_text, glossary_text,
                                 benchmark_text or "<<BENCHMARK FILE MISSING>>")
    # Stop instead of sending empty tables (only when the benchmark lines are switched on)
    if any("<<BENCHMARK FILE MISSING>>" in part for part in prefix):
        raise SystemExit(f"Benchmark lines are on, but {BENCHMARK_PATH} is missing or empty.")

    backend.prepare(prefix)

    lock = threading.Lock()
    t0, total_fail = time.time(), 0
    ex = ThreadPoolExecutor(max_workers=max(1, WORKERS))
    try:
        futs = {ex.submit(run_project, idx, format_building(meta.loc[idx]), backend, args, out_dir): idx
                for idx in projects}
        for n, fut in enumerate(as_completed(futs), 1):
            try:
                name, ran, failed = fut.result()
            except Exception as e:
                print(f"[{n}/{len(projects)}] project-{futs[fut]} crashed: {e}")
                continue
            total_fail += failed
            with lock:
                msg = "already complete" if ran == 0 else f"{ran} run(s), {failed} failed"
                print(f"[{n}/{len(projects)}] {name}: {msg}")
    except KeyboardInterrupt:
        STOP.set()
        print("\nCtrl+C received: finishing the calls in progress, then stopping. "
              "Re-run with the same --run-name to resume.")
    finally:
        ex.shutdown(wait=True, cancel_futures=True)
        backend.close()   # runs on normal finish, errors and Ctrl+C
    print(f"Done in {(time.time() - t0) / 60:.1f} min. Failed runs this session: {total_fail}")


if __name__ == "__main__":
    main()