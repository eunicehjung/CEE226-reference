# Guide: Summing GWP with `sum_gwp.py`

This guide is for anyone who needs to run `sum_gwp.py` but doesn't usually write code. You don't need to understand Python. You only need to edit a few lines in a settings block and run one command.

---

## 1. What the script does

The file `full_lca_results.xlsx` has about 928,000 rows. Each row is one material entry for one project, at one life cycle stage, with a GWP (global warming potential) value in kg CO2e.
Download full_lca_results.xlsx from the link under GWP-calculation: https://figshare.com/articles/dataset/A_Harmonized_Dataset_of_High-Resolution_Whole_Building_Life_Cycle_Assessment_Results_in_North_America_i_Data_only_-_i_i_First_Public_Release_i_/28462145/1


The script lets you choose:

1. **Which projects** (`project_index`), such as project 1, several projects, or all 292.
2. **Which building elements** (`omniclass_element`), such as substructure, superstructure, or enclosure.
3. **Which life cycle stages** (`life_cycle_stage`), such as A1-A3 only or A1 through C4.

It then adds up the GWP of every row that matches all three choices and gives you **one total per project**.

---

## 2. One-time setup

You only do this the first time.

### 2.1 Install Python

If you don't have Python, install **Anaconda** from <https://www.anaconda.com/download>. It includes Python and most of the packages the script needs.

To check that Python works, open a terminal:

- **Mac:** open the *Terminal* app.
- **Windows:** open *Anaconda Prompt* from the Start menu.

Then type:

```
python --version
```

If you see something like `Python 3.11.5`, Python is installed.

### 2.2 Install the two packages the script uses

In the same terminal, type:

```
pip install pandas openpyxl
```

`pandas` handles the table of data, and `openpyxl` lets Python read Excel files. If they are already installed, this command does nothing.

### 2.3 Put the files in one folder

Put these two files in the **same folder**, for example a folder called `lca_gwp` on your Desktop:

```
lca_gwp/
├── sum_gwp.py
└── full_lca_results.xlsx
```

The script looks for the Excel file in its own folder. If the Excel file is somewhere else, the script will not find it.

---

## 3. Choose what to sum (edit the settings)

Open `sum_gwp.py` in any text editor, such as VS Code, Notepad, TextEdit (in plain-text mode), or Spyder from Anaconda. Near the top you'll find a block marked `SETTINGS`. **This is the only part you need to change.**

### 3.1 Projects

```python
PROJECTS = [1]
```

| What you want | What to type |
|---|---|
| One project | `PROJECTS = [1]` |
| Several projects | `PROJECTS = [1, 5, 12]` |
| All projects | `PROJECTS = "all"` |

Keep the square brackets `[ ]` around numbers and the quotation marks around `"all"`.

### 3.2 Building elements

```python
ELEMENTS = [
    "Substructure",
    "Shell - Superstructure",
    "Shell - Enclosure",
    # "Interiors - Construction",
    # "Interiors - Finishes",
    # "Unknown",
]
```

A `#` at the start of a line **turns that line off**. In the example above, only Substructure, Superstructure, and Enclosure are included.

- To **include** an element, delete the `#` at the start of its line.
- To **exclude** an element, add a `#` at the start of its line.

Don't change the spelling, spaces, or hyphens inside the quotes. They must match the Excel file exactly. Also keep the comma at the end of each line.

### 3.3 Life cycle stages

```python
STAGES = [
    "A1-A3",
    "A4",
    "B4-B5",
    "C2-C4",
    # "D",
]
```

This works the same way: a `#` turns a stage off.

| Stage | Meaning |
|---|---|
| A1-A3 | Product stage: raw material supply, transport to factory, manufacturing |
| A4 | Transport to the construction site |
| B4-B5 | Replacement and refurbishment during the building's life |
| C2-C4 | End of life: transport, waste processing, disposal |
| D | Benefits beyond the building's life, such as recycling credits |

**About stage D:** its values are often negative, because they are credits. Under the EN 15978 standard, D is usually reported **separately** from A–C. When you compare projects or benchmarks, leaving D out (with `#`) is the safer choice.

### 3.4 Save the file

Save `sum_gwp.py` after editing. The script uses whatever settings are saved in the file.

---

## 4. Run the script

1. Open a terminal (Terminal on Mac, Anaconda Prompt on Windows).
2. Move into the folder with the files. Type `cd`, a space, then drag the folder into the terminal window and press Enter:
   ```
   cd /Users/yourname/Desktop/lca_gwp
   ```
3. Run:
   ```
   python sum_gwp.py
   ```

**The first run takes a few minutes** because the Excel file is large. The terminal will show:

```
Reading full_lca_results.xlsx (first run is slow; caching to full_lca_results.pkl) ...
```

Wait for it to finish. Later runs take only a few seconds (see Section 6).

When it finishes, the terminal shows something like:

```
Elements: ['Substructure', 'Shell - Superstructure', ...]
Stages:   ['A1-A3', 'A4', 'B4-B5', 'C2-C4']
Rows used: 5,123

 project_index  total_gwp_kgCO2e
             1      28,142,947.4

1 warning(s) — saved to 20261023_06_gwp-total_warnings.csv

Saved to /Users/yourname/Desktop/lca_gwp/20261023_06_gwp-total.csv
```

---

## 5. Read the results

The script creates files in the same folder. Their names start with the date and hour of the run (`YYYYMMDD_HH`).

### 5.1 `YYYYMMDD_HH_gwp-total.csv`: the totals

This is the main result. You can open it in Excel.

| project_index | total_gwp_kgCO2e |
|---|---|
| 1 | 28142947.45 |

- Units are **kg CO2e**. Divide by 1,000 for tonnes CO2e.
- Every project you asked for appears, even if nothing matched. In that case its total is `0`.

### 5.2 `YYYYMMDD_HH_gwp-total_warnings.csv`: things to know

This file is only created when there is something to report. It lists projects that have **no data** for an element or stage you selected, for example:

| project_index | issue |
|---|---|
| 1 | no rows for omniclass_element = Unknown |

A warning is **not an error**. It explains why a total may be smaller than you'd expect: the data for that category doesn't exist for that project, so there was nothing to add.

### 5.3 Record your settings

The CSV files **don't record which elements and stages you selected**. Write the settings down in your notes, or rename the output file, for example `20261023_06_gwp-total_A1-C4_noD.csv`.

**Runs in the same hour overwrite each other**, because the file name includes only the date and hour. Rename or move a result you want to keep before running again in the same hour.

---

## 6. What is `full_lca_results.pkl`?

After the first run, a new file called `full_lca_results.pkl` appears in the folder. It is a **fast copy** of the four columns the script needs from the Excel file, saved in a format Python loads quickly. This is why later runs take seconds instead of minutes.

- **You don't need to do anything with it.**
- **You can delete it anytime.** The next run will be slow once and then recreate it.
- **If you replace the Excel file** with an updated version, the script notices and rebuilds the `.pkl` automatically.
- **Don't send it to others or open `.pkl` files from other people.** Only Python can read them, and they can contain code. Share the Excel or CSV files instead.

---

## 7. Common problems

| Message in the terminal | What it means | Fix |
|---|---|---|
| `python: command not found` or `'python' is not recognized` | Python isn't installed, or you're not in Anaconda Prompt (Windows) | Install Anaconda (Section 2.1). On Mac, try `python3 sum_gwp.py`. |
| `No such file or directory: 'sum_gwp.py'` | The terminal isn't in the right folder | Repeat step 2 of Section 4 (`cd` into the folder). |
| `ModuleNotFoundError: No module named 'pandas'` (or `openpyxl`) | A package is missing | Run `pip install pandas openpyxl`. |
| `FileNotFoundError: ... full_lca_results.xlsx` | The Excel file isn't next to the script, or has a different name | Put it in the same folder and name it exactly `full_lca_results.xlsx`. |
| `ELEMENTS contains values not in the file: [...]` | A typo inside the quotes in `ELEMENTS` or `STAGES` | The message lists the valid options. Copy one exactly. |
| `SyntaxError` | A comma, bracket, or quotation mark was deleted by accident | Compare your settings block with the examples in Section 3. Every item needs quotes and a comma at the end. |
| `PermissionError` | The output CSV is open in Excel and the script can't overwrite it | Close the file in Excel and run again. |
| Total is `0` | No rows matched your choices for that project | Check the warnings file. The project may not have the elements or stages you selected. |

---

## 8. Quick checklist

1. ☐ `sum_gwp.py` and `full_lca_results.xlsx` are in the same folder.
2. ☐ `PROJECTS`, `ELEMENTS`, and `STAGES` are set, and the file is saved.
3. ☐ The terminal is in that folder (`cd ...`).
4. ☐ Run `python sum_gwp.py`.
5. ☐ Open the `_gwp-total.csv` file, and the `_warnings.csv` file if there is one.
6. ☐ Write down which settings you used.
