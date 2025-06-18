# 📓 Journal — June 18

## 🛠️ Task
Today, I was assigned to consolidate **2 months' worth of Call Detail Report (CDR)** data into a single file. The raw files were in `.xlsx` format, and each file contained approximately **500,000 rows**.

## ⚠️ Challenge
I initially tried to use **Python** (both `.ipynb` and `.py` scripts) to perform the merge and apply a filter for `"AGENT - CUSTOMER RPC"`. However, the processing was **very slow** — the script took too long to load and filter the large datasets.

## ✅ Solution
In the end, I switched to using a **VBA macro**, which successfully **consolidated the files and applied the filter in just 5 minutes**. Despite Python being my preferred tool for automation, **VBA outperformed it in this specific case**, especially with large `.xlsx` files and local Excel operations.

## 💭 Reflection
This experience reminded me that **VBA still has an edge** in certain Excel-based tasks, particularly for handling large row counts directly inside Excel.

While I’ll continue to explore optimizing my Python scripts (perhaps using `openpyxl`, `pyxlsb`, or chunk-based processing), **VBA remains a powerful fallback for massive Excel file automation**.

---

# 🧾 Work Journal — June 18, 2025

## 🛠️ Task
Merge two months’ worth of `.xlsx` Call Detail Report (CDR) files (~500,000 rows each) into a single consolidated dataset with a filter applied for:

> `"AGENT - CUSTOMER RPC"` in the `DESCRIPTION` column.

---

## ⚠️ What Happened

- Tried using **Python (`.ipynb` and `.py`)** to merge and filter `.xlsx` files.
- The process was **very slow**, especially when loading multiple large sheets per file.
- Estimated processing time: **30+ minutes** and memory spike.
- Switched to **VBA macro**, which successfully completed the task in **~10 minutes**.
- VBA handled Excel-native operations (AutoFilter, Copy-Paste) efficiently.

---

## ✅ Final Solution

Used a VBA macro to:
- Loop through `.xlsx` files in a folder.
- Apply `AutoFilter` on each worksheet for `"AGENT - CUSTOMER RPC"`.
- Copy filtered rows and paste into a summary sheet.
- Result: **Fast**, efficient, no crashes.

---

## 📚 Key Learnings

### What Slowed Python Down:
- `pandas.read_excel()` loads **entire sheet or column** into memory.
- `.xlsx` files are **XML-based and compressed**, not optimal for bulk read.
- Python lacks true `chunksize` support for Excel files.
- No native streaming from `.xlsx` unless using `openpyxl`.

---

## 💡 Optimization Strategies

| Strategy | Description |
|----------|-------------|
| **Use `.csv` over `.xlsx`** | CSV is faster and works with `chunksize`. |
| **Chunk processing** | Use `pd.read_csv(..., chunksize=100000)` to avoid memory spikes. |
| **Use `openpyxl` in read-only mode** | Iterate row-by-row in large `.xlsx` files. |
| **Filter early, write immediately** | Avoid buffering large filtered results in memory. |
| **Limit columns** | Load only the columns needed using `usecols`. |
| **Multiprocessing (advanced)** | Split work across multiple CPU cores for parallel file processing. |
| **VBA as fallback** | When Excel-native tasks are needed, VBA is sometimes faster and more reliable. |

---

## 🧠 Final Thoughts

This experience reminded me that:
- Python is powerful for automation—but **not always best** for native Excel tasks.
- For `.xlsx` heavy filtering or consolidation, **VBA is still the king**.
- Going forward, I’ll pre-process large Excel files using VBA or convert them to `.csv` before using Python.

---

# ✅ Lessons Learned: Excel File Consolidation (June 18, 2025)

## 📘 What We Learned

- **VBA is still faster than Python** when processing `.xlsx` files inside Excel, especially with filters and copy-paste operations across large sheets.
- **Python with pandas** can process Excel files, but:
  - It loads entire columns into memory.
  - It slows down significantly with high row counts (e.g., 500,000+ rows).
  - `.xlsx` is **not efficient** for large-scale automation.
- The `.py` script used was **well-structured** but **not memory-efficient**, because `pandas.read_excel()` does **not support chunking**.
- Performance bottlenecks depend heavily on:
  - **File format**: `.xlsx` is slower than `.csv`.
  - **Method used**: `read_excel` vs `openpyxl` vs VBA.
  - **Size of dataset**.

---

## 🛠️ Optimization Strategies (Future Proofing)

### 1. ❌ Avoid `.xlsx` for Bulk Processing
- ✅ Use `.csv` instead of `.xlsx` whenever possible.
- 🔍 **Why**: `.csv` loads faster and supports `chunksize` in pandas.

---

### 2. ✅ Use Chunking in Python
For `.csv` files:
```python
for chunk in pd.read_csv(file, chunksize=100000):
    # process filtered chunk
```
🔥 This avoids loading the entire file in memory at once.

---

### 3. 🧵 Use `openpyxl` for Row-by-Row Reading
Unlike `pandas`, `openpyxl` allows:
- Read-only mode
- Iterating row-by-row
- Applying conditions with low memory usage

**Example:**
```python
from openpyxl import load_workbook

wb = load_workbook(filename, read_only=True)
for sheet in wb.worksheets:
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[col_index] == "AGENT - CUSTOMER RPC":
            # append to CSV or list
```

---

### 4. 🎯 Extract Only Needed Columns
- ✅ Use `usecols` in `read_excel()` or `read_csv()` to limit memory usage.
```python
pd.read_excel(file, usecols=["DESCRIPTION"])
```

---

### 5. 💾 Write Filtered Output Directly (No Buffering)
- ✅ Use `to_csv(..., mode="a")` to **append output directly** and avoid storing results in memory.

---

### 6. 📊 Track Bottlenecks with Logging
- Use Python’s `time` module or `logging` to monitor time spent per file or step.

---

### 7. ⚙️ Use Multiprocessing (Advanced)
- Use `concurrent.futures` or `multiprocessing.Pool` to split file processing across CPU cores.

---

### 8. 📄 Pre-process in Excel Using VBA
- If `.xlsx` files are provided:
  - Use **VBA to pre-filter** and convert to `.csv`
  - Then use **Python** for downstream processing.

---

## 🧠 Final Thought

> Use **VBA** when working with Excel-heavy reports involving filters, formatting, or sheet-level logic.  
> Use **Python** for `.csv`, databases, or APIs — especially with chunking, multiprocessing, or low-memory pipelines.

---
