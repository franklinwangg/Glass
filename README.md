# 🕒 Edge Browser History Extractor

A lightweight Python utility that exports your **Microsoft Edge browsing history** for a specific date into clean `.csv` and `.xlsx` tables.

---

## ⚙️ Setup

```bash
pip install pandas openpyxl
```

---

## 🚀 Usage

```bash
python script.py <username> <month> <day> <year>
```

### Example

```bash
python script.py frank 11 9 2025
```

This will:

- Copy `C:\Users\frank\AppData\Local\Microsoft\Edge\User Data\Default\History`
- Extract visits from **Nov 9, 2025**
- Output:

  ```
  output/11-9-25.csv
  output/11-9-25.xlsx
  ```

---

## 🧾 Output Format

| Title   | Visited (Local Time) |
| ------- | -------------------- |
| Google  | 11:19:10 PM          |
| Youtube | 11:19:09 PM          |
| Clickup | 10:34:42 PM          |

---

## 🧠 Notes

- Close Edge before running (it locks the History DB).
- Temporary copy of browsing history file is auto-deleted after export.
- `/output` is Git-ignored except for `.gitkeep`.

---

MIT License
