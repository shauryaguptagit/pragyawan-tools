<div align="center">
  <h1>Pragyawan Tools</h1>
  <p>Internal automation scripts built during a data management internship — Excel processing and file distribution utilities.</p>
  <img src="https://img.shields.io/badge/Python-3776AB?style=for-the-badge&logo=python&logoColor=white"/>
  <img src="https://img.shields.io/badge/Pandas-150458?style=for-the-badge&logo=pandas&logoColor=white"/>
  <img src="https://img.shields.io/badge/Tkinter-FF6F00?style=for-the-badge&logoColor=white"/>
</div>

---

## Tools

### 1. Excel Data Automation

**Problem:** Manual data entry into Excel was consuming hours of team time each week.

**Solution:** Python + Pandas script that reads, cleans, and consolidates data from multiple sources into a single structured Excel file — automating the most repetitive parts of the pipeline.

```
auto-excel-updater/
├── main.py          # core processing script
└── MasterSheet.xlsx # template structure
```

**Stack:** Python, Pandas, openpyxl

---

### 2. USB File Distribution Utility

**Problem:** Staff manually copied specific files to network folders multiple times per day — tedious and error-prone.

**Solution:** Desktop GUI (Tkinter) with a single-click file distribution action. Select source, select destination, done.

```
usb-copier/
├── main.py    # application entry point
└── main.spec  # PyInstaller build spec (distributable .exe)
```

**Stack:** Python, Tkinter, shutil, PyInstaller

---

## Context

Built during an internship at Pragyawan Technologies Pvt. Ltd. (Jul–Sep 2025). Both tools were deployed and used by the operations team.