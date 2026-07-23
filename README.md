# Repeatability Analysis Tool

A desktop utility I built in Python to check measurement repeatability across multiple test runs and flag values that drift beyond a set tolerance. It reads several CSV exports, compares readings point by point, and produces a color-coded Excel report so an engineer can see at a glance where a machine is stable and where it is not.

## What it does
- Loads two or more CSV files through a simple Tkinter interface.
- Compares readings across runs against a user-defined tolerance.
- Highlights out-of-tolerance points in the generated Excel report.
- Saves a timestamped report and opens it automatically.

## Tech
Python, pandas, NumPy, openpyxl, Tkinter.

## Context
Built to speed up repeatability studies on precision alignment equipment. Some prompts in the interface are in Chinese because it was used in a Taiwan production environment.

## Files
- `project92.py`: current version.
- `project92_v1_old.py`: earlier version, kept for reference.
