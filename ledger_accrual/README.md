# Ledger Auto-Accrual Processor

A small utility that reads Excel ledger files, computes interest on the running
balance, and writes formatted, print-ready Excel output with auto-generated
interest entries.

## What it does

- Scans the current directory for `.xlsx` / `.xls` ledger files
- Reads interest frequency and rate ranges from a `rate.txt` file next to the ledger
- Accrues interest daily (`balance × rate ÷ 365`) and posts it at each period boundary
- Supports rate changes mid-period, splitting the interest label into segments

Balance logic: 
`Balance = Balance + Credit − Debit`.


## Requirements

- Python 3
- [`openpyxl`](https://pypi.org/project/openpyxl/)

```
pip install openpyxl
```

## Usage

Run from the directory containing your ledger file(s) and `rate.txt`:

```
python ledger_accrual.py
```

Output workbooks are written to `output/`, named like
`<ledger>_<freq>_<rates>pct.xlsx`.
