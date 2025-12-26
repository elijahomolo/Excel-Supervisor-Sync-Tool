# Excel-Supervisor-Sync-Tool

**Excel Supervisor Sync**

A Go CLI tool that synchronizes supervisor data from a master HR export into an Excel lookup sheet without modifying formula-driven reports.

Built for real-world Excel files, not idealized tables.

What It Does
	•	Reads a master workforce report (e.g. ALLOPS)
	•	Builds a mapping of:

        Colleague ID → Full Name + Supervisor

	•	Updates the vlookup sheet in a target workbook:
	•	Updates supervisor if the driver exists
	•	Appends driver name, ID, and supervisor if missing
	•	Leaves all report/formula sheets untouched

⸻

Why This Exists

Most Excel reports:
	•	Contain title rows above headers
	•	Use inconsistent header naming
	•	Rely on lookup sheets as the true data source

Editing report sheets directly breaks formulas.
This tool only updates safe, authoritative lookup tables.

⸻

**Features**
	•	Dynamic header row detection
	•	Fuzzy, case-insensitive header matching
	•	Name construction from Preferred First Name + Legal Last Name
	•	Dry-run mode
	•	Deterministic updates
	•	Unit-tested with real Excel files

**Usage**

`go build -o excel-sync

./excel-sync \
  --master ALLOPS_Workforce_Report.xlsx \
  --target shorts_report.xlsx \
  --vlookup-sheet vlookup \
  --out shorts_report_updated.xlsx`

*Dry Run*

`./excel-sync --master ALLOPS.xlsx --target report.xlsx --dry-run`

**Required Columns**

Master File
	•	Colleague ID
	•	Preferred First Name
	•	Legal Last Name
	•	Manager Name

Target (vlookup Sheet)
	•	Driver Name
	•	Driver Number
	•	Supervisor

Headers are matched fuzzily (case, punctuation, spacing ignored).

**Testing**

`go test ./...`

**Design Philosophy**
	•	Lookup tables are source-of-truth
	•	Reports are read-only artifacts
	•	Fail fast on structural mismatches
	•	Be resilient to Excel chaos

