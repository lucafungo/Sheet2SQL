# Sheet2SQL
# (An Excel to SQL Script Generator)

## 📄 Table of Contents
- [Introduction](#introduction)
- [Technical Overview](#technical-overview)
- [Dependencies & Installation](#dependencies--installation)
- [Directory and File Requirements](#directory-and-file-requirements)
- [Program Flow](#program-flow)
- [Detailed Code Walkthrough](#detailed-code-walkthrough)
- [Sample Input and Output](#sample-input-and-output)
- [Edge Cases & Handling](#edge-cases--handling)

---

## Introduction
This Python script reads an Excel spreadsheet and generates an SQL script to:
- Create a database table
- Populate it with data from the spreadsheet
- Optionally wrap the table as a temporary table

Supports SQL Server style formatting and escaping. Best used for ETL tasks, migrations, or data seeding.

---

## Technical Overview
- **Language:** Python 3.6+
- **Excel Parser:** `pandas.read_excel` with `openpyxl`
- **Output Format:** `.sql` file
- **SQL Constructs:** `CREATE TABLE`, `INSERT INTO`, and optional `SELECT`/`DROP`

---

## Dependencies & Installation
Install the following with pip:

```bash
pip install pandas openpyxl
```

---

## Directory and File Requirements
- Excel file must be `.xlsx`
- It should exist in the **same directory** as the script
- Column names should be unique
- Avoid merged cells and formulas
---

## Program Flow
1. User runs the script
2. Inputs Excel filename, table type, and table name
3. Data is parsed and headers are sanitized
4. SQL is generated and saved as `<filename>_output.sql`
5. Program loops until user types `exit`

---

## Detailed Code Walkthrough


### Core Logic: `generate_sql(input_excel, output_sql, table_name)`

#### Reading Excel
```python
df = pd.read_excel(input_excel, dtype=str, keep_default_na=False)
```
Reads all cells as strings. Empty cells become `''`.

#### Header Cleanup
```python
headers = [header.replace(" ", "_") for header in df.columns]
```
Replaces spaces in headers with underscores for SQL compatibility.

#### CREATE TABLE
```python
CREATE TABLE <table_name> (
   [Column1] varchar(255),
   ...
);
```
Hardcoded column types for simplicity.

#### INSERT INTO
```python
INSERT INTO <table_name> (
   [Col1], [Col2], ...
) VALUES
('val1', 'val2', ...),
...
```
Batches of 999 rows to stay within SQL Server limits.

#### Escaping
```python
str(cell).replace("'", "''").replace("\\", "/")
```
Escapes single quotes for SQL and avoids backslash issues.

#### Writing Output
```python
with open(output_file_path, 'w', encoding='utf-8') as sqlfile:
    sqlfile.write(sql_script)
```
Writes to `.sql` file in the same directory.

---

## Sample Input and Output

**Input Excel: `departments.xlsx`**

| DeptID | Dept Name   | Location   |
|--------|-------------|------------|
| 101    | HR          | New York   |
| 102    | Engineering | San Diego  |

**User Input:**
```
Enter the name of the Excel file: departments
Is this a temporary table? (yes/no): no
Enter the table name: departments
```

**Output SQL:**
```sql
CREATE TABLE departments (
    [DeptID] varchar(255),
    [Dept_Name] varchar(255),
    [Location] varchar(255)
);

-- SELECT * FROM departments;
-- DROP TABLE departments;

INSERT INTO departments (
   [DeptID],
   [Dept_Name],
   [Location]
) VALUES
('101', 'HR', 'New York'),
('102', 'Engineering', 'San Diego');
```

---

## Edge Cases & Handling

| Scenario            | Handling Description                          |
|---------------------|-----------------------------------------------|
| File not found      | Prompt and retry                              |
| Empty cells         | Converted to `''` in SQL                      |
| Special characters  | Single quotes escaped, backslashes replaced   |
| Duplicate headers   | Not handled yet (will crash)                  |
| Large files         | Inserted in batches of 999 rows               |

---


