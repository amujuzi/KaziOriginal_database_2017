# Database Extraction Scripts

## Overview

This directory contains scripts to extract schema, data, and VBA code from the MS Access database file.

## Database File

**File Name:** `Real - Operations Activity (07NOV2017) with campus visitors -master Copy.accdb`

**Location:** Root of repository

## Extraction Methods

### Method 1: Using Microsoft Access (Windows)

#### Extract Schema

1. Open the database in Microsoft Access
2. Go to **Database Tools** → **Database Documenter**
3. Select all tables, queries, forms, and modules
4. Click OK to generate report
5. Export to PDF or print to save

#### Export Tables to CSV

1. Select a table in the navigation pane
2. **External Data** → **Text File**
3. Export to `data/[TableName].csv`
4. Repeat for all tables

#### Export VBA Code

Run this VBA script in Access (Alt+F11 to open VBA Editor):

```vba
Sub ExportAllCode()
    Dim obj As AccessObject
    Dim db As Object
    Dim exportPath As String

    exportPath = CurrentProject.Path & "\vba-export\"

    ' Create export directory if it doesn't exist
    If Dir(exportPath, vbDirectory) = "" Then
        MkDir exportPath
    End If

    Set db = Application.CurrentProject

    ' Export standard modules
    For Each obj In db.AllModules
        On Error Resume Next
        Application.SaveAsText acModule, obj.Name, exportPath & obj.Name & ".bas"
    Next obj

    ' Export form modules
    For Each obj In db.AllForms
        On Error Resume Next
        Application.SaveAsText acForm, obj.Name, exportPath & "Form_" & obj.Name & ".txt"
    Next obj

    ' Export report modules
    For Each obj In db.AllReports
        On Error Resume Next
        Application.SaveAsText acReport, obj.Name, exportPath & "Report_" & obj.Name & ".txt"
    Next obj

    MsgBox "Export complete! Files saved to: " & exportPath
End Sub
```

### Method 2: Using mdbtools (Linux/Mac)

Install mdbtools:
```bash
# Ubuntu/Debian
sudo apt-get install mdbtools

# Mac
brew install mdbtools
```

Run the extraction script:
```bash
chmod +x scripts/extract-access-data.sh
./scripts/extract-access-data.sh "Real - Operations Activity (07NOV2017) with campus visitors -master Copy.accdb"
```

### Method 3: Using Python (Cross-platform)

Install required packages:
```bash
pip install pyodbc pandas
```

Run the Python extraction script:
```bash
python scripts/extract-access-data.py "Real - Operations Activity (07NOV2017) with campus visitors -master Copy.accdb"
```

## Output Structure

After extraction, you should have:

```
KaziOriginal_database_2017/
├── data/
│   ├── schema.sql              # Database schema
│   ├── tables/                 # CSV exports
│   │   ├── tblEmployees.csv
│   │   ├── tblDepartments.csv
│   │   └── ...
│   └── queries/                # Saved queries
│       ├── qryActiveEmployees.sql
│       └── ...
├── vba-export/                 # VBA code
│   ├── modUtilities.bas
│   ├── Form_frmEmployees.txt
│   └── ...
└── database-documentation.md   # Auto-generated docs
```

## Next Steps

After extraction:

1. Review the extracted data
2. Update `docs/02-DATABASE-SCHEMA.md` with actual table information
3. Update `docs/03-VBA-MODULES.md` with actual VBA code
4. Update `docs/04-FORMS-UI.md` with actual form information
5. Commit the extracted data and updated documentation
