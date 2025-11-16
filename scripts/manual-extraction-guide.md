# Manual Extraction Guide for MS Access Database

This guide helps you manually extract information from the Access database when automated tools are not available.

## Database File

**File Name:** `Real - Operations Activity (07NOV2017) with campus visitors -master Copy.accdb`

## Step-by-Step Manual Extraction

### 1. Document Tables

1. Open the database in Microsoft Access
2. In the Navigation Pane, note all tables (usually start with `tbl`)
3. For each table:
   - Right-click → **Design View**
   - Document each field:
     - Field Name
     - Data Type
     - Field Size
     - Required (Yes/No)
     - Default Value
     - Description
   - Note the Primary Key (key icon)
   - **Database Tools** → **Relationships** to see foreign keys

**Template for documenting a table:**

```markdown
### Table: tbl[TableName]

**Purpose:** [What this table stores]

| Field Name | Data Type | Size | Required | Default | Description |
|------------|-----------|------|----------|---------|-------------|
| ID | AutoNumber | - | Yes (PK) | - | Primary key |
| FieldName | Text | 50 | Yes | - | Description |
| ... | ... | ... | ... | ... | ... |

**Relationships:**
- Field → OtherTable.Field (One-to-Many)
```

### 2. Document Queries

1. In Navigation Pane, click **Queries**
2. For each query:
   - Right-click → **SQL View**
   - Copy the SQL statement
   - Note what it's used for

**Template:**

```markdown
### Query: qry[QueryName]

**Purpose:** [What this query does]

**SQL:**
```sql
SELECT ...
FROM ...
WHERE ...
```

**Used By:** [Which forms/reports use it]
```

### 3. Document Forms

1. In Navigation Pane, click **Forms**
2. For each form:
   - Open in **Design View**
   - Note **Record Source** (from Property Sheet)
   - List all controls (textboxes, combos, buttons, etc.)
   - Right-click form → **View Code** to see VBA
   - Take a screenshot of the form layout

**Template:**

```markdown
### Form: frm[FormName]

**Purpose:** [What this form does]

**Record Source:** [Table or Query]

**Screenshot:** ![Form Layout](../screenshots/frm[FormName].png)

**Controls:**

| Control Name | Type | Bound To | Purpose |
|--------------|------|----------|---------|
| txt[Name] | TextBox | FieldName | Input/display field |
| cbo[Name] | ComboBox | FieldName | Dropdown selection |
| btn[Name] | Button | - | Performs action |

**VBA Code:** See `vba-export/Form_frm[FormName].txt`

**Events:**
- Form_Load: [Description]
- btn[Name]_Click: [Description]
```

### 4. Document Reports

1. In Navigation Pane, click **Reports**
2. For each report:
   - Open in **Design View**
   - Note **Record Source**
   - Document grouping levels
   - Note any calculated fields
   - View the VBA code

### 5. Document VBA Modules

1. Press **Alt+F11** to open VBA Editor
2. In Project Explorer, you'll see:
   - Microsoft Access Class Objects (forms, reports)
   - Modules (standard code modules)

3. For each module:
   - Double-click to view code
   - Document each Public Function/Sub
   - Copy the code

**Template:**

```markdown
### Module: mod[ModuleName]

**Purpose:** [What this module does]

#### Function: [FunctionName]

**Purpose:** [What this function does]

**VBA Code:**
```vba
Public Function [FunctionName](param As Type) As ReturnType
    ' Code here
End Function
```

**Laravel Equivalent:**
```php
public function functionName($param): ReturnType
{
    // Code here
}
```

**Used By:** [Where this is called from]
```

### 6. Export VBA Code

To export all VBA code at once, create a new module in the VBA Editor and paste this code:

```vba
Sub ExportAllVBACode()
    Dim obj As AccessObject
    Dim db As Object
    Dim exportPath As String
    Dim fso As Object
    Dim textStream As Object

    ' Set export path
    exportPath = CurrentProject.Path & "\vba-export\"

    ' Create FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Create directory if it doesn't exist
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder exportPath
    End If

    Set db = Application.CurrentProject

    ' Export standard modules
    Debug.Print "Exporting standard modules..."
    For Each obj In db.AllModules
        On Error Resume Next
        Application.SaveAsText acModule, obj.Name, exportPath & obj.Name & ".bas"
        Debug.Print "  - " & obj.Name & ".bas"
    Next obj

    ' Export form modules
    Debug.Print "Exporting form modules..."
    For Each obj In db.AllForms
        On Error Resume Next
        Application.SaveAsText acForm, obj.Name, exportPath & "Form_" & obj.Name & ".txt"
        Debug.Print "  - Form_" & obj.Name & ".txt"
    Next obj

    ' Export report modules
    Debug.Print "Exporting report modules..."
    For Each obj In db.AllReports
        On Error Resume Next
        Application.SaveAsText acReport, obj.Name, exportPath & "Report_" & obj.Name & ".txt"
        Debug.Print "  - Report_" & obj.Name & ".txt"
    Next obj

    MsgBox "VBA export complete! Files saved to: " & vbCrLf & exportPath, vbInformation
End Sub
```

**To run it:**
1. Paste into a new module
2. Press F5 or click Run
3. Files will be exported to `vba-export\` folder

### 7. Export Data

To export all tables to CSV:

```vba
Sub ExportAllTablesToCSV()
    Dim db As DAO.Database
    Dim tbl As DAO.TableDef
    Dim exportPath As String
    Dim fso As Object

    Set db = CurrentDb
    Set fso = CreateObject("Scripting.FileSystemObject")

    exportPath = CurrentProject.Path & "\data\tables\"

    ' Create directory
    If Not fso.FolderExists(CurrentProject.Path & "\data") Then
        fso.CreateFolder CurrentProject.Path & "\data"
    End If
    If Not fso.FolderExists(exportPath) Then
        fso.CreateFolder exportPath
    End If

    Debug.Print "Exporting tables to CSV..."

    For Each tbl In db.TableDefs
        ' Skip system tables
        If Left(tbl.Name, 4) <> "MSys" And Left(tbl.Name, 1) <> "~" Then
            On Error Resume Next
            DoCmd.TransferText acExportDelim, , tbl.Name, exportPath & tbl.Name & ".csv", True
            Debug.Print "  - " & tbl.Name & ".csv (" & DCount("*", tbl.Name) & " rows)"
        End If
    Next tbl

    MsgBox "CSV export complete! Files saved to: " & vbCrLf & exportPath, vbInformation
End Sub
```

### 8. Generate Database Documentation

Use Access's built-in Database Documenter:

1. **Database Tools** → **Database Documenter**
2. Select **All Object Types**
3. Click **Options** and select what to include:
   - Properties
   - Relationships
   - Permissions
4. Click **OK**
5. The report will open in Print Preview
6. **External Data** → **PDF or XPS** to save as PDF

### 9. Create Entity Relationship Diagram

1. **Database Tools** → **Relationships**
2. If relationships aren't showing:
   - Right-click → **Show Table**
   - Add all tables
3. Access will show relationships as lines between tables
4. Take a screenshot or:
   - **Design** → **Relationship Report**
   - Save as PDF

## Checklist

Use this checklist to ensure complete extraction:

### Tables
- [ ] List of all tables documented
- [ ] Each table's fields and properties documented
- [ ] Primary keys identified
- [ ] Foreign keys identified
- [ ] Relationships mapped
- [ ] Data exported to CSV

### Queries
- [ ] All queries listed
- [ ] SQL for each query extracted
- [ ] Purpose of each query documented

### Forms
- [ ] All forms listed
- [ ] Each form's controls documented
- [ ] VBA code extracted
- [ ] Screenshots taken
- [ ] Record sources noted

### Reports
- [ ] All reports listed
- [ ] Report layouts documented
- [ ] VBA code extracted

### VBA Code
- [ ] All modules exported
- [ ] Form code modules exported
- [ ] Report code modules exported
- [ ] Functions documented with purpose

### Other
- [ ] Macros documented (if any)
- [ ] Database Documenter report generated
- [ ] Relationships diagram created
- [ ] Overall database purpose documented

## Next Steps

After completing the extraction:

1. **Update documentation files:**
   - `docs/02-DATABASE-SCHEMA.md` - Add table information
   - `docs/03-VBA-MODULES.md` - Add VBA code and functions
   - `docs/04-FORMS-UI.md` - Add form layouts and controls

2. **Organize extracted files:**
   ```
   KaziOriginal_database_2017/
   ├── data/
   │   ├── tables/              # CSV exports
   │   ├── database-doc.pdf     # Database Documenter output
   │   └── relationships.pdf    # ER diagram
   ├── vba-export/              # VBA code
   │   ├── modUtilities.bas
   │   ├── Form_frmEmployees.txt
   │   └── ...
   └── screenshots/             # Form/report screenshots
       ├── frmEmployees.png
       └── ...
   ```

3. **Commit to repository:**
   ```bash
   git add data/ vba-export/ screenshots/
   git commit -m "Add extracted Access database data and code"
   git push
   ```

## Tips

- **Take your time** - Thorough documentation now saves time later
- **Test queries** - Run each query to see sample results
- **Document assumptions** - Note any business rules you discover
- **Ask questions** - If you're unsure about something, document it as a question
- **Keep original** - Always keep a backup of the original .accdb file

## Getting Help

If you encounter issues:

1. Check Access version (File → Account → About Access)
2. Ensure you have permission to open the database
3. Try opening in Safe Mode: Hold Shift while opening
4. Check for password protection
5. Consult the Access documentation at microsoft.com/access
