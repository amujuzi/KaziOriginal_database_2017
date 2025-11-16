# Database Extraction Template

**Database:** Real - Operations Activity (07NOV2017) with campus visitors -master Copy.accdb

Use this template to document the database as you explore it. Copy sections as needed.

## Quick Database Overview

**Purpose:** [What is this database used for?]

**Main Entities:** [e.g., Students, Visitors, Operations, Activities]

**Date Range of Data:** [What time period does the data cover?]

**Primary Users:** [Who uses this database?]

---

## Tables

### Table: [TableName]

**Purpose:** [What this table stores]

**Naming Convention:** [e.g., tblStudents, Operations, etc.]

| Field Name | Data Type | Size | Required | PK/FK | Default | Description |
|------------|-----------|------|----------|-------|---------|-------------|
| | | | | | | |

**Relationships:**
- [FieldName] → [OtherTable].[FieldName] (Relationship Type)

**Sample Data:**
[Paste or describe a few sample rows to understand the data]

**Business Rules:**
- [Any validation rules, constraints, or business logic]

**Data Volume:** [Approximate number of records]

---

## Queries

### Query: [QueryName]

**Purpose:** [What this query does]

**Type:** [Select/Update/Append/Delete/Crosstab]

**SQL:**
```sql
[Paste SQL here]
```

**Returns:** [What kind of data/how many rows]

**Used By:** [Which forms, reports, or other queries use this]

**Laravel Equivalent Approach:**
```php
[How you would do this in Laravel - can fill this in later]
```

---

## Forms

### Form: [FormName]

**Purpose:** [What this form is used for]

**Type:** [Single Form/Continuous Form/Datasheet/Dialog]

**Record Source:** [Table or Query name]

**Screenshot:** [Reference to screenshot file]

#### Properties
- Allow Additions: [Yes/No]
- Allow Edits: [Yes/No]
- Allow Deletions: [Yes/No]
- Data Entry: [Yes/No]
- Modal: [Yes/No]
- Pop Up: [Yes/No]

#### Controls

| Control Name | Type | Bound To | Properties | Purpose |
|--------------|------|----------|------------|---------|
| | | | | |

#### Subforms
- [Subform Name]: [Purpose and link fields]

#### VBA Events

**Form_Load:**
```vba
[Code or description]
```

**Form_Current:**
```vba
[Code or description]
```

**[Control]_Click:**
```vba
[Code or description]
```

#### Notes
[Any important observations about this form]

---

## Reports

### Report: [ReportName]

**Purpose:** [What this report shows]

**Record Source:** [Table or Query]

**Grouping:** [Any group levels]

**Sorting:** [Default sort order]

**Parameters:** [Any input parameters]

**Calculated Fields:**
- [FieldName]: [Calculation]

#### Notes
[Any important observations]

---

## VBA Modules

### Module: [ModuleName]

**Purpose:** [What this module does]

**Type:** [Standard Module/Class Module/Form Module]

#### Functions/Procedures

**Function/Sub: [Name]**

**Parameters:**
- [paramName] As [Type]: [Description]

**Returns:** [Return type and description]

**Purpose:** [What it does]

**Called By:** [Where it's used]

**VBA Code:**
```vba
[Paste code here]
```

**Laravel Equivalent:**
[Where this logic should go in Laravel - Service/Controller/Model/Helper]

```php
[Suggested implementation]
```

**Business Logic:**
[Describe any important business rules]

---

## Macros

### Macro: [MacroName]

**Purpose:** [What it does]

**Actions:**
1. [Action 1]
2. [Action 2]
...

**Triggered By:** [What triggers this macro]

**Laravel/Vue Equivalent:** [How to implement this in modern stack]

---

## Relationships

Document the relationships between tables:

```
[ParentTable]
    ├── [ChildTable1] (via FieldName)
    │   └── [GrandchildTable] (via FieldName)
    └── [ChildTable2] (via FieldName)
```

**Referential Integrity:**
- Cascade Update: [Yes/No - which relationships]
- Cascade Delete: [Yes/No - which relationships]

---

## Users & Security

**User-Level Security:** [Yes/No]

**Groups:**
- [Group Name]: [Permissions]

**Access Control:**
[How is access controlled in the application]

**Laravel Equivalent:**
- Roles needed: [List]
- Permissions needed: [List]

---

## Business Rules Discovered

List any business rules you discover:

1. **Rule:** [Description]
   - **Implementation:** [Where/how it's enforced]
   - **Laravel Approach:** [How to implement]

2. **Rule:** [Description]
   - **Implementation:** [Where/how it's enforced]
   - **Laravel Approach:** [How to implement]

---

## Data Flow

Describe the typical data flow/workflows:

### Workflow: [Name]

1. [Step 1]
2. [Step 2]
3. [Step 3]
...

**Forms Used:** [List]
**Tables Updated:** [List]
**Laravel Implementation:** [Suggested approach]

---

## Questions & Unknowns

Document things you're unsure about:

1. **Question:** [Your question]
   - **Context:** [Where you encountered this]
   - **Impact:** [Why this matters]

---

## Migration Priority

Based on what you've discovered, prioritize for migration:

### High Priority (Core Functionality)
1. [Feature/Table/Form]
2. [Feature/Table/Form]

### Medium Priority (Important but not critical)
1. [Feature/Table/Form]
2. [Feature/Table/Form]

### Low Priority (Nice to have)
1. [Feature/Table/Form]
2. [Feature/Table/Form]

### Can be Simplified/Modernized
1. [Feature/Table/Form]: [How it can be improved]

---

## Notes

Add any general notes or observations:

- [Observation 1]
- [Observation 2]
- [Insight 1]
- [Potential issue 1]

---

## Next Steps

After completing this documentation:

1. [ ] Update `docs/02-DATABASE-SCHEMA.md` with actual table info
2. [ ] Update `docs/03-VBA-MODULES.md` with actual VBA functions
3. [ ] Update `docs/04-FORMS-UI.md` with actual form layouts
4. [ ] Create Laravel migrations based on schema
5. [ ] Map VBA business logic to Laravel services
6. [ ] Design Vue components based on forms
7. [ ] Create data migration plan
