# Database Schema Documentation

## Overview

This document provides comprehensive documentation of the MS Access database schema, including tables, relationships, queries, and indexes. This information is essential for creating Laravel migrations and Eloquent models.

## How to Extract Schema Information

### Using Microsoft Access

1. **Open the Database** in Microsoft Access
2. **Database Tools** → **Relationships** to view relationships
3. **Table Design View** to see field properties
4. **Database Documenter** (Database Tools → Database Documenter) for comprehensive report

### Using mdbtools (Linux/Mac)

```bash
# List all tables
mdb-tables your_database.accdb

# Export schema
mdb-schema your_database.accdb mysql > schema.sql

# Get table structure
mdb-schema your_database.accdb mysql table_name
```

### Using ODBC/PHP Script

See the extraction scripts in `/scripts/extract-schema.php`

## Tables

[TO BE COMPLETED - Add tables from the actual database]

### Example Table Documentation Format

For each table in the database, document using this format:

---

#### Table: `tbl_employees`

**Purpose:** Stores employee information

**Fields:**

| Field Name | Data Type | Size | Required | Default | Description |
|------------|-----------|------|----------|---------|-------------|
| EmployeeID | AutoNumber | - | Yes (PK) | - | Unique employee identifier |
| FirstName | Text | 50 | Yes | - | Employee's first name |
| LastName | Text | 50 | Yes | - | Employee's last name |
| Email | Text | 100 | Yes | - | Employee email (unique) |
| Phone | Text | 20 | No | - | Contact phone number |
| HireDate | Date/Time | - | Yes | Date() | Date employee was hired |
| Salary | Currency | - | No | 0 | Annual salary |
| DepartmentID | Number (Long) | - | Yes (FK) | - | Reference to tblDepartments |
| ManagerID | Number (Long) | - | No | - | Reference to another employee |
| IsActive | Yes/No | - | Yes | True | Employment status |
| Notes | Memo | - | No | - | Additional notes |
| CreatedDate | Date/Time | - | Yes | Now() | Record creation timestamp |
| ModifiedDate | Date/Time | - | No | - | Last modification timestamp |

**Indexes:**

- Primary Key: EmployeeID
- Unique: Email
- Index: DepartmentID (for joins)
- Index: ManagerID (for joins)

**Relationships:**

- `tbl_employees.DepartmentID` → `tbl_departments.DepartmentID` (Many-to-One)
- `tbl_employees.ManagerID` → `tbl_employees.EmployeeID` (Self-referencing)

**Validation Rules:**

- Email must be unique
- HireDate cannot be in the future
- Salary must be >= 0

**Laravel Migration Equivalent:**

```php
Schema::create('employees', function (Blueprint $table) {
    $table->id();
    $table->string('first_name', 50);
    $table->string('last_name', 50);
    $table->string('email', 100)->unique();
    $table->string('phone', 20)->nullable();
    $table->date('hire_date');
    $table->decimal('salary', 10, 2)->nullable()->default(0);
    $table->foreignId('department_id')->constrained('departments');
    $table->foreignId('manager_id')->nullable()->constrained('employees');
    $table->boolean('is_active')->default(true);
    $table->text('notes')->nullable();
    $table->timestamps();
    $table->softDeletes();

    $table->index('department_id');
    $table->index('manager_id');
});
```

---

## Table Inventory

[TO BE COMPLETED - List all tables found in the database]

Example format:

| Table Name | Purpose | Record Count | Related Tables |
|------------|---------|--------------|----------------|
| tbl_employees | Employee records | ~500 | tbl_departments, tbl_positions |
| tbl_departments | Department information | ~20 | tbl_employees |
| tbl_projects | Project tracking | ~100 | tbl_employees, tbl_clients |
| tbl_tasks | Task assignments | ~1000 | tbl_projects, tbl_employees |
| ... | ... | ... | ... |

## Relationships Diagram

[TO BE COMPLETED - Add relationship diagram]

```
tbl_departments (1) ──< (Many) tbl_employees
                                    │
                                    │ (Many)
                                    v
                              tbl_tasks (Many) >── (1) tbl_projects
```

## Common Patterns

### Naming Conventions

| Pattern | Access | Laravel Recommendation |
|---------|--------|------------------------|
| Table names | tbl_employees | employees (plural) |
| ID fields | EmployeeID | id |
| Foreign keys | DepartmentID | department_id |
| Booleans | IsActive | is_active |
| Dates | CreatedDate | created_at |

### Lookup Tables

Document any lookup/reference tables:

**Example: `tbl_status`**

| StatusID | StatusName | Description |
|----------|------------|-------------|
| 1 | Active | Currently active |
| 2 | Inactive | Temporarily inactive |
| 3 | Terminated | Employment terminated |

**Laravel Enum Alternative:**
```php
enum EmployeeStatus: string
{
    case Active = 'active';
    case Inactive = 'inactive';
    case Terminated = 'terminated';
}
```

## Queries

[TO BE COMPLETED - Document saved queries]

### Example Query Documentation

**Query Name:** `qry_active_employees`

**SQL:**
```sql
SELECT e.EmployeeID, e.FirstName, e.LastName, e.Email,
       d.DepartmentName, e.HireDate
FROM tbl_employees e
INNER JOIN tbl_departments d ON e.DepartmentID = d.DepartmentID
WHERE e.IsActive = True
ORDER BY e.LastName, e.FirstName;
```

**Purpose:** List all active employees with department information

**Laravel Eloquent Equivalent:**
```php
Employee::with('department')
    ->where('is_active', true)
    ->orderBy('last_name')
    ->orderBy('first_name')
    ->get(['id', 'first_name', 'last_name', 'email', 'hire_date']);
```

**Used By:**
- frmEmployeeList (form)
- rptEmployeeDirectory (report)

---

## Queries Inventory

[TO BE COMPLETED - List all queries]

| Query Name | Type | Purpose | Complexity |
|------------|------|---------|------------|
| qry_active_employees | Select | List active employees | Simple |
| qry_employee_summary | Select | Employee statistics | Complex |
| qry_update_salaries | Action | Bulk salary update | Action |
| ... | ... | ... | ... |

## Indexes

Document critical indexes for performance:

[TO BE COMPLETED]

| Table | Index Name | Fields | Type |
|-------|------------|--------|------|
| tbl_employees | PK_EmployeeID | EmployeeID | Primary |
| tbl_employees | idx_email | Email | Unique |
| tbl_employees | idx_department | DepartmentID | Index |

## Data Integrity Rules

### Referential Integrity

[TO BE COMPLETED - Document cascade rules]

Example:
- When Department is deleted → Prevent if employees exist
- When Employee is deleted → Cascade delete related tasks
- When Project is deleted → Set employee assignments to NULL

### Validation Rules

[TO BE COMPLETED]

Document field-level validation:
- Email format validation
- Date range restrictions
- Numeric range limits
- Required field combinations

## Migration Notes

### Data Types Requiring Special Handling

1. **AutoNumber → id()**
   - Access AutoNumber starts at 1 and increments
   - Preserve IDs during migration if referenced externally

2. **Yes/No → boolean**
   - Access stores as -1 (True) and 0 (False)
   - Convert to 1/0 for MySQL/PostgreSQL

3. **Memo → text()**
   - Access Memo can store rich text
   - May need HTML cleanup

4. **Attachment → File Storage**
   - Extract files during migration
   - Store file paths in new system
   - Upload to Laravel storage

5. **Currency → decimal(10, 2)**
   - Currency stored with 4 decimal places in Access
   - Round to 2 decimals for standard currency

### Known Issues

[TO BE COMPLETED - Document any schema issues]

Example:
- Table X has no primary key (needs to be added)
- Circular references in table Y and Z
- Denormalized data in table W

## Migration Checklist

- [ ] Extract complete schema from Access
- [ ] Document all tables with fields and types
- [ ] Map all relationships
- [ ] Document all queries and their purposes
- [ ] Identify and document indexes
- [ ] Note validation rules and constraints
- [ ] Create Laravel migrations for all tables
- [ ] Test data type conversions
- [ ] Verify referential integrity
- [ ] Plan for data migration/seeding

## Tools and Scripts

Helpful scripts for schema extraction can be found in `/scripts/`:

- `extract-schema.php` - Extract schema via ODBC
- `extract-data.sh` - Export all data to CSV
- `generate-migrations.php` - Generate Laravel migrations

---

**Next Steps:**

1. Open the Access database
2. Use Database Documenter to generate a complete schema report
3. Fill in the [TO BE COMPLETED] sections above
4. Create Laravel migrations based on this documentation

**See Also:**
- [Laravel Migration Guide](./05-LARAVEL-MIGRATION.md)
- [VBA Modules Reference](./03-VBA-MODULES.md)
