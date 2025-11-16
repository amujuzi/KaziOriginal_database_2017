# VBA Modules Reference

## Overview

This document catalogs all VBA code modules found in the MS Access database, including their purposes, functions, and Laravel/Vue equivalents.

## How to Extract VBA Code

### Using Microsoft Access

1. Open the database in Access
2. Press `Alt+F11` to open VBA Editor
3. View → Project Explorer
4. Export each module: Right-click → Export File

### Using Access VBA Export Script

```vba
' Run this in Access VBA to export all modules
Sub ExportAllModules()
    Dim obj As AccessObject
    Dim db As Object

    Set db = Application.CurrentProject

    ' Export standard modules
    For Each obj In db.AllModules
        Application.SaveAsText acModule, obj.Name, _
            "C:\export\" & obj.Name & ".bas"
    Next obj

    ' Export form modules
    For Each obj In db.AllForms
        If obj.IsLoaded Or HasModule(acForm, obj.Name) Then
            Application.SaveAsText acForm, obj.Name, _
                "C:\export\Form_" & obj.Name & ".bas"
        End If
    Next obj

    MsgBox "Export complete!"
End Sub

Function HasModule(objType As AcObjectType, objName As String) As Boolean
    On Error Resume Next
    HasModule = (Len(Application.GetOption(objType, objName)) > 0)
End Function
```

## Module Categories

[TO BE COMPLETED - Add actual modules from the database]

VBA modules typically fall into these categories:

1. **Standard Modules** - Reusable functions and procedures
2. **Form Modules** - Code behind forms (event handlers)
3. **Report Modules** - Code behind reports
4. **Class Modules** - Custom objects and classes

## Standard Modules

### Example Module Documentation Format

---

#### Module: `modUtilities`

**Purpose:** General utility functions used throughout the application

**Functions:**

##### `GetCurrentUserID() As Long`

**Purpose:** Returns the current user's ID

**VBA Code:**
```vba
Public Function GetCurrentUserID() As Long
    Dim rs As DAO.Recordset
    Set rs = CurrentDb.OpenRecordset("SELECT UserID FROM tblUsers WHERE Username = '" & CurrentUser() & "'")

    If Not rs.EOF Then
        GetCurrentUserID = rs!UserID
    Else
        GetCurrentUserID = 0
    End If

    rs.Close
End Function
```

**Laravel Equivalent:**
```php
// In a helper file or service
function getCurrentUserId(): int
{
    return auth()->id() ?? 0;
}

// Or directly use
auth()->id()
```

**Used By:**
- frmEmployees (Form_Load)
- frmProjects (btnAssign_Click)
- Multiple reports

---

##### `IsValidEmail(emailAddress As String) As Boolean`

**Purpose:** Validates email address format

**VBA Code:**
```vba
Public Function IsValidEmail(emailAddress As String) As Boolean
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")

    regex.Pattern = "^[\w\.-]+@[\w\.-]+\.\w+$"
    IsValidEmail = regex.Test(emailAddress)
End Function
```

**Laravel Equivalent:**
```php
// Use Laravel validation
$request->validate([
    'email' => 'required|email'
]);

// Or custom function
function isValidEmail(string $email): bool
{
    return filter_var($email, FILTER_VALIDATE_EMAIL) !== false;
}
```

**Vue/JavaScript Equivalent:**
```javascript
function isValidEmail(email) {
  const regex = /^[\w.-]+@[\w.-]+\.\w+$/
  return regex.test(email)
}

// Or use a validation library like Vuelidate
```

---

#### Module: `modDatabase`

**Purpose:** Database operations and data access functions

##### `ExecuteSQL(sqlStatement As String) As Boolean`

**VBA Code:**
```vba
Public Function ExecuteSQL(sqlStatement As String) As Boolean
    On Error GoTo ErrorHandler

    CurrentDb.Execute sqlStatement
    ExecuteSQL = True
    Exit Function

ErrorHandler:
    MsgBox "Database Error: " & Err.Description
    ExecuteSQL = False
End Function
```

**Laravel Equivalent:**
```php
use Illuminate\Support\Facades\DB;
use Illuminate\Support\Facades\Log;

function executeSql(string $sql): bool
{
    try {
        DB::statement($sql);
        return true;
    } catch (\Exception $e) {
        Log::error('Database Error: ' . $e->getMessage());
        return false;
    }
}

// Better approach: Use Eloquent or Query Builder instead of raw SQL
```

---

##### `GetRecordCount(tableName As String, whereClause As String) As Long`

**VBA Code:**
```vba
Public Function GetRecordCount(tableName As String, Optional whereClause As String = "") As Long
    Dim sql As String
    Dim rs As DAO.Recordset

    sql = "SELECT COUNT(*) AS Total FROM " & tableName
    If whereClause <> "" Then
        sql = sql & " WHERE " & whereClause
    End If

    Set rs = CurrentDb.OpenRecordset(sql)
    GetRecordCount = rs!Total
    rs.Close
End Function
```

**Laravel Equivalent:**
```php
use Illuminate\Support\Facades\DB;

function getRecordCount(string $table, ?string $whereColumn = null, $whereValue = null): int
{
    $query = DB::table($table);

    if ($whereColumn && $whereValue !== null) {
        $query->where($whereColumn, $whereValue);
    }

    return $query->count();
}

// Or using Eloquent:
// Employee::where('department_id', 5)->count()
```

---

#### Module: `modCalculations`

**Purpose:** Business logic calculations

##### `CalculateTotalCompensation(employeeID As Long) As Currency`

**VBA Code:**
```vba
Public Function CalculateTotalCompensation(employeeID As Long) As Currency
    Dim rs As DAO.Recordset
    Dim sql As String

    sql = "SELECT Salary, Bonus, Commission FROM tblEmployees WHERE EmployeeID = " & employeeID
    Set rs = CurrentDb.OpenRecordset(sql)

    If Not rs.EOF Then
        CalculateTotalCompensation = Nz(rs!Salary, 0) + Nz(rs!Bonus, 0) + Nz(rs!Commission, 0)
    Else
        CalculateTotalCompensation = 0
    End If

    rs.Close
End Function
```

**Laravel Equivalent:**
```php
// app/Services/EmployeeService.php
namespace App\Services;

use App\Models\Employee;

class EmployeeService
{
    public function calculateTotalCompensation(int $employeeId): float
    {
        $employee = Employee::find($employeeId);

        if (!$employee) {
            return 0;
        }

        return ($employee->salary ?? 0)
             + ($employee->bonus ?? 0)
             + ($employee->commission ?? 0);
    }
}

// Or as a model accessor in Employee model:
public function getTotalCompensationAttribute(): float
{
    return ($this->salary ?? 0)
         + ($this->bonus ?? 0)
         + ($this->commission ?? 0);
}

// Usage: $employee->total_compensation
```

---

## Form Modules

### Example Form Module

#### Form: `frmEmployees`

**Purpose:** Employee entry and editing form

**Key Event Handlers:**

##### `Form_Load()`

**VBA Code:**
```vba
Private Sub Form_Load()
    ' Load departments into combo box
    Me.cboDepartment.RowSource = "SELECT DepartmentID, DepartmentName FROM tblDepartments ORDER BY DepartmentName"

    ' Set default values for new records
    If Me.NewRecord Then
        Me.HireDate = Date
        Me.IsActive = True
        Me.CreatedBy = CurrentUser()
    End If

    ' Load user permissions
    If Not HasPermission("EditEmployees") Then
        Me.btnSave.Enabled = False
        Me.AllowEdits = False
    End If
End Sub
```

**Vue Component Equivalent:**
```vue
<script setup>
import { ref, onMounted, computed } from 'vue'
import { useRoute } from 'vue-router'
import { useAuthStore } from '@/stores/auth'

const route = useRoute()
const authStore = useAuthStore()

const employee = ref({
  hire_date: new Date().toISOString().split('T')[0],
  is_active: true,
  created_by: authStore.user.username
})

const departments = ref([])
const isNewRecord = computed(() => !route.params.id)
const canEdit = computed(() => authStore.hasPermission('edit-employees'))

onMounted(async () => {
  // Load departments
  const response = await fetch('/api/departments?sort=name')
  departments.value = await response.json()

  // Load employee if editing
  if (!isNewRecord.value) {
    const empResponse = await fetch(`/api/employees/${route.params.id}`)
    employee.value = await empResponse.json()
  }
})
</script>

<template>
  <form @submit.prevent="saveEmployee">
    <select v-model="employee.department_id" :disabled="!canEdit">
      <option v-for="dept in departments" :key="dept.id" :value="dept.id">
        {{ dept.name }}
      </option>
    </select>

    <button type="submit" :disabled="!canEdit">Save</button>
  </form>
</template>
```

---

##### `btnSave_Click()`

**VBA Code:**
```vba
Private Sub btnSave_Click()
    On Error GoTo ErrorHandler

    ' Validation
    If IsNull(Me.FirstName) Or IsNull(Me.LastName) Then
        MsgBox "First name and last name are required", vbCritical
        Exit Sub
    End If

    If Not IsValidEmail(Me.Email) Then
        MsgBox "Please enter a valid email address", vbCritical
        Me.Email.SetFocus
        Exit Sub
    End If

    ' Save record
    DoCmd.RunCommand acCmdSaveRecord

    ' Log activity
    Call LogActivity("Employee Save", "Saved employee: " & Me.FirstName & " " & Me.LastName)

    MsgBox "Employee saved successfully", vbInformation
    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrorHandler:
    MsgBox "Error saving employee: " & Err.Description, vbCritical
End Sub
```

**Laravel Controller Equivalent:**
```php
// app/Http/Controllers/EmployeeController.php
namespace App\Http\Controllers;

use App\Http\Requests\StoreEmployeeRequest;
use App\Models\Employee;
use App\Services\ActivityLogger;

class EmployeeController extends Controller
{
    public function store(StoreEmployeeRequest $request)
    {
        // Validation happens automatically in FormRequest

        $employee = Employee::create($request->validated());

        // Log activity
        ActivityLogger::log(
            'Employee Save',
            "Saved employee: {$employee->first_name} {$employee->last_name}"
        );

        return response()->json([
            'message' => 'Employee saved successfully',
            'employee' => $employee
        ], 201);
    }
}

// app/Http/Requests/StoreEmployeeRequest.php
class StoreEmployeeRequest extends FormRequest
{
    public function rules(): array
    {
        return [
            'first_name' => 'required|string|max:50',
            'last_name' => 'required|string|max:50',
            'email' => 'required|email|unique:employees',
            'department_id' => 'required|exists:departments,id',
        ];
    }
}
```

**Vue Component Equivalent:**
```vue
<script setup>
import { useRouter } from 'vue-router'
import { useToast } from '@/composables/useToast'

const router = useRouter()
const toast = useToast()

const saveEmployee = async () => {
  try {
    const response = await fetch('/api/employees', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(employee.value)
    })

    if (!response.ok) {
      const errors = await response.json()
      toast.error('Validation failed')
      return
    }

    toast.success('Employee saved successfully')
    router.push({ name: 'employees' })
  } catch (error) {
    toast.error('Error saving employee: ' + error.message)
  }
}
</script>
```

---

##### `cboDepartment_AfterUpdate()`

**VBA Code:**
```vba
Private Sub cboDepartment_AfterUpdate()
    ' When department changes, update manager dropdown
    If Not IsNull(Me.cboDepartment) Then
        Me.cboManager.RowSource = "SELECT EmployeeID, FirstName & ' ' & LastName AS FullName " & _
                                  "FROM tblEmployees " & _
                                  "WHERE DepartmentID = " & Me.cboDepartment & _
                                  " AND IsActive = True " & _
                                  "ORDER BY LastName"
        Me.cboManager.Requery
    Else
        Me.cboManager.RowSource = ""
    End If
End Sub
```

**Vue Component Equivalent:**
```vue
<script setup>
import { ref, watch } from 'vue'

const selectedDepartment = ref(null)
const managers = ref([])

// Watch for department changes
watch(selectedDepartment, async (newDepartmentId) => {
  if (newDepartmentId) {
    const response = await fetch(
      `/api/employees?department_id=${newDepartmentId}&is_active=true&role=manager`
    )
    managers.value = await response.json()
  } else {
    managers.value = []
  }
})
</script>

<template>
  <select v-model="selectedDepartment">
    <!-- Department options -->
  </select>

  <select v-model="employee.manager_id" :disabled="!selectedDepartment">
    <option v-for="manager in managers" :key="manager.id" :value="manager.id">
      {{ manager.full_name }}
    </option>
  </select>
</template>
```

---

## Report Modules

#### Report: `rptEmployeeSummary`

**Purpose:** Employee summary report with calculations

##### `Report_Load()`

**VBA Code:**
```vba
Private Sub Report_Load()
    Me.lblReportDate = "Generated: " & Format(Now(), "mm/dd/yyyy hh:nn AM/PM")
    Me.lblUserName = "By: " & CurrentUser()
End Sub
```

**Laravel/Vue Equivalent:**

Generate PDF reports using Laravel + DomPDF:

```php
// app/Http/Controllers/ReportController.php
use Barryvdh\DomPDF\Facade\Pdf;

public function employeeSummary()
{
    $employees = Employee::with('department')->get();

    $pdf = Pdf::loadView('reports.employee-summary', [
        'employees' => $employees,
        'generated_at' => now()->format('m/d/Y h:i A'),
        'generated_by' => auth()->user()->name
    ]);

    return $pdf->download('employee-summary.pdf');
}
```

```html
<!-- resources/views/reports/employee-summary.blade.php -->
<!DOCTYPE html>
<html>
<head>
    <title>Employee Summary Report</title>
</head>
<body>
    <h1>Employee Summary Report</h1>
    <p>Generated: {{ $generated_at }}</p>
    <p>By: {{ $generated_by }}</p>

    <table>
        <thead>
            <tr>
                <th>Name</th>
                <th>Department</th>
                <th>Hire Date</th>
            </tr>
        </thead>
        <tbody>
            @foreach($employees as $employee)
            <tr>
                <td>{{ $employee->full_name }}</td>
                <td>{{ $employee->department->name }}</td>
                <td>{{ $employee->hire_date->format('m/d/Y') }}</td>
            </tr>
            @endforeach
        </tbody>
    </table>
</body>
</html>
```

---

## Class Modules

#### Class: `clsEmployee`

**Purpose:** Employee business object with methods

**VBA Code:**
```vba
' Class module
Private mEmployeeID As Long
Private mFirstName As String
Private mLastName As String
Private mSalary As Currency

' Properties
Public Property Get EmployeeID() As Long
    EmployeeID = mEmployeeID
End Property

Public Property Let EmployeeID(value As Long)
    mEmployeeID = value
End Property

' Methods
Public Function GetFullName() As String
    GetFullName = mFirstName & " " & mLastName
End Function

Public Function ApplyRaise(percentage As Double) As Currency
    mSalary = mSalary * (1 + percentage / 100)
    ApplyRaise = mSalary
End Function
```

**Laravel Model Equivalent:**
```php
// app/Models/Employee.php
namespace App\Models;

use Illuminate\Database\Eloquent\Model;

class Employee extends Model
{
    protected $fillable = [
        'first_name',
        'last_name',
        'salary'
    ];

    protected $casts = [
        'salary' => 'decimal:2'
    ];

    // Accessor (getter)
    public function getFullNameAttribute(): string
    {
        return "{$this->first_name} {$this->last_name}";
    }

    // Method
    public function applyRaise(float $percentage): float
    {
        $this->salary = $this->salary * (1 + $percentage / 100);
        $this->save();
        return $this->salary;
    }
}

// Usage:
$employee = Employee::find(1);
echo $employee->full_name;  // Accessor
$newSalary = $employee->applyRaise(5);  // Method
```

---

## Module Inventory

[TO BE COMPLETED - List all modules found in the database]

| Module Name | Type | Purpose | Functions Count | Priority |
|-------------|------|---------|-----------------|----------|
| modUtilities | Standard | General utilities | 15 | High |
| modDatabase | Standard | Database operations | 8 | High |
| modCalculations | Standard | Business calculations | 12 | High |
| modValidation | Standard | Data validation | 6 | Medium |
| modReports | Standard | Report generation | 4 | Medium |
| frmEmployees | Form | Employee form events | 10 | High |
| ... | ... | ... | ... | ... |

## Migration Strategy

### Step 1: Identify Dependencies
- List all modules and their dependencies
- Identify which modules call other modules
- Map module relationships

### Step 2: Categorize by Purpose
- **Utilities** → Laravel Helpers or Services
- **Database Access** → Eloquent Models/Query Builder
- **Calculations** → Laravel Services
- **Validation** → Form Requests
- **Form Events** → Vue Component Methods

### Step 3: Prioritize Migration
1. Core utilities and database functions
2. Business logic calculations
3. Form event handlers
4. Report generation
5. Advanced features

### Step 4: Create Laravel Structure
```
app/
├── Services/           # Business logic from modules
│   ├── EmployeeService.php
│   ├── CalculationService.php
│   └── ValidationService.php
├── Helpers/           # Utility functions
│   └── helpers.php
├── Http/
│   ├── Controllers/   # Form actions
│   └── Requests/      # Validation
└── Models/            # Database access
```

## Testing Converted Code

Create tests for migrated VBA functions:

```php
// tests/Feature/EmployeeServiceTest.php
namespace Tests\Feature;

use Tests\TestCase;
use App\Services\EmployeeService;
use App\Models\Employee;

class EmployeeServiceTest extends TestCase
{
    public function test_calculate_total_compensation()
    {
        $employee = Employee::factory()->create([
            'salary' => 50000,
            'bonus' => 5000,
            'commission' => 2000
        ]);

        $service = new EmployeeService();
        $total = $service->calculateTotalCompensation($employee->id);

        $this->assertEquals(57000, $total);
    }
}
```

## Migration Checklist

- [ ] Export all VBA modules from Access
- [ ] Catalog all modules and functions
- [ ] Identify module dependencies
- [ ] Create Laravel services for business logic
- [ ] Create helper functions for utilities
- [ ] Convert form events to Vue components
- [ ] Migrate validation to Form Requests
- [ ] Test all migrated functions
- [ ] Document API endpoints
- [ ] Update this documentation with actual modules

---

**See Also:**
- [VBA Fundamentals](./01-VBA-FUNDAMENTALS.md)
- [Business Logic Mapping](./08-BUSINESS-LOGIC.md)
- [Translation Dictionary](./07-TRANSLATION-DICTIONARY.md)
