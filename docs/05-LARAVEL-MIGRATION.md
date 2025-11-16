# Laravel Migration Guide

## Overview

This guide provides step-by-step instructions for migrating the MS Access database and VBA business logic to Laravel.

## Table of Contents

1. [Database Migration Strategy](#database-migration-strategy)
2. [Creating Eloquent Models](#creating-eloquent-models)
3. [Migrating Business Logic](#migrating-business-logic)
4. [Authentication & Authorization](#authentication--authorization)
5. [API Design](#api-design)
6. [File Uploads & Storage](#file-uploads--storage)
7. [Reports & Exports](#reports--exports)

## Database Migration Strategy

### Step 1: Extract Schema from Access

Use one of these methods to extract the schema:

**Method 1: Using mdbtools (Linux)**
```bash
# Install mdbtools
sudo apt-get install mdbtools

# List tables
mdb-tables your_database.accdb

# Export schema
mdb-schema your_database.accdb mysql > schema.sql

# Export data
for table in $(mdb-tables your_database.accdb); do
    mdb-export your_database.accdb "$table" > "${table}.csv"
done
```

**Method 2: Using Microsoft Access**
1. Open the database in Access
2. External Data → Export → Text File
3. Repeat for each table

**Method 3: Using ODBC Connection (PHP Script)**
```php
<?php
// On Windows with Access Database Engine installed
$conn = odbc_connect("Driver={Microsoft Access Driver (*.mdb, *.accdb)};Dbq=C:\path\to\database.accdb", '', '');

$tables = odbc_tables($conn);
while (odbc_fetch_row($tables)) {
    $tableName = odbc_result($tables, "TABLE_NAME");
    echo "Table: $tableName\n";
}
```

### Step 2: Create Laravel Migrations

Convert Access tables to Laravel migrations:

**Access Table Example:**
```
Table: tbl_users
- UserID (AutoNumber, Primary Key)
- Username (Text, 50)
- Email (Text, 100)
- Password (Text, 255)
- IsActive (Yes/No)
- CreatedDate (Date/Time)
- DepartmentID (Number, Foreign Key)
```

**Laravel Migration:**
```php
<?php
// database/migrations/2024_01_01_000001_create_users_table.php

use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;

return new class extends Migration
{
    public function up(): void
    {
        Schema::create('users', function (Blueprint $table) {
            $table->id(); // Replaces AutoNumber
            $table->string('username', 50);
            $table->string('email', 100)->unique();
            $table->string('password', 255);
            $table->boolean('is_active')->default(true);
            $table->foreignId('department_id')
                  ->constrained('departments')
                  ->onDelete('cascade');
            $table->timestamps(); // created_at, updated_at
            $table->softDeletes(); // deleted_at (optional)
        });
    }

    public function down(): void
    {
        Schema::dropIfExists('users');
    }
};
```

### Data Type Mapping

| Access Data Type | Laravel Migration Method | Notes |
|------------------|-------------------------|-------|
| AutoNumber | `$table->id()` | Unsigned BIGINT, auto-increment |
| Text (n) | `$table->string('field', n)` | VARCHAR(n) |
| Memo / Long Text | `$table->text('field')` | TEXT |
| Number (Integer) | `$table->integer('field')` | INT |
| Number (Long) | `$table->bigInteger('field')` | BIGINT |
| Number (Decimal) | `$table->decimal('field', 10, 2)` | DECIMAL |
| Currency | `$table->decimal('field', 10, 2)` | Store in cents or use decimal |
| Yes/No | `$table->boolean('field')` | TINYINT(1) |
| Date/Time | `$table->dateTime('field')` | DATETIME |
| Date | `$table->date('field')` | DATE |
| Time | `$table->time('field')` | TIME |
| Attachment | `$table->string('field')` | Store file path |
| OLE Object | `$table->binary('field')` | BLOB (avoid if possible) |
| Hyperlink | `$table->string('field')` | VARCHAR |
| Lookup | `$table->foreignId('field')` | Use relationships |

### Step 3: Data Migration

Create a seeder or migration script to import data:

```php
<?php
// database/seeders/ImportAccessDataSeeder.php

namespace Database\Seeders;

use Illuminate\Database\Seeder;
use Illuminate\Support\Facades\DB;
use League\Csv\Reader;

class ImportAccessDataSeeder extends Seeder
{
    public function run(): void
    {
        $this->importUsers();
        $this->importDepartments();
        // ... other tables
    }

    private function importUsers(): void
    {
        $csv = Reader::createFromPath(database_path('imports/users.csv'), 'r');
        $csv->setHeaderOffset(0);

        DB::transaction(function () use ($csv) {
            foreach ($csv as $record) {
                DB::table('users')->insert([
                    'id' => $record['UserID'],
                    'username' => $record['Username'],
                    'email' => $record['Email'],
                    'password' => bcrypt($record['Password']), // Hash if plain text
                    'is_active' => $record['IsActive'] === 'True',
                    'department_id' => $record['DepartmentID'],
                    'created_at' => $record['CreatedDate'],
                    'updated_at' => $record['CreatedDate'],
                ]);
            }
        });
    }
}
```

## Creating Eloquent Models

### Basic Model Structure

**Access Table: tbl_employees**

**Laravel Model:**
```php
<?php
// app/Models/Employee.php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
use Illuminate\Database\Eloquent\Relations\BelongsTo;
use Illuminate\Database\Eloquent\Relations\HasMany;
use Illuminate\Database\Eloquent\SoftDeletes;

class Employee extends Model
{
    use HasFactory, SoftDeletes;

    protected $fillable = [
        'first_name',
        'last_name',
        'email',
        'phone',
        'hire_date',
        'salary',
        'department_id',
        'is_active',
    ];

    protected $casts = [
        'hire_date' => 'date',
        'salary' => 'decimal:2',
        'is_active' => 'boolean',
    ];

    // Relationships
    public function department(): BelongsTo
    {
        return $this->belongsTo(Department::class);
    }

    public function tasks(): HasMany
    {
        return $this->hasMany(Task::class);
    }

    // Accessors (like calculated fields in Access)
    public function getFullNameAttribute(): string
    {
        return "{$this->first_name} {$this->last_name}";
    }

    // Scopes (like saved queries in Access)
    public function scopeActive($query)
    {
        return $query->where('is_active', true);
    }

    public function scopeInDepartment($query, $departmentId)
    {
        return $query->where('department_id', $departmentId);
    }
}
```

### Relationship Mapping

**Access Relationships → Laravel Eloquent**

**One-to-Many (Department → Employees):**
```php
// Department Model
public function employees(): HasMany
{
    return $this->hasMany(Employee::class);
}

// Employee Model
public function department(): BelongsTo
{
    return $this->belongsTo(Department::class);
}

// Usage
$department = Department::find(1);
$employees = $department->employees; // Get all employees

$employee = Employee::find(1);
$department = $employee->department; // Get department
```

**Many-to-Many (Employees ↔ Projects):**
```php
// Employee Model
public function projects(): BelongsToMany
{
    return $this->belongsToMany(Project::class, 'employee_project')
                ->withPivot('role', 'hours')
                ->withTimestamps();
}

// Project Model
public function employees(): BelongsToMany
{
    return $this->belongsToMany(Employee::class, 'employee_project')
                ->withPivot('role', 'hours')
                ->withTimestamps();
}

// Usage
$employee = Employee::find(1);
$projects = $employee->projects; // Get all projects

// Attach relationship
$employee->projects()->attach($projectId, ['role' => 'Developer', 'hours' => 40]);
```

## Migrating Business Logic

### VBA Procedures → Laravel Services

**VBA Code:**
```vba
Public Function CalculateEmployeeSalary(employeeID As Long) As Currency
    Dim rs As DAO.Recordset
    Dim baseSalary As Currency
    Dim bonus As Currency

    Set rs = CurrentDb.OpenRecordset("SELECT BaseSalary, Bonus FROM tblEmployees WHERE EmployeeID = " & employeeID)

    If Not rs.EOF Then
        baseSalary = rs!BaseSalary
        bonus = rs!Bonus
        CalculateEmployeeSalary = baseSalary + bonus
    Else
        CalculateEmployeeSalary = 0
    End If

    rs.Close
End Function
```

**Laravel Service:**
```php
<?php
// app/Services/EmployeeService.php

namespace App\Services;

use App\Models\Employee;

class EmployeeService
{
    public function calculateEmployeeSalary(int $employeeId): float
    {
        $employee = Employee::find($employeeId);

        if (!$employee) {
            return 0;
        }

        return $employee->base_salary + $employee->bonus;
    }

    // Or using accessor on the model
    // In Employee model:
    // public function getTotalSalaryAttribute(): float
    // {
    //     return $this->base_salary + $this->bonus;
    // }
    //
    // Usage: $employee->total_salary
}
```

### VBA Form Actions → Laravel Controllers

**VBA Button Click:**
```vba
Private Sub btnSave_Click()
    On Error GoTo ErrorHandler

    If IsNull(Me.txtFirstName) Or IsNull(Me.txtLastName) Then
        MsgBox "First name and last name are required", vbCritical
        Exit Sub
    End If

    DoCmd.RunSQL "INSERT INTO tblEmployees (FirstName, LastName, Email) VALUES ('" & _
                 Me.txtFirstName & "', '" & Me.txtLastName & "', '" & Me.txtEmail & "')"

    MsgBox "Employee saved successfully", vbInformation
    DoCmd.Close acForm, Me.Name
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub
```

**Laravel Controller + Form Request:**
```php
<?php
// app/Http/Requests/StoreEmployeeRequest.php

namespace App\Http\Requests;

use Illuminate\Foundation\Http\FormRequest;

class StoreEmployeeRequest extends FormRequest
{
    public function rules(): array
    {
        return [
            'first_name' => 'required|string|max:50',
            'last_name' => 'required|string|max:50',
            'email' => 'required|email|unique:employees',
        ];
    }
}

// app/Http/Controllers/EmployeeController.php

namespace App\Http\Controllers;

use App\Http\Requests\StoreEmployeeRequest;
use App\Models\Employee;

class EmployeeController extends Controller
{
    public function store(StoreEmployeeRequest $request)
    {
        $employee = Employee::create($request->validated());

        return response()->json([
            'message' => 'Employee saved successfully',
            'employee' => $employee
        ], 201);
    }
}
```

### Query Migration

**Access Query → Eloquent**

**Access Saved Query:**
```sql
SELECT e.EmployeeID, e.FirstName, e.LastName, d.DepartmentName,
       e.Salary, e.Bonus, (e.Salary + e.Bonus) AS TotalCompensation
FROM tblEmployees e
INNER JOIN tblDepartments d ON e.DepartmentID = d.DepartmentID
WHERE e.IsActive = True AND d.DepartmentName = 'IT'
ORDER BY TotalCompensation DESC
```

**Laravel Eloquent:**
```php
$employees = Employee::with('department')
    ->select('employees.*')
    ->selectRaw('(salary + bonus) as total_compensation')
    ->join('departments', 'employees.department_id', '=', 'departments.id')
    ->where('employees.is_active', true)
    ->where('departments.name', 'IT')
    ->orderByDesc('total_compensation')
    ->get();

// Or using relationships
$employees = Employee::active()
    ->whereHas('department', function ($query) {
        $query->where('name', 'IT');
    })
    ->with('department')
    ->get()
    ->map(function ($employee) {
        $employee->total_compensation = $employee->salary + $employee->bonus;
        return $employee;
    })
    ->sortByDesc('total_compensation');
```

## Authentication & Authorization

### Replace Access User/Group Security

**Laravel Breeze/Sanctum Setup:**
```bash
composer require laravel/sanctum
php artisan vendor:publish --provider="Laravel\Sanctum\SanctumServiceProvider"
php artisan migrate
```

**User Model:**
```php
<?php
namespace App\Models;

use Illuminate\Foundation\Auth\User as Authenticatable;
use Laravel\Sanctum\HasApiTokens;

class User extends Authenticatable
{
    use HasApiTokens;

    protected $fillable = [
        'name',
        'email',
        'password',
    ];

    protected $hidden = [
        'password',
        'remember_token',
    ];
}
```

**Authorization (Roles & Permissions):**
```bash
composer require spatie/laravel-permission
```

```php
<?php
// Assign roles
$user->assignRole('admin');

// Check permissions
if ($user->can('edit-employees')) {
    // Allow action
}

// In controller
public function update(Request $request, Employee $employee)
{
    $this->authorize('update', $employee);
    // Update logic
}

// Policy
class EmployeePolicy
{
    public function update(User $user, Employee $employee)
    {
        return $user->hasRole('admin') ||
               $user->department_id === $employee->department_id;
    }
}
```

## API Design

### RESTful API Structure

**Routes:**
```php
<?php
// routes/api.php

use App\Http\Controllers\EmployeeController;

Route::middleware('auth:sanctum')->group(function () {
    Route::apiResource('employees', EmployeeController::class);
    Route::get('employees/{employee}/tasks', [EmployeeController::class, 'tasks']);
    Route::post('employees/{employee}/assign-project', [EmployeeController::class, 'assignProject']);
});
```

**API Resource (Data Transformation):**
```php
<?php
// app/Http/Resources/EmployeeResource.php

namespace App\Http\Resources;

use Illuminate\Http\Resources\Json\JsonResource;

class EmployeeResource extends JsonResource
{
    public function toArray($request)
    {
        return [
            'id' => $this->id,
            'full_name' => $this->full_name,
            'email' => $this->email,
            'department' => new DepartmentResource($this->whenLoaded('department')),
            'hire_date' => $this->hire_date->format('Y-m-d'),
            'is_active' => $this->is_active,
            'created_at' => $this->created_at->toISOString(),
        ];
    }
}
```

## File Uploads & Storage

**Replace Access Attachments:**

```php
<?php
// Controller
public function uploadDocument(Request $request, Employee $employee)
{
    $request->validate([
        'document' => 'required|file|mimes:pdf,doc,docx|max:10240'
    ]);

    $path = $request->file('document')->store('employee-documents', 'public');

    $employee->documents()->create([
        'filename' => $request->file('document')->getClientOriginalName(),
        'path' => $path,
        'size' => $request->file('document')->getSize(),
    ]);

    return response()->json(['path' => $path]);
}

// config/filesystems.php - configure storage
```

## Reports & Exports

**Replace Access Reports:**

```php
<?php
// Using Laravel Excel
composer require maatwebsite/excel

// app/Exports/EmployeesExport.php
namespace App\Exports;

use App\Models\Employee;
use Maatwebsite\Excel\Concerns\FromCollection;

class EmployeesExport implements FromCollection
{
    public function collection()
    {
        return Employee::with('department')->get();
    }
}

// Controller
use App\Exports\EmployeesExport;
use Maatwebsite\Excel\Facades\Excel;

public function export()
{
    return Excel::download(new EmployeesExport, 'employees.xlsx');
}

// For PDF reports
composer require barryvdh/laravel-dompdf
```

## Testing

**Write tests for business logic:**

```php
<?php
// tests/Feature/EmployeeTest.php

namespace Tests\Feature;

use Tests\TestCase;
use App\Models\Employee;
use Illuminate\Foundation\Testing\RefreshDatabase;

class EmployeeTest extends TestCase
{
    use RefreshDatabase;

    public function test_can_create_employee()
    {
        $response = $this->postJson('/api/employees', [
            'first_name' => 'John',
            'last_name' => 'Doe',
            'email' => 'john@example.com',
        ]);

        $response->assertStatus(201);
        $this->assertDatabaseHas('employees', ['email' => 'john@example.com']);
    }
}
```

## Migration Checklist

- [ ] Extract Access schema and data
- [ ] Create Laravel migrations for all tables
- [ ] Define Eloquent models with relationships
- [ ] Import data from Access to new database
- [ ] Migrate VBA functions to Services
- [ ] Create controllers for CRUD operations
- [ ] Implement authentication/authorization
- [ ] Design RESTful API
- [ ] Handle file uploads/storage
- [ ] Create reports and exports
- [ ] Write feature tests
- [ ] Performance optimization (indexes, caching)
- [ ] API documentation

---

**Next:** [Vue.js UI Migration Guide](./06-VUE-UI-MIGRATION.md)
