# Business Logic Mapping

## Overview

This document maps VBA business logic to Laravel implementations, focusing on preserving functionality while leveraging modern patterns.

## Business Logic Categories

1. **Data Validation** - Ensuring data integrity
2. **Calculations** - Mathematical and business calculations
3. **Workflows** - Multi-step business processes
4. **Permissions** - Access control logic
5. **Notifications** - User alerts and communications
6. **Integrations** - External system interactions

## Validation Logic

### VBA Validation Pattern

**Access:**
```vba
Private Sub Form_BeforeUpdate(Cancel As Integer)
    ' Email validation
    If Not IsValidEmail(Me.Email) Then
        MsgBox "Invalid email format"
        Cancel = True
        Exit Sub
    End If

    ' Hire date validation
    If Me.HireDate > Date Then
        MsgBox "Hire date cannot be in the future"
        Cancel = True
        Exit Sub
    End If

    ' Salary validation
    If Me.Salary < 0 Then
        MsgBox "Salary must be positive"
        Cancel = True
        Exit Sub
    End If

    ' Custom business rule
    If Me.Position = "Manager" And IsNull(Me.ManagerTraining) Then
        MsgBox "Managers must complete manager training"
        Cancel = True
        Exit Sub
    End If
End Sub
```

### Laravel Form Request

**Laravel:**
```php
<?php
// app/Http/Requests/StoreEmployeeRequest.php

namespace App\Http\Requests;

use Illuminate\Foundation\Http\FormRequest;
use Illuminate\Validation\Rule;

class StoreEmployeeRequest extends FormRequest
{
    public function authorize(): bool
    {
        return auth()->check();
    }

    public function rules(): array
    {
        return [
            'first_name' => 'required|string|max:50',
            'last_name' => 'required|string|max:50',
            'email' => [
                'required',
                'email',
                Rule::unique('employees')->ignore($this->employee)
            ],
            'hire_date' => 'required|date|before_or_equal:today',
            'salary' => 'required|numeric|min:0',
            'department_id' => 'required|exists:departments,id',
            'position' => 'required|string',
            'manager_training' => [
                Rule::requiredIf(function () {
                    return $this->input('position') === 'Manager';
                })
            ]
        ];
    }

    public function messages(): array
    {
        return [
            'email.email' => 'Invalid email format',
            'hire_date.before_or_equal' => 'Hire date cannot be in the future',
            'salary.min' => 'Salary must be positive',
            'manager_training.required' => 'Managers must complete manager training'
        ];
    }

    // Custom validation
    public function withValidator($validator)
    {
        $validator->after(function ($validator) {
            if ($this->hasConflictingSchedule()) {
                $validator->errors()->add('schedule', 'Schedule conflicts with existing assignment');
            }
        });
    }

    private function hasConflictingSchedule(): bool
    {
        // Custom business logic
        return false;
    }
}
```

**Vue Client-Side Validation:**
```vue
<script setup>
import { useVuelidate } from '@vuelidate/core'
import { required, email, minValue, helpers } from '@vuelidate/validators'

const managerTrainingRequired = helpers.withMessage(
  'Managers must complete manager training',
  (value, siblings) => {
    if (siblings.position === 'Manager') {
      return value !== null && value !== ''
    }
    return true
  }
)

const rules = {
  first_name: { required },
  last_name: { required },
  email: { required, email },
  hire_date: {
    required,
    beforeToday: helpers.withMessage(
      'Hire date cannot be in the future',
      (value) => new Date(value) <= new Date()
    )
  },
  salary: {
    required,
    minValue: helpers.withMessage('Salary must be positive', minValue(0))
  },
  manager_training: { managerTrainingRequired }
}

const v$ = useVuelidate(rules, formData)
</script>
```

## Calculation Logic

### VBA Calculation Pattern

**Access:**
```vba
Public Function CalculateEmployeeBenefits(employeeID As Long) As Currency
    Dim rs As DAO.Recordset
    Dim baseSalary As Currency
    Dim healthInsurance As Currency
    Dim retirement401k As Currency
    Dim totalBenefits As Currency

    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblEmployees WHERE EmployeeID = " & employeeID)

    If Not rs.EOF Then
        baseSalary = rs!Salary

        ' Health insurance: 5% of salary
        healthInsurance = baseSalary * 0.05

        ' 401k match: up to 6% of salary
        Dim employeeContribution As Currency
        employeeContribution = rs!Retirement401kContribution
        If employeeContribution > baseSalary * 0.06 Then
            retirement401k = baseSalary * 0.06
        Else
            retirement401k = employeeContribution
        End If

        totalBenefits = healthInsurance + retirement401k
    End If

    rs.Close
    CalculateEmployeeBenefits = totalBenefits
End Function
```

### Laravel Service Pattern

**Laravel:**
```php
<?php
// app/Services/BenefitsCalculator.php

namespace App\Services;

use App\Models\Employee;

class BenefitsCalculator
{
    private const HEALTH_INSURANCE_RATE = 0.05;
    private const MAX_401K_MATCH_RATE = 0.06;

    public function calculateEmployeeBenefits(int $employeeId): float
    {
        $employee = Employee::findOrFail($employeeId);

        $healthInsurance = $this->calculateHealthInsurance($employee);
        $retirement401k = $this->calculate401kMatch($employee);

        return $healthInsurance + $retirement401k;
    }

    private function calculateHealthInsurance(Employee $employee): float
    {
        return $employee->salary * self::HEALTH_INSURANCE_RATE;
    }

    private function calculate401kMatch(Employee $employee): float
    {
        $maxMatch = $employee->salary * self::MAX_401K_MATCH_RATE;
        $employeeContribution = $employee->retirement_401k_contribution ?? 0;

        return min($employeeContribution, $maxMatch);
    }

    public function getBenefitsBreakdown(int $employeeId): array
    {
        $employee = Employee::findOrFail($employeeId);

        return [
            'health_insurance' => $this->calculateHealthInsurance($employee),
            '401k_match' => $this->calculate401kMatch($employee),
            'total' => $this->calculateEmployeeBenefits($employeeId)
        ];
    }
}

// Usage in Controller
use App\Services\BenefitsCalculator;

public function showBenefits(Employee $employee, BenefitsCalculator $calculator)
{
    $benefits = $calculator->getBenefitsBreakdown($employee->id);
    return response()->json($benefits);
}
```

**Vue Display:**
```vue
<script setup>
import { ref, onMounted } from 'vue'

const props = defineProps({
  employeeId: Number
})

const benefits = ref(null)

onMounted(async () => {
  const response = await fetch(`/api/employees/${props.employeeId}/benefits`)
  benefits.value = await response.json()
})

const formatCurrency = (amount) => {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD'
  }).format(amount)
}
</script>

<template>
  <div v-if="benefits" class="benefits-breakdown">
    <h3>Benefits Breakdown</h3>
    <table>
      <tr>
        <td>Health Insurance:</td>
        <td>{{ formatCurrency(benefits.health_insurance) }}</td>
      </tr>
      <tr>
        <td>401k Match:</td>
        <td>{{ formatCurrency(benefits['401k_match']) }}</td>
      </tr>
      <tr class="total">
        <td><strong>Total Benefits:</strong></td>
        <td><strong>{{ formatCurrency(benefits.total) }}</strong></td>
      </tr>
    </table>
  </div>
</template>
```

## Workflow Logic

### VBA Multi-Step Workflow

**Access:**
```vba
Public Sub ProcessEmployeeOnboarding(employeeID As Long)
    Dim db As DAO.Database
    Set db = CurrentDb

    ' Step 1: Create employee record
    DoCmd.OpenForm "frmNewEmployee", , , "EmployeeID = " & employeeID

    ' Step 2: Assign equipment
    Call AssignEquipment(employeeID)

    ' Step 3: Enroll in benefits
    Call EnrollInBenefits(employeeID)

    ' Step 4: Schedule training
    Call ScheduleTraining(employeeID)

    ' Step 5: Notify HR and Manager
    Call NotifyStakeholders(employeeID)

    ' Step 6: Update status
    db.Execute "UPDATE tblEmployees SET OnboardingStatus = 'Completed', OnboardingDate = #" & Date & "# WHERE EmployeeID = " & employeeID

    MsgBox "Onboarding process completed for employee " & employeeID
End Sub
```

### Laravel Job Chain Pattern

**Laravel:**
```php
<?php
// app/Services/OnboardingService.php

namespace App\Services;

use App\Models\Employee;
use App\Jobs\AssignEquipment;
use App\Jobs\EnrollInBenefits;
use App\Jobs\ScheduleTraining;
use App\Jobs\NotifyStakeholders;
use Illuminate\Support\Facades\Bus;

class OnboardingService
{
    public function processEmployeeOnboarding(Employee $employee): void
    {
        // Chain jobs for sequential execution
        Bus::chain([
            new AssignEquipment($employee),
            new EnrollInBenefits($employee),
            new ScheduleTraining($employee),
            new NotifyStakeholders($employee),
            function () use ($employee) {
                $employee->update([
                    'onboarding_status' => 'Completed',
                    'onboarding_date' => now()
                ]);
            }
        ])->dispatch();
    }
}

// app/Jobs/AssignEquipment.php
namespace App\Jobs;

use App\Models\Employee;
use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;

class AssignEquipment implements ShouldQueue
{
    use Dispatchable, Queueable;

    public function __construct(
        public Employee $employee
    ) {}

    public function handle(): void
    {
        // Assign laptop
        $this->employee->equipment()->create([
            'type' => 'laptop',
            'assigned_date' => now()
        ]);

        // Assign phone
        $this->employee->equipment()->create([
            'type' => 'phone',
            'assigned_date' => now()
        ]);
    }
}

// Similar pattern for other jobs...
```

**Alternative: Event/Listener Pattern**

```php
<?php
// app/Events/EmployeeCreated.php
namespace App\Events;

use App\Models\Employee;
use Illuminate\Foundation\Events\Dispatchable;

class EmployeeCreated
{
    use Dispatchable;

    public function __construct(
        public Employee $employee
    ) {}
}

// app/Listeners/StartOnboardingProcess.php
namespace App\Listeners;

use App\Events\EmployeeCreated;
use App\Services\OnboardingService;

class StartOnboardingProcess
{
    public function __construct(
        private OnboardingService $onboardingService
    ) {}

    public function handle(EmployeeCreated $event): void
    {
        $this->onboardingService->processEmployeeOnboarding($event->employee);
    }
}

// In EventServiceProvider
protected $listen = [
    EmployeeCreated::class => [
        StartOnboardingProcess::class,
    ],
];

// Trigger event after creating employee
event(new EmployeeCreated($employee));
```

## Permission Logic

### VBA Permission Check

**Access:**
```vba
Public Function HasPermission(permissionName As String) As Boolean
    Dim rs As DAO.Recordset
    Dim userGroup As String

    ' Get current user's group
    userGroup = DLookup("UserGroup", "tblUsers", "Username = '" & CurrentUser() & "'")

    ' Check if group has permission
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblPermissions WHERE GroupName = '" & userGroup & "' AND PermissionName = '" & permissionName & "'")

    HasPermission = Not rs.EOF
    rs.Close
End Function

Private Sub Form_Load()
    If Not HasPermission("EditEmployees") Then
        Me.AllowEdits = False
        Me.btnSave.Enabled = False
    End If

    If Not HasPermission("DeleteEmployees") Then
        Me.btnDelete.Visible = False
    End If
End Sub
```

### Laravel Policy Pattern

**Laravel:**
```php
<?php
// app/Policies/EmployeePolicy.php

namespace App\Policies;

use App\Models\User;
use App\Models\Employee;

class EmployeePolicy
{
    public function viewAny(User $user): bool
    {
        return $user->hasPermission('view-employees');
    }

    public function view(User $user, Employee $employee): bool
    {
        return $user->hasPermission('view-employees');
    }

    public function create(User $user): bool
    {
        return $user->hasPermission('create-employees');
    }

    public function update(User $user, Employee $employee): bool
    {
        return $user->hasPermission('edit-employees')
            || ($user->id === $employee->manager_id);
    }

    public function delete(User $user, Employee $employee): bool
    {
        return $user->hasPermission('delete-employees')
            && $employee->id !== $user->id; // Can't delete yourself
    }
}

// app/Models/User.php
public function hasPermission(string $permission): bool
{
    return $this->roles()
        ->whereHas('permissions', function ($query) use ($permission) {
            $query->where('name', $permission);
        })
        ->exists();
}

// In Controller
use App\Models\Employee;

public function update(Request $request, Employee $employee)
{
    $this->authorize('update', $employee);

    // Update logic
}

// Or in Blade/API
@can('update', $employee)
    <button>Edit</button>
@endcan
```

**Vue with Composable:**
```vue
<script setup>
import { useAuthStore } from '@/stores/auth'

const authStore = useAuthStore()

const canEdit = computed(() => {
  return authStore.hasPermission('edit-employees')
})

const canDelete = computed(() => {
  return authStore.hasPermission('delete-employees')
})
</script>

<template>
  <div>
    <button v-if="canEdit" @click="editEmployee">Edit</button>
    <button v-if="canDelete" @click="deleteEmployee">Delete</button>
  </div>
</template>
```

## Notification Logic

### VBA Notification Pattern

**Access:**
```vba
Public Sub NotifyManager(employeeID As Long, message As String)
    Dim rs As DAO.Recordset
    Dim managerEmail As String

    ' Get manager email
    Set rs = CurrentDb.OpenRecordset("SELECT m.Email FROM tblEmployees e INNER JOIN tblEmployees m ON e.ManagerID = m.EmployeeID WHERE e.EmployeeID = " & employeeID)

    If Not rs.EOF Then
        managerEmail = rs!Email

        ' Send email (using CDO or Outlook automation)
        Call SendEmail(managerEmail, "Employee Update", message)

        ' Log notification
        CurrentDb.Execute "INSERT INTO tblNotifications (EmployeeID, NotificationType, NotificationDate, Status) VALUES (" & employeeID & ", 'ManagerNotification', #" & Now & "#, 'Sent')"
    End If

    rs.Close
End Sub
```

### Laravel Notification Pattern

**Laravel:**
```php
<?php
// app/Notifications/EmployeeUpdated.php

namespace App\Notifications;

use App\Models\Employee;
use Illuminate\Bus\Queueable;
use Illuminate\Notifications\Notification;
use Illuminate\Notifications\Messages\MailMessage;

class EmployeeUpdated extends Notification
{
    use Queueable;

    public function __construct(
        public Employee $employee,
        public string $message
    ) {}

    public function via($notifiable): array
    {
        return ['mail', 'database'];
    }

    public function toMail($notifiable): MailMessage
    {
        return (new MailMessage)
            ->subject('Employee Update')
            ->line($this->message)
            ->action('View Employee', url("/employees/{$this->employee->id}"));
    }

    public function toArray($notifiable): array
    {
        return [
            'employee_id' => $this->employee->id,
            'message' => $this->message
        ];
    }
}

// Usage
$employee = Employee::find($employeeId);
$manager = $employee->manager;

$manager->notify(new EmployeeUpdated($employee, 'Employee information has been updated'));

// Or using event listener
// In EmployeeObserver
public function updated(Employee $employee)
{
    if ($employee->manager) {
        $employee->manager->notify(
            new EmployeeUpdated($employee, 'Employee information has been updated')
        );
    }
}
```

## Integration Logic

### VBA External Integration

**Access:**
```vba
Public Function SyncWithPayrollSystem(employeeID As Long) As Boolean
    Dim http As Object
    Dim jsonData As String
    Dim rs As DAO.Recordset

    Set http = CreateObject("MSXML2.XMLHTTP")
    Set rs = CurrentDb.OpenRecordset("SELECT * FROM tblEmployees WHERE EmployeeID = " & employeeID)

    If Not rs.EOF Then
        ' Build JSON
        jsonData = "{""employee_id"":" & rs!EmployeeID & _
                   ",""name"":""" & rs!FirstName & " " & rs!LastName & """" & _
                   ",""salary"":" & rs!Salary & "}"

        ' Send to external API
        http.Open "POST", "https://payroll.example.com/api/sync", False
        http.setRequestHeader "Content-Type", "application/json"
        http.Send jsonData

        If http.Status = 200 Then
            SyncWithPayrollSystem = True
        Else
            SyncWithPayrollSystem = False
        End If
    End If

    rs.Close
End Function
```

### Laravel HTTP Client Pattern

**Laravel:**
```php
<?php
// app/Services/PayrollIntegrationService.php

namespace App\Services;

use App\Models\Employee;
use Illuminate\Support\Facades\Http;
use Illuminate\Support\Facades\Log;

class PayrollIntegrationService
{
    private string $baseUrl;
    private string $apiKey;

    public function __construct()
    {
        $this->baseUrl = config('services.payroll.url');
        $this->apiKey = config('services.payroll.api_key');
    }

    public function syncEmployee(Employee $employee): bool
    {
        try {
            $response = Http::withToken($this->apiKey)
                ->timeout(30)
                ->post("{$this->baseUrl}/api/sync", [
                    'employee_id' => $employee->id,
                    'name' => $employee->full_name,
                    'salary' => $employee->salary,
                    'department' => $employee->department->name
                ]);

            if ($response->successful()) {
                Log::info("Employee {$employee->id} synced successfully");
                return true;
            }

            Log::error("Failed to sync employee {$employee->id}: " . $response->body());
            return false;

        } catch (\Exception $e) {
            Log::error("Exception syncing employee {$employee->id}: " . $e->getMessage());
            return false;
        }
    }

    public function syncAllEmployees(): array
    {
        $results = ['success' => 0, 'failed' => 0];

        Employee::active()->chunk(100, function ($employees) use (&$results) {
            foreach ($employees as $employee) {
                if ($this->syncEmployee($employee)) {
                    $results['success']++;
                } else {
                    $results['failed']++;
                }
            }
        });

        return $results;
    }
}
```

## Business Rule Examples

### Rule: Automatic Status Updates

**VBA:**
```vba
Private Sub txtEndDate_AfterUpdate()
    If Not IsNull(Me.txtEndDate) And Me.txtEndDate <= Date Then
        Me.Status = "Completed"
        Me.CompletedBy = CurrentUser()
    End If
End Sub
```

**Laravel Observer:**
```php
<?php
// app/Observers/ProjectObserver.php

namespace App\Observers;

use App\Models\Project;

class ProjectObserver
{
    public function updating(Project $project): void
    {
        if ($project->isDirty('end_date')) {
            if ($project->end_date && $project->end_date->isPast()) {
                $project->status = 'Completed';
                $project->completed_by = auth()->id();
            }
        }
    }
}
```

### Rule: Dependent Field Updates

**VBA:**
```vba
Private Sub cboEmploymentType_AfterUpdate()
    If Me.cboEmploymentType = "Full-Time" Then
        Me.Benefits = True
        Me.VacationDays = 15
    ElseIf Me.cboEmploymentType = "Part-Time" Then
        Me.Benefits = False
        Me.VacationDays = 5
    Else ' Contract
        Me.Benefits = False
        Me.VacationDays = 0
    End If
End Sub
```

**Laravel Mutator:**
```php
<?php
// app/Models/Employee.php

public function setEmploymentTypeAttribute($value): void
{
    $this->attributes['employment_type'] = $value;

    // Auto-update related fields
    match($value) {
        'Full-Time' => [
            $this->attributes['benefits'] = true,
            $this->attributes['vacation_days'] = 15
        ],
        'Part-Time' => [
            $this->attributes['benefits'] = false,
            $this->attributes['vacation_days'] = 5
        ],
        'Contract' => [
            $this->attributes['benefits'] = false,
            $this->attributes['vacation_days'] = 0
        ]
    };
}
```

## Migration Checklist

- [ ] Identify all business rules in VBA code
- [ ] Categorize rules (validation, calculation, workflow, etc.)
- [ ] Map VBA functions to Laravel equivalents
- [ ] Create services for complex business logic
- [ ] Implement validation in Form Requests
- [ ] Set up policies for authorization
- [ ] Create observers for model events
- [ ] Implement notifications
- [ ] Set up job queues for workflows
- [ ] Write tests for all business logic
- [ ] Document all business rules
- [ ] Verify calculations match VBA results

---

**See Also:**
- [VBA Modules Reference](./03-VBA-MODULES.md)
- [Laravel Migration Guide](./05-LARAVEL-MIGRATION.md)
- [Translation Dictionary](./07-TRANSLATION-DICTIONARY.md)
