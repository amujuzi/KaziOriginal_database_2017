# Forms and UI Documentation

## Overview

This document catalogs all MS Access forms, their purposes, controls, and provides guidance for recreating them in Vue.js.

## How to Document Forms

### Using Microsoft Access

1. Open each form in Design View
2. Note all controls and their properties
3. Review Form properties (RecordSource, AllowEdits, etc.)
4. Document event procedures (View Code button)
5. Take screenshots for reference

### Form Export

```vba
' Export form layouts to text files
Sub ExportForms()
    Dim obj As AccessObject
    For Each obj In CurrentProject.AllForms
        Application.SaveAsText acForm, obj.Name, "C:\export\forms\" & obj.Name & ".txt"
    Next obj
End Sub
```

## Forms Inventory

[TO BE COMPLETED - List all forms from the database]

| Form Name | Purpose | Type | Record Source | Complexity |
|-----------|---------|------|---------------|------------|
| frmMain | Main menu/dashboard | Unbound | None | Low |
| frmEmployees | Employee entry | Bound | tblEmployees | Medium |
| frmEmployeeList | Employee list view | Continuous | qryEmployees | Medium |
| frmProjects | Project management | Bound | tblProjects | High |
| ... | ... | ... | ... | ... |

## Form Types and Vue Equivalents

### 1. Single Form (Detail View)

**Access:** Bound form showing one record at a time

**Vue Equivalent:** Form component with single object

**Example:**

```vue
<!-- EmployeeDetail.vue -->
<script setup>
import { ref, onMounted } from 'vue'
import { useRoute } from 'vue-router'

const route = useRoute()
const employee = ref(null)

onMounted(async () => {
  const response = await fetch(`/api/employees/${route.params.id}`)
  employee.value = await response.json()
})
</script>

<template>
  <div v-if="employee" class="employee-detail">
    <h2>{{ employee.full_name }}</h2>
    <div class="field">
      <label>Email:</label>
      <span>{{ employee.email }}</span>
    </div>
    <div class="field">
      <label>Department:</label>
      <span>{{ employee.department.name }}</span>
    </div>
  </div>
</template>
```

### 2. Continuous Forms (List View)

**Access:** Multiple records displayed in a list

**Vue Equivalent:** List component with v-for

**Example:**

```vue
<!-- EmployeeList.vue -->
<script setup>
import { ref, onMounted } from 'vue'

const employees = ref([])

onMounted(async () => {
  const response = await fetch('/api/employees')
  employees.value = await response.json()
})
</script>

<template>
  <div class="employee-list">
    <div v-for="employee in employees" :key="employee.id" class="employee-item">
      <span>{{ employee.full_name }}</span>
      <span>{{ employee.department.name }}</span>
      <button @click="editEmployee(employee.id)">Edit</button>
    </div>
  </div>
</template>
```

### 3. Datasheet Forms

**Access:** Spreadsheet-like view

**Vue Equivalent:** Data table component

### 4. Split Forms

**Access:** Form + Datasheet combined

**Vue Equivalent:** Master-detail layout with separate components

### 5. Dialog Forms

**Access:** Popup forms

**Vue Equivalent:** Modal components

## Example Form Documentation

---

### Form: `frmEmployees`

**Purpose:** Add and edit employee records

**Form Properties:**
- Record Source: `tblEmployees`
- Default View: Single Form
- Allow Additions: Yes
- Allow Deletions: Yes
- Allow Edits: Yes
- Data Entry: No
- Modal: No
- Pop Up: No

**Controls:**

| Control Name | Type | Control Source | Properties | Purpose |
|--------------|------|----------------|------------|---------|
| txtEmployeeID | TextBox | EmployeeID | Enabled: No | Display ID (AutoNumber) |
| txtFirstName | TextBox | FirstName | Required | First name entry |
| txtLastName | TextBox | LastName | Required | Last name entry |
| txtEmail | TextBox | Email | Validation Rule | Email address |
| cboDepartment | ComboBox | DepartmentID | Row Source: qryDepartments | Select department |
| cboManager | ComboBox | ManagerID | Row Source: Dynamic | Select manager |
| txtHireDate | TextBox | HireDate | Format: Short Date | Hire date |
| chkIsActive | CheckBox | IsActive | Default: True | Active status |
| txtSalary | TextBox | Salary | Format: Currency | Salary amount |
| txtNotes | TextBox | Notes | Multiline | Additional notes |
| btnSave | Button | - | - | Save record |
| btnCancel | Button | - | - | Cancel changes |
| btnDelete | Button | - | - | Delete record |

**Subforms:**
- `subfrmEmployeeTasks` - Displays tasks assigned to employee

**Event Procedures:**
- Form_Load - Initialize form, load dropdowns
- Form_BeforeUpdate - Validate before saving
- Form_Current - Update UI based on current record
- cboDepartment_AfterUpdate - Filter managers by department
- btnSave_Click - Save record with validation
- btnCancel_Click - Undo changes
- btnDelete_Click - Delete with confirmation

**Vue Component Structure:**

```vue
<!-- EmployeeForm.vue -->
<script setup>
import { ref, onMounted, watch } from 'vue'
import { useRoute, useRouter } from 'vue-router'
import { useVuelidate } from '@vuelidate/core'
import { required, email } from '@vuelidate/validators'
import EmployeeTasks from './EmployeeTasks.vue'

const route = useRoute()
const router = useRouter()

const employee = ref({
  id: null,
  first_name: '',
  last_name: '',
  email: '',
  department_id: null,
  manager_id: null,
  hire_date: new Date().toISOString().split('T')[0],
  is_active: true,
  salary: 0,
  notes: ''
})

const departments = ref([])
const managers = ref([])
const isNewRecord = computed(() => !route.params.id)

// Validation rules
const rules = {
  first_name: { required },
  last_name: { required },
  email: { required, email },
  department_id: { required }
}

const v$ = useVuelidate(rules, employee)

// Load form data
onMounted(async () => {
  // Load departments
  const deptResponse = await fetch('/api/departments')
  departments.value = await deptResponse.json()

  // Load employee if editing
  if (!isNewRecord.value) {
    const empResponse = await fetch(`/api/employees/${route.params.id}`)
    employee.value = await empResponse.json()
  }
})

// Watch department changes to filter managers
watch(() => employee.value.department_id, async (newDeptId) => {
  if (newDeptId) {
    const response = await fetch(`/api/employees?department_id=${newDeptId}&is_manager=true`)
    managers.value = await response.json()
  } else {
    managers.value = []
    employee.value.manager_id = null
  }
})

// Save employee
const saveEmployee = async () => {
  const isValid = await v$.value.$validate()
  if (!isValid) return

  try {
    const url = isNewRecord.value
      ? '/api/employees'
      : `/api/employees/${employee.value.id}`

    const response = await fetch(url, {
      method: isNewRecord.value ? 'POST' : 'PUT',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(employee.value)
    })

    if (response.ok) {
      toast.success('Employee saved successfully')
      router.push({ name: 'employees' })
    }
  } catch (error) {
    toast.error('Error saving employee')
  }
}

// Cancel changes
const cancelChanges = () => {
  router.back()
}

// Delete employee
const deleteEmployee = async () => {
  if (!confirm('Are you sure you want to delete this employee?')) return

  try {
    const response = await fetch(`/api/employees/${employee.value.id}`, {
      method: 'DELETE'
    })

    if (response.ok) {
      toast.success('Employee deleted successfully')
      router.push({ name: 'employees' })
    }
  } catch (error) {
    toast.error('Error deleting employee')
  }
}
</script>

<template>
  <div class="employee-form">
    <h2>{{ isNewRecord ? 'New Employee' : 'Edit Employee' }}</h2>

    <form @submit.prevent="saveEmployee">
      <div class="form-row">
        <div class="form-group">
          <label>First Name: *</label>
          <input v-model="employee.first_name" type="text" />
          <span v-if="v$.first_name.$error" class="error">Required</span>
        </div>

        <div class="form-group">
          <label>Last Name: *</label>
          <input v-model="employee.last_name" type="text" />
          <span v-if="v$.last_name.$error" class="error">Required</span>
        </div>
      </div>

      <div class="form-group">
        <label>Email: *</label>
        <input v-model="employee.email" type="email" />
        <span v-if="v$.email.$error" class="error">Valid email required</span>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label>Department: *</label>
          <select v-model="employee.department_id">
            <option :value="null">Select Department</option>
            <option
              v-for="dept in departments"
              :key="dept.id"
              :value="dept.id"
            >
              {{ dept.name }}
            </option>
          </select>
        </div>

        <div class="form-group">
          <label>Manager:</label>
          <select
            v-model="employee.manager_id"
            :disabled="!employee.department_id"
          >
            <option :value="null">Select Manager</option>
            <option
              v-for="manager in managers"
              :key="manager.id"
              :value="manager.id"
            >
              {{ manager.full_name }}
            </option>
          </select>
        </div>
      </div>

      <div class="form-row">
        <div class="form-group">
          <label>Hire Date:</label>
          <input v-model="employee.hire_date" type="date" />
        </div>

        <div class="form-group">
          <label>Salary:</label>
          <input v-model.number="employee.salary" type="number" step="0.01" />
        </div>
      </div>

      <div class="form-group">
        <label>
          <input v-model="employee.is_active" type="checkbox" />
          Active
        </label>
      </div>

      <div class="form-group">
        <label>Notes:</label>
        <textarea v-model="employee.notes" rows="4"></textarea>
      </div>

      <!-- Subform: Employee Tasks -->
      <EmployeeTasks
        v-if="!isNewRecord"
        :employee-id="employee.id"
      />

      <div class="form-actions">
        <button type="submit" class="btn-primary">Save</button>
        <button type="button" class="btn-secondary" @click="cancelChanges">
          Cancel
        </button>
        <button
          v-if="!isNewRecord"
          type="button"
          class="btn-danger"
          @click="deleteEmployee"
        >
          Delete
        </button>
      </div>
    </form>
  </div>
</template>

<style scoped>
.employee-form {
  max-width: 800px;
  margin: 0 auto;
  padding: 20px;
}

.form-row {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 15px;
}

.form-group {
  margin-bottom: 15px;
}

.form-group label {
  display: block;
  margin-bottom: 5px;
  font-weight: bold;
}

.form-group input,
.form-group select,
.form-group textarea {
  width: 100%;
  padding: 8px;
  border: 1px solid #ccc;
  border-radius: 4px;
}

.error {
  color: red;
  font-size: 0.85em;
}

.form-actions {
  margin-top: 20px;
  display: flex;
  gap: 10px;
}

.btn-primary {
  background-color: #007bff;
  color: white;
  padding: 10px 20px;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.btn-secondary {
  background-color: #6c757d;
  color: white;
  padding: 10px 20px;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}

.btn-danger {
  background-color: #dc3545;
  color: white;
  padding: 10px 20px;
  border: none;
  border-radius: 4px;
  cursor: pointer;
}
</style>
```

---

### Form: `frmMain`

**Purpose:** Main navigation menu and dashboard

**Type:** Unbound form (no data source)

**Controls:**
- Navigation buttons for each major form
- Summary labels showing key metrics
- Welcome message with current user

**Vue Equivalent:**

```vue
<!-- Dashboard.vue -->
<script setup>
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { useAuthStore } from '@/stores/auth'

const router = useRouter()
const authStore = useAuthStore()

const stats = ref({
  total_employees: 0,
  active_projects: 0,
  pending_tasks: 0
})

onMounted(async () => {
  const response = await fetch('/api/dashboard/stats')
  stats.value = await response.json()
})

const navigate = (routeName) => {
  router.push({ name: routeName })
}
</script>

<template>
  <div class="dashboard">
    <h1>Welcome, {{ authStore.user.name }}</h1>

    <div class="stats">
      <div class="stat-card">
        <h3>{{ stats.total_employees }}</h3>
        <p>Total Employees</p>
      </div>
      <div class="stat-card">
        <h3>{{ stats.active_projects }}</h3>
        <p>Active Projects</p>
      </div>
      <div class="stat-card">
        <h3>{{ stats.pending_tasks }}</h3>
        <p>Pending Tasks</p>
      </div>
    </div>

    <div class="navigation">
      <button @click="navigate('employees')">Employees</button>
      <button @click="navigate('projects')">Projects</button>
      <button @click="navigate('tasks')">Tasks</button>
      <button @click="navigate('reports')">Reports</button>
    </div>
  </div>
</template>
```

---

## Common Form Patterns

### Pattern 1: Lookup/Dropdown Population

**Access:**
```vba
Me.cboDepartment.RowSource = "SELECT DepartmentID, DepartmentName FROM tblDepartments ORDER BY DepartmentName"
```

**Vue:**
```vue
<script setup>
const departments = ref([])

onMounted(async () => {
  const response = await fetch('/api/departments?sort=name')
  departments.value = await response.json()
})
</script>

<template>
  <select v-model="selectedDepartment">
    <option v-for="dept in departments" :key="dept.id" :value="dept.id">
      {{ dept.name }}
    </option>
  </select>
</template>
```

### Pattern 2: Cascade Dropdowns

**Access:**
```vba
Private Sub cboDepartment_AfterUpdate()
    Me.cboManager.RowSource = "SELECT EmployeeID, FullName FROM tblEmployees WHERE DepartmentID = " & Me.cboDepartment
    Me.cboManager.Requery
End Sub
```

**Vue:**
```vue
<script setup>
watch(() => selectedDepartment.value, async (deptId) => {
  if (deptId) {
    const response = await fetch(`/api/employees?department_id=${deptId}`)
    managers.value = await response.json()
  } else {
    managers.value = []
  }
})
</script>
```

### Pattern 3: Calculated Fields

**Access:**
```vba
Private Sub txtQuantity_AfterUpdate()
    Me.txtTotal = Me.txtQuantity * Me.txtUnitPrice
End Sub
```

**Vue:**
```vue
<script setup>
const quantity = ref(0)
const unitPrice = ref(0)
const total = computed(() => quantity.value * unitPrice.value)
</script>
```

### Pattern 4: Conditional Visibility

**Access:**
```vba
If Me.Status = "Completed" Then
    Me.txtCompletedDate.Visible = True
Else
    Me.txtCompletedDate.Visible = False
End If
```

**Vue:**
```vue
<input
  v-show="status === 'Completed'"
  v-model="completedDate"
  type="date"
/>
```

### Pattern 5: Master-Detail (Subforms)

**Access:**
Subform control with LinkMasterFields and LinkChildFields

**Vue:**
Parent-child components with props

## Form Migration Checklist

For each form:

- [ ] Document form purpose and type
- [ ] List all controls and properties
- [ ] Identify record source (table/query)
- [ ] Document all event procedures
- [ ] Take screenshots
- [ ] Note validation rules
- [ ] Identify subforms
- [ ] Map to Vue component structure
- [ ] Create component file
- [ ] Implement data fetching
- [ ] Add validation
- [ ] Test all interactions

## UI/UX Improvements for Modern Web

When migrating to Vue, consider these improvements:

1. **Responsive Design** - Forms adapt to screen size
2. **Real-time Validation** - Immediate feedback
3. **Better Error Messages** - Clear, actionable messages
4. **Loading States** - Show progress indicators
5. **Toast Notifications** - Non-intrusive feedback
6. **Keyboard Shortcuts** - Improve efficiency
7. **Accessibility** - ARIA labels, keyboard navigation
8. **Modern UI Components** - Use component libraries (Vuetify, PrimeVue, etc.)

---

**See Also:**
- [Vue UI Migration Guide](./06-VUE-UI-MIGRATION.md)
- [VBA Modules Reference](./03-VBA-MODULES.md)
- [Translation Dictionary](./07-TRANSLATION-DICTIONARY.md)
