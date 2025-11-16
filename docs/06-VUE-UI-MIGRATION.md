# Vue.js UI Migration Guide

## Overview

This guide helps you translate MS Access forms and controls into modern Vue.js components, maintaining functionality while improving user experience.

## Table of Contents

1. [Form Architecture](#form-architecture)
2. [Access Controls → Vue Components](#access-controls--vue-components)
3. [Data Binding](#data-binding)
4. [Form Validation](#form-validation)
5. [Navigation & Routing](#navigation--routing)
6. [Subforms & Nested Components](#subforms--nested-components)
7. [Reports & Data Display](#reports--data-display)
8. [Best Practices](#best-practices)

## Form Architecture

### Access Forms vs Vue Components

**Access Form Structure:**
- Bound to a table/query
- Contains controls (textboxes, combos, buttons)
- VBA code handles events
- Forms navigation (DoCmd)

**Vue Component Structure:**
- Reactive data from API
- Template with UI elements
- Composition API handles logic
- Vue Router for navigation

### Basic Form Migration Pattern

**Access Form (Employee Entry):**
```vba
' Form: frmEmployees
' RecordSource: tblEmployees

Private Sub Form_Load()
    Me.cboDepartment.RowSource = "SELECT DepartmentID, DepartmentName FROM tblDepartments"
End Sub

Private Sub btnSave_Click()
    DoCmd.RunCommand acCmdSaveRecord
    MsgBox "Employee saved"
End Sub
```

**Vue 3 Component:**
```vue
<!-- EmployeeForm.vue -->
<script setup>
import { ref, onMounted } from 'vue'
import { useRouter } from 'vue-router'
import { useToast } from '@/composables/useToast'

const router = useRouter()
const toast = useToast()

// Form data (bound to inputs)
const employee = ref({
  firstName: '',
  lastName: '',
  email: '',
  departmentId: null,
  isActive: true
})

const departments = ref([])
const loading = ref(false)

// Load departments (like RowSource)
onMounted(async () => {
  const response = await fetch('/api/departments')
  departments.value = await response.json()
})

// Save employee (like btnSave_Click)
const saveEmployee = async () => {
  loading.value = true
  try {
    const response = await fetch('/api/employees', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(employee.value)
    })

    if (response.ok) {
      toast.success('Employee saved successfully')
      router.push({ name: 'employees' })
    }
  } catch (error) {
    toast.error('Error saving employee')
  } finally {
    loading.value = false
  }
}
</script>

<template>
  <div class="employee-form">
    <h2>Employee Entry</h2>

    <form @submit.prevent="saveEmployee">
      <div class="form-group">
        <label>First Name:</label>
        <input v-model="employee.firstName" type="text" required />
      </div>

      <div class="form-group">
        <label>Last Name:</label>
        <input v-model="employee.lastName" type="text" required />
      </div>

      <div class="form-group">
        <label>Email:</label>
        <input v-model="employee.email" type="email" required />
      </div>

      <div class="form-group">
        <label>Department:</label>
        <select v-model="employee.departmentId" required>
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
        <label>
          <input v-model="employee.isActive" type="checkbox" />
          Active
        </label>
      </div>

      <button type="submit" :disabled="loading">
        {{ loading ? 'Saving...' : 'Save Employee' }}
      </button>
    </form>
  </div>
</template>

<style scoped>
.employee-form {
  max-width: 600px;
  margin: 0 auto;
  padding: 20px;
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
.form-group select {
  width: 100%;
  padding: 8px;
  border: 1px solid #ccc;
  border-radius: 4px;
}
</style>
```

## Access Controls → Vue Components

### Control Mapping Reference

| Access Control | Vue Element | Notes |
|----------------|-------------|-------|
| TextBox | `<input type="text">` | Use v-model for binding |
| ComboBox (dropdown) | `<select>` | Populate with v-for |
| ComboBox (autocomplete) | `<input>` + library | Use vue-select or similar |
| ListBox | `<select multiple>` | Or custom component |
| CheckBox | `<input type="checkbox">` | Boolean binding |
| OptionGroup | `<input type="radio">` | Same name attribute |
| Button | `<button>` | Use @click event |
| Label | `<label>` or `<span>` | Static or dynamic text |
| Image | `<img>` | Bind :src attribute |
| SubForm | Custom Component | Pass data as props |
| Tab Control | Tab Component | Use vue-tabs or custom |
| Command Button | `<button>` | Style appropriately |

### TextBox Examples

**Access:**
```vba
Me.txtFirstName.Value = "John"
Me.txtFirstName.SetFocus
Me.txtFirstName.Enabled = False
```

**Vue:**
```vue
<script setup>
import { ref } from 'vue'

const firstName = ref('John')
const firstNameInput = ref(null)
const isDisabled = ref(false)

const focusFirstName = () => {
  firstNameInput.value.focus()
}
</script>

<template>
  <input
    ref="firstNameInput"
    v-model="firstName"
    :disabled="isDisabled"
    type="text"
  />
</template>
```

### ComboBox/Dropdown Examples

**Access:**
```vba
Me.cboDepartment.RowSource = "SELECT DepartmentID, DepartmentName FROM tblDepartments"
Me.cboDepartment.Value = 5
```

**Vue (Simple Select):**
```vue
<script setup>
import { ref, onMounted } from 'vue'

const selectedDepartment = ref(null)
const departments = ref([])

onMounted(async () => {
  const response = await fetch('/api/departments')
  departments.value = await response.json()
  selectedDepartment.value = 5
})
</script>

<template>
  <select v-model="selectedDepartment">
    <option :value="null">Select Department</option>
    <option
      v-for="dept in departments"
      :key="dept.id"
      :value="dept.id"
    >
      {{ dept.name }}
    </option>
  </select>
</template>
```

**Vue (Searchable - using vue-select):**
```bash
npm install vue-select
```

```vue
<script setup>
import { ref, onMounted } from 'vue'
import vSelect from 'vue-select'
import 'vue-select/dist/vue-select.css'

const selectedDepartment = ref(null)
const departments = ref([])

onMounted(async () => {
  const response = await fetch('/api/departments')
  departments.value = await response.json()
})
</script>

<template>
  <v-select
    v-model="selectedDepartment"
    :options="departments"
    label="name"
    :reduce="dept => dept.id"
    placeholder="Search departments..."
  />
</template>
```

### Checkbox Examples

**Access:**
```vba
Me.chkActive.Value = True
If Me.chkActive.Value = True Then
    ' Do something
End If
```

**Vue:**
```vue
<script setup>
import { ref, watch } from 'vue'

const isActive = ref(true)

watch(isActive, (newValue) => {
  if (newValue) {
    console.log('Checkbox is checked')
  }
})
</script>

<template>
  <label>
    <input v-model="isActive" type="checkbox" />
    Active
  </label>
</template>
```

### Option Group/Radio Buttons

**Access:**
```vba
Me.optGender.Value = 1 ' Male
```

**Vue:**
```vue
<script setup>
import { ref } from 'vue'

const gender = ref('male')
</script>

<template>
  <div>
    <label>
      <input v-model="gender" type="radio" value="male" />
      Male
    </label>
    <label>
      <input v-model="gender" type="radio" value="female" />
      Female
    </label>
    <label>
      <input v-model="gender" type="radio" value="other" />
      Other
    </label>
  </div>
</template>
```

### Date Picker

**Access:**
Uses built-in calendar control

**Vue (using VueDatePicker):**
```bash
npm install @vuepic/vue-datepicker
```

```vue
<script setup>
import { ref } from 'vue'
import VueDatePicker from '@vuepic/vue-datepicker'
import '@vuepic/vue-datepicker/dist/main.css'

const hireDate = ref(new Date())
</script>

<template>
  <VueDatePicker v-model="hireDate" />
</template>
```

## Data Binding

### Two-Way Binding

**Access (Manual):**
```vba
Private Sub txtFirstName_AfterUpdate()
    ' Manual sync to variable or field
    currentFirstName = Me.txtFirstName.Value
End Sub
```

**Vue (Automatic with v-model):**
```vue
<script setup>
import { ref } from 'vue'

const firstName = ref('')
// firstName automatically syncs with input
</script>

<template>
  <input v-model="firstName" type="text" />
  <p>You typed: {{ firstName }}</p>
</template>
```

### Computed Properties (Calculated Fields)

**Access:**
```vba
Private Sub Form_Current()
    Me.txtFullName = Me.txtFirstName & " " & Me.txtLastName
    Me.txtTotal = Me.txtQuantity * Me.txtPrice
End Sub
```

**Vue:**
```vue
<script setup>
import { ref, computed } from 'vue'

const firstName = ref('')
const lastName = ref('')
const quantity = ref(0)
const price = ref(0)

const fullName = computed(() => `${firstName.value} ${lastName.value}`)
const total = computed(() => quantity.value * price.value)
</script>

<template>
  <p>Full Name: {{ fullName }}</p>
  <p>Total: ${{ total.toFixed(2) }}</p>
</template>
```

### Conditional Rendering

**Access:**
```vba
Private Sub Form_Current()
    If Me.Status = "Inactive" Then
        Me.btnEdit.Enabled = False
        Me.lblWarning.Visible = True
    Else
        Me.btnEdit.Enabled = True
        Me.lblWarning.Visible = False
    End If
End Sub
```

**Vue:**
```vue
<script setup>
import { ref } from 'vue'

const status = ref('Active')
</script>

<template>
  <button :disabled="status === 'Inactive'">Edit</button>
  <p v-show="status === 'Inactive'" class="warning">
    This record is inactive
  </p>
</template>
```

## Form Validation

### Access Validation

**Access:**
```vba
Private Sub btnSave_Click()
    If IsNull(Me.txtEmail) Then
        MsgBox "Email is required", vbCritical
        Me.txtEmail.SetFocus
        Exit Sub
    End If

    If Not IsValidEmail(Me.txtEmail) Then
        MsgBox "Invalid email format", vbCritical
        Exit Sub
    End If

    ' Save logic
End Sub
```

### Vue Validation (using Vuelidate)

```bash
npm install @vuelidate/core @vuelidate/validators
```

```vue
<script setup>
import { ref, computed } from 'vue'
import { useVuelidate } from '@vuelidate/core'
import { required, email, minLength } from '@vuelidate/validators'

const formData = ref({
  firstName: '',
  lastName: '',
  email: '',
  password: ''
})

const rules = computed(() => ({
  firstName: { required },
  lastName: { required },
  email: { required, email },
  password: { required, minLength: minLength(8) }
}))

const v$ = useVuelidate(rules, formData)

const submitForm = async () => {
  const isValid = await v$.value.$validate()

  if (!isValid) {
    return
  }

  // Submit logic
  console.log('Form is valid!', formData.value)
}
</script>

<template>
  <form @submit.prevent="submitForm">
    <div>
      <input v-model="formData.firstName" type="text" />
      <span v-if="v$.firstName.$error" class="error">
        First name is required
      </span>
    </div>

    <div>
      <input v-model="formData.email" type="email" />
      <span v-if="v$.email.$error" class="error">
        Valid email is required
      </span>
    </div>

    <div>
      <input v-model="formData.password" type="password" />
      <span v-if="v$.password.$error" class="error">
        Password must be at least 8 characters
      </span>
    </div>

    <button type="submit">Save</button>
  </form>
</template>
```

### Custom Validation Rules

```vue
<script setup>
import { helpers } from '@vuelidate/validators'

const isUniqueUsername = async (value) => {
  if (!value) return true
  const response = await fetch(`/api/check-username/${value}`)
  const { available } = await response.json()
  return available
}

const rules = computed(() => ({
  username: {
    required,
    unique: helpers.withAsync(isUniqueUsername)
  }
}))
</script>
```

## Navigation & Routing

### Access Navigation

**Access:**
```vba
DoCmd.OpenForm "frmEmployeeList"
DoCmd.OpenForm "frmEmployeeEdit", , , "EmployeeID = " & Me.EmployeeID
DoCmd.Close acForm, "frmEmployees"
DoCmd.GoToRecord , , acNewRec
```

### Vue Router

**Setup:**
```bash
npm install vue-router
```

**Router Configuration:**
```javascript
// router/index.js
import { createRouter, createWebHistory } from 'vue-router'

const routes = [
  {
    path: '/employees',
    name: 'employees',
    component: () => import('@/views/EmployeeList.vue')
  },
  {
    path: '/employees/create',
    name: 'employees.create',
    component: () => import('@/views/EmployeeForm.vue')
  },
  {
    path: '/employees/:id/edit',
    name: 'employees.edit',
    component: () => import('@/views/EmployeeForm.vue'),
    props: true
  },
  {
    path: '/employees/:id',
    name: 'employees.show',
    component: () => import('@/views/EmployeeDetail.vue'),
    props: true
  }
]

const router = createRouter({
  history: createWebHistory(),
  routes
})

export default router
```

**Navigation in Components:**
```vue
<script setup>
import { useRouter } from 'vue-router'

const router = useRouter()

// Navigate to list
const goToList = () => {
  router.push({ name: 'employees' })
}

// Navigate to edit with ID
const editEmployee = (employeeId) => {
  router.push({ name: 'employees.edit', params: { id: employeeId } })
}

// Go back
const goBack = () => {
  router.back()
}
</script>

<template>
  <button @click="goToList">Back to List</button>
  <button @click="editEmployee(123)">Edit Employee</button>

  <!-- Or use router-link -->
  <router-link :to="{ name: 'employees' }">Employee List</router-link>
</template>
```

## Subforms & Nested Components

### Access Subform

**Access:**
```vba
' Main Form: frmOrders
' SubForm: subfrmOrderItems
' Link: OrderID → OrderID

Private Sub Form_Current()
    Me.subfrmOrderItems.Requery
End Sub

' Access subform control
Me.subfrmOrderItems.Form!txtQuantity.Value = 10
```

### Vue Parent-Child Components

**Parent Component (Order.vue):**
```vue
<script setup>
import { ref, onMounted } from 'vue'
import OrderItems from '@/components/OrderItems.vue'

const props = defineProps({
  id: [String, Number]
})

const order = ref(null)
const orderItems = ref([])

onMounted(async () => {
  const response = await fetch(`/api/orders/${props.id}`)
  const data = await response.json()
  order.value = data
  orderItems.value = data.items
})

const handleItemsUpdate = (updatedItems) => {
  orderItems.value = updatedItems
}
</script>

<template>
  <div v-if="order" class="order-detail">
    <h2>Order #{{ order.id }}</h2>
    <p>Customer: {{ order.customer_name }}</p>
    <p>Date: {{ order.order_date }}</p>

    <!-- Child component (like subform) -->
    <OrderItems
      :items="orderItems"
      :order-id="order.id"
      @update="handleItemsUpdate"
    />
  </div>
</template>
```

**Child Component (OrderItems.vue):**
```vue
<script setup>
import { ref } from 'vue'

const props = defineProps({
  items: Array,
  orderId: [String, Number]
})

const emit = defineEmits(['update'])

const localItems = ref([...props.items])

const addItem = () => {
  localItems.value.push({
    product_id: null,
    quantity: 1,
    price: 0
  })
}

const removeItem = (index) => {
  localItems.value.splice(index, 1)
  emit('update', localItems.value)
}

const saveItems = async () => {
  await fetch(`/api/orders/${props.orderId}/items`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(localItems.value)
  })
  emit('update', localItems.value)
}
</script>

<template>
  <div class="order-items">
    <h3>Order Items</h3>

    <table>
      <thead>
        <tr>
          <th>Product</th>
          <th>Quantity</th>
          <th>Price</th>
          <th>Total</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>
        <tr v-for="(item, index) in localItems" :key="index">
          <td>
            <select v-model="item.product_id">
              <!-- Product options -->
            </select>
          </td>
          <td>
            <input v-model.number="item.quantity" type="number" min="1" />
          </td>
          <td>
            <input v-model.number="item.price" type="number" step="0.01" />
          </td>
          <td>{{ (item.quantity * item.price).toFixed(2) }}</td>
          <td>
            <button @click="removeItem(index)">Remove</button>
          </td>
        </tr>
      </tbody>
    </table>

    <button @click="addItem">Add Item</button>
    <button @click="saveItems">Save</button>
  </div>
</template>
```

## Reports & Data Display

### Datasheet View → Data Table

**Access:**
- Datasheet view of form/query
- Built-in sorting/filtering

**Vue (using a library like vue3-easy-data-table):**
```bash
npm install vue3-easy-data-table
```

```vue
<script setup>
import { ref, onMounted } from 'vue'
import Vue3EasyDataTable from 'vue3-easy-data-table'
import 'vue3-easy-data-table/dist/style.css'

const employees = ref([])

const headers = [
  { text: 'ID', value: 'id', sortable: true },
  { text: 'Name', value: 'full_name', sortable: true },
  { text: 'Email', value: 'email', sortable: true },
  { text: 'Department', value: 'department.name', sortable: true },
  { text: 'Actions', value: 'actions' }
]

onMounted(async () => {
  const response = await fetch('/api/employees')
  employees.value = await response.json()
})

const viewEmployee = (employee) => {
  router.push({ name: 'employees.show', params: { id: employee.id } })
}
</script>

<template>
  <Vue3EasyDataTable
    :headers="headers"
    :items="employees"
    :search-value="searchValue"
    border-cell
    buttons-pagination
  >
    <template #item-actions="{ id }">
      <button @click="viewEmployee(id)">View</button>
      <button @click="editEmployee(id)">Edit</button>
    </template>
  </Vue3EasyDataTable>
</template>
```

### Export to Excel/PDF

```vue
<script setup>
import { exportToExcel, exportToPDF } from '@/utils/export'

const exportData = async (format) => {
  const response = await fetch('/api/employees/export')
  const data = await response.json()

  if (format === 'excel') {
    exportToExcel(data, 'employees.xlsx')
  } else if (format === 'pdf') {
    exportToPDF(data, 'employees.pdf')
  }
}
</script>

<template>
  <button @click="exportData('excel')">Export to Excel</button>
  <button @click="exportData('pdf')">Export to PDF</button>
</template>
```

## Best Practices

### 1. Component Organization

```
src/
├── components/
│   ├── common/           # Reusable components
│   │   ├── BaseInput.vue
│   │   ├── BaseSelect.vue
│   │   └── BaseButton.vue
│   ├── employees/        # Feature-specific
│   │   ├── EmployeeForm.vue
│   │   ├── EmployeeList.vue
│   │   └── EmployeeCard.vue
├── composables/          # Reusable logic
│   ├── useApi.js
│   ├── useForm.js
│   └── useToast.js
├── views/                # Page components
│   ├── EmployeesView.vue
│   └── DashboardView.vue
└── stores/               # State management
    └── employeeStore.js
```

### 2. Composables for Reusable Logic

```javascript
// composables/useApi.js
import { ref } from 'vue'

export function useApi(url) {
  const data = ref(null)
  const loading = ref(false)
  const error = ref(null)

  const fetchData = async () => {
    loading.value = true
    error.value = null

    try {
      const response = await fetch(url)
      if (!response.ok) throw new Error('Failed to fetch')
      data.value = await response.json()
    } catch (e) {
      error.value = e.message
    } finally {
      loading.value = false
    }
  }

  return { data, loading, error, fetchData }
}

// Usage in component
import { useApi } from '@/composables/useApi'

const { data: employees, loading, fetchData } = useApi('/api/employees')
onMounted(() => fetchData())
```

### 3. State Management with Pinia

```bash
npm install pinia
```

```javascript
// stores/employeeStore.js
import { defineStore } from 'pinia'

export const useEmployeeStore = defineStore('employee', {
  state: () => ({
    employees: [],
    currentEmployee: null,
    loading: false
  }),

  actions: {
    async fetchEmployees() {
      this.loading = true
      const response = await fetch('/api/employees')
      this.employees = await response.json()
      this.loading = false
    },

    async saveEmployee(employee) {
      const response = await fetch('/api/employees', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(employee)
      })
      const saved = await response.json()
      this.employees.push(saved)
      return saved
    }
  }
})

// Usage in component
import { useEmployeeStore } from '@/stores/employeeStore'

const employeeStore = useEmployeeStore()
await employeeStore.fetchEmployees()
```

### 4. Error Handling & User Feedback

```vue
<script setup>
import { ref } from 'vue'

const toast = ref({ show: false, message: '', type: 'success' })

const showToast = (message, type = 'success') => {
  toast.value = { show: true, message, type }
  setTimeout(() => {
    toast.value.show = false
  }, 3000)
}

const saveEmployee = async () => {
  try {
    await api.saveEmployee(employee.value)
    showToast('Employee saved successfully', 'success')
  } catch (error) {
    showToast('Error saving employee', 'error')
  }
}
</script>

<template>
  <div v-if="toast.show" :class="`toast toast-${toast.type}`">
    {{ toast.message }}
  </div>
</template>
```

## Migration Checklist

- [ ] List all Access forms
- [ ] Identify reusable components
- [ ] Create component hierarchy
- [ ] Set up Vue Router
- [ ] Implement form validation
- [ ] Create data table components
- [ ] Handle file uploads
- [ ] Implement search/filter functionality
- [ ] Add loading states
- [ ] Add error handling
- [ ] Implement responsive design
- [ ] Add accessibility features
- [ ] Test on different browsers

---

**See Also:**
- [Forms and UI Reference](./04-FORMS-UI.md)
- [Laravel Migration Guide](./05-LARAVEL-MIGRATION.md)
- [Translation Dictionary](./07-TRANSLATION-DICTIONARY.md)
