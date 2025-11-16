# Translation Dictionary: VBA/Access to Laravel/Vue

## Overview

This dictionary provides quick reference for translating MS Access/VBA concepts to their Laravel and Vue.js equivalents.

## Database Concepts

| Access/VBA | Laravel | Vue.js | Notes |
|------------|---------|--------|-------|
| Table | Eloquent Model | - | e.g., `tbl_users` → `User::class` |
| AutoNumber | `$table->id()` | - | Auto-incrementing primary key |
| Recordset | Collection | Array/Object | `Employee::all()` returns Collection |
| CurrentDb | `DB::` facade | - | Database access |
| Query | Query Builder / Eloquent | - | Fluent query interface |
| Relationship | Eloquent Relationships | - | `hasMany()`, `belongsTo()`, etc. |
| Lookup Field | Foreign Key + Relationship | - | `foreignId()` in migrations |
| Memo Field | `text()` | - | Long text storage |
| Yes/No Field | `boolean()` | Boolean | True/False values |
| Currency Field | `decimal(10, 2)` | Number | Monetary values |
| Date/Time | `datetime()` / Carbon | Date object | Date handling |
| Attachment | File path (string) | File upload | Store path, not binary |

## Data Access

| Access/VBA | Laravel | Notes |
|------------|---------|-------|
| `CurrentDb.OpenRecordset("SELECT...")` | `DB::select("SELECT...")` | Raw SQL query |
| `rs!FieldName` | `$model->field_name` | Field access |
| `rs.AddNew` | `Model::create([])` | Create new record |
| `rs.Update` | `$model->save()` | Save changes |
| `rs.Delete` | `$model->delete()` | Delete record |
| `rs.MoveFirst` | `$collection->first()` | First record |
| `rs.MoveLast` | `$collection->last()` | Last record |
| `rs.MoveNext` | Loop iteration | Next record |
| `rs.EOF` | Loop end condition | End of records |
| `rs.RecordCount` | `$collection->count()` | Count records |
| `DLookup()` | `Model::where()->value()` | Lookup single value |
| `DCount()` | `Model::where()->count()` | Count records |
| `DSum()` | `Model::where()->sum()` | Sum values |
| `DMax()` | `Model::where()->max()` | Maximum value |
| `DMin()` | `Model::where()->min()` | Minimum value |

## Form/UI Elements

| Access Control | Vue Component | Example |
|----------------|---------------|---------|
| TextBox | `<input type="text">` | `<input v-model="name">` |
| ComboBox (dropdown) | `<select>` | `<select v-model="selected">` |
| ListBox | `<select multiple>` | `<select v-model="items" multiple>` |
| CheckBox | `<input type="checkbox">` | `<input v-model="isActive" type="checkbox">` |
| OptionGroup | `<input type="radio">` | `<input v-model="choice" type="radio">` |
| Label | `<label>` or `<span>` | `<label>{{ text }}</label>` |
| Button | `<button>` | `<button @click="save">Save</button>` |
| SubForm | Child Component | `<ChildComponent :data="items" />` |
| Tab Control | Tab Component | Vue tabs library or custom |
| Image | `<img>` | `<img :src="imageUrl">` |

## Form Events

| Access Event | Vue.js Equivalent | Notes |
|--------------|-------------------|-------|
| `Form_Load` | `onMounted()` | Component initialization |
| `Form_Current` | `watch()` on data | Data change reaction |
| `Form_BeforeUpdate` | Form validation | Before save validation |
| `Form_AfterUpdate` | API callback | After save success |
| `Form_Close` | `onUnmounted()` | Component cleanup |
| `Click` | `@click` | Click handler |
| `DblClick` | `@dblclick` | Double-click handler |
| `Change` | `@change` or `v-model` | Value change |
| `AfterUpdate` | `@input` or `@change` | After value update |
| `GotFocus` | `@focus` | Focus event |
| `LostFocus` | `@blur` | Blur event |
| `KeyPress` | `@keypress` | Key press |
| `MouseMove` | `@mousemove` | Mouse movement |

## Navigation

| Access VBA | Laravel | Vue Router | Notes |
|------------|---------|------------|-------|
| `DoCmd.OpenForm "frmName"` | `return view('form')` | `router.push('/form')` | Navigate to form |
| `DoCmd.Close` | - | `router.back()` | Close/go back |
| `DoCmd.GoToRecord acNewRec` | - | `router.push('/create')` | New record |
| `DoCmd.GoToRecord acNext` | - | Navigate in UI | Next record |
| `DoCmd.GoToRecord acPrevious` | - | Navigate in UI | Previous record |
| `DoCmd.OpenReport` | Generate PDF | Open report view | View report |

## String Functions

| VBA Function | PHP (Laravel) | JavaScript (Vue) |
|--------------|---------------|------------------|
| `Len(str)` | `strlen($str)` | `str.length` |
| `Left(str, n)` | `substr($str, 0, $n)` | `str.substring(0, n)` |
| `Right(str, n)` | `substr($str, -$n)` | `str.slice(-n)` |
| `Mid(str, start, len)` | `substr($str, $start, $len)` | `str.substring(start, end)` |
| `UCase(str)` | `strtoupper($str)` | `str.toUpperCase()` |
| `LCase(str)` | `strtolower($str)` | `str.toLowerCase()` |
| `Trim(str)` | `trim($str)` | `str.trim()` |
| `LTrim(str)` | `ltrim($str)` | `str.trimStart()` |
| `RTrim(str)` | `rtrim($str)` | `str.trimEnd()` |
| `Replace(str, find, replace)` | `str_replace($find, $replace, $str)` | `str.replace(find, replace)` |
| `InStr(str, search)` | `strpos($str, $search)` | `str.indexOf(search)` |
| `Split(str, delimiter)` | `explode($delimiter, $str)` | `str.split(delimiter)` |
| `Join(array, delimiter)` | `implode($delimiter, $array)` | `array.join(delimiter)` |

## Date/Time Functions

| VBA Function | PHP (Laravel) | JavaScript (Vue) |
|--------------|---------------|------------------|
| `Now()` | `now()` | `new Date()` |
| `Date()` | `today()` | `new Date().toDateString()` |
| `Time()` | `now()->format('H:i:s')` | `new Date().toTimeString()` |
| `Year(date)` | `$date->year` | `date.getFullYear()` |
| `Month(date)` | `$date->month` | `date.getMonth() + 1` |
| `Day(date)` | `$date->day` | `date.getDate()` |
| `Hour(date)` | `$date->hour` | `date.getHours()` |
| `Minute(date)` | `$date->minute` | `date.getMinutes()` |
| `Second(date)` | `$date->second` | `date.getSeconds()` |
| `DateAdd("d", n, date)` | `$date->addDays($n)` | Use library (date-fns) |
| `DateDiff("d", date1, date2)` | `$date1->diffInDays($date2)` | Calculate manually |
| `Format(date, "mm/dd/yyyy")` | `$date->format('m/d/Y')` | `date.toLocaleDateString()` |

## Numeric Functions

| VBA Function | PHP (Laravel) | JavaScript (Vue) |
|--------------|---------------|------------------|
| `Abs(n)` | `abs($n)` | `Math.abs(n)` |
| `Round(n, decimals)` | `round($n, $decimals)` | `n.toFixed(decimals)` |
| `Int(n)` | `floor($n)` | `Math.floor(n)` |
| `Fix(n)` | `intval($n)` | `Math.trunc(n)` |
| `Sqr(n)` | `sqrt($n)` | `Math.sqrt(n)` |
| `Rnd()` | `rand(0, 1)` | `Math.random()` |
| `FormatNumber(n, decimals)` | `number_format($n, $decimals)` | `n.toLocaleString()` |
| `FormatCurrency(n)` | `number_format($n, 2)` | `n.toFixed(2)` |
| `FormatPercent(n)` | `($n * 100) . '%'` | `(n * 100) + '%'` |

## Conditional Functions

| VBA Function | PHP (Laravel) | JavaScript (Vue) |
|--------------|---------------|------------------|
| `IIf(condition, true, false)` | `$condition ? $true : $false` | `condition ? true : false` |
| `IsNull(value)` | `is_null($value)` | `value === null` |
| `Nz(value, default)` | `$value ?? $default` | `value ?? default` |
| `IsNumeric(value)` | `is_numeric($value)` | `!isNaN(value)` |
| `IsDate(value)` | Check with try/catch | `!isNaN(Date.parse(value))` |
| `IsEmpty(value)` | `empty($value)` | `!value` |

## Conversion Functions

| VBA Function | PHP (Laravel) | JavaScript (Vue) |
|--------------|---------------|------------------|
| `CStr(value)` | `(string) $value` | `String(value)` |
| `CInt(value)` | `(int) $value` | `parseInt(value)` |
| `CLng(value)` | `(int) $value` | `parseInt(value)` |
| `CDbl(value)` | `(float) $value` | `parseFloat(value)` |
| `CBool(value)` | `(bool) $value` | `Boolean(value)` |
| `CDate(value)` | `Carbon::parse($value)` | `new Date(value)` |
| `Val(string)` | `floatval($string)` | `parseFloat(string)` |

## Control Structures

### If Statement

**VBA:**
```vba
If condition Then
    ' code
ElseIf anotherCondition Then
    ' code
Else
    ' code
End If
```

**PHP:**
```php
if ($condition) {
    // code
} elseif ($anotherCondition) {
    // code
} else {
    // code
}
```

**JavaScript:**
```javascript
if (condition) {
  // code
} else if (anotherCondition) {
  // code
} else {
  // code
}
```

### Select Case / Switch

**VBA:**
```vba
Select Case variable
    Case "value1"
        ' code
    Case "value2"
        ' code
    Case Else
        ' code
End Select
```

**PHP:**
```php
switch ($variable) {
    case 'value1':
        // code
        break;
    case 'value2':
        // code
        break;
    default:
        // code
}

// Or PHP 8 match
match($variable) {
    'value1' => /* expression */,
    'value2' => /* expression */,
    default => /* expression */
}
```

**JavaScript:**
```javascript
switch (variable) {
  case 'value1':
    // code
    break
  case 'value2':
    // code
    break
  default:
    // code
}
```

### Loops

**VBA For Loop:**
```vba
For i = 1 To 10
    ' code
Next i
```

**PHP:**
```php
for ($i = 1; $i <= 10; $i++) {
    // code
}
```

**JavaScript:**
```javascript
for (let i = 1; i <= 10; i++) {
  // code
}
```

**VBA For Each:**
```vba
For Each item In collection
    ' code
Next item
```

**PHP:**
```php
foreach ($collection as $item) {
    // code
}
```

**JavaScript:**
```javascript
collection.forEach(item => {
  // code
})

// or
for (const item of collection) {
  // code
}
```

**VBA While Loop:**
```vba
Do While condition
    ' code
Loop
```

**PHP/JavaScript:**
```php
while ($condition) {
    // code
}
```

## Error Handling

**VBA:**
```vba
On Error GoTo ErrorHandler
' code
Exit Sub

ErrorHandler:
    MsgBox Err.Description
End Sub
```

**Laravel:**
```php
try {
    // code
} catch (\Exception $e) {
    Log::error($e->getMessage());
    // handle error
}
```

**Vue:**
```javascript
try {
  // code
} catch (error) {
  console.error(error.message)
  // handle error
}
```

## User Interaction

| VBA | Laravel | Vue.js |
|-----|---------|--------|
| `MsgBox "message"` | Flash message / toast | Toast notification |
| `InputBox("prompt")` | Form input | Modal with input |
| `MsgBox "text", vbYesNo` | Confirmation dialog | Confirm modal |

**Vue Example (using a toast library):**
```javascript
// Instead of MsgBox
toast.success('Operation completed')
toast.error('An error occurred')

// Instead of InputBox
const result = await showModal({
  title: 'Enter name',
  type: 'input'
})
```

## Authentication

| Access | Laravel |
|--------|---------|
| User-level security | Laravel Auth / Sanctum |
| Workgroup file | Roles & Permissions |
| CurrentUser() | `auth()->user()` |
| User groups | Roles (Spatie Laravel Permission) |

## Common Patterns

### Pattern: Load Dropdown from Table

**VBA:**
```vba
Me.cboCategory.RowSource = "SELECT CategoryID, CategoryName FROM tblCategories"
```

**Laravel API:**
```php
public function getCategories() {
    return Category::all(['id', 'name']);
}
```

**Vue:**
```vue
<script setup>
const categories = ref([])
onMounted(async () => {
  categories.value = await fetch('/api/categories').then(r => r.json())
})
</script>
<template>
  <select v-model="selectedCategory">
    <option v-for="cat in categories" :key="cat.id" :value="cat.id">
      {{ cat.name }}
    </option>
  </select>
</template>
```

### Pattern: Calculate Field on Change

**VBA:**
```vba
Private Sub txtQuantity_AfterUpdate()
    Me.txtTotal = Me.txtQuantity * Me.txtPrice
End Sub
```

**Vue:**
```vue
<script setup>
const quantity = ref(0)
const price = ref(0)
const total = computed(() => quantity.value * price.value)
</script>
<template>
  <input v-model.number="quantity" type="number">
  <input v-model.number="price" type="number">
  <p>Total: {{ total }}</p>
</template>
```

### Pattern: Master-Detail Relationship

**VBA:**
```vba
Me.subfrmDetails.LinkMasterFields = "OrderID"
Me.subfrmDetails.LinkChildFields = "OrderID"
```

**Laravel:**
```php
$order = Order::with('details')->find($id);
```

**Vue:**
```vue
<script setup>
const order = ref(null)
const orderDetails = computed(() => order.value?.details ?? [])

onMounted(async () => {
  order.value = await fetch(`/api/orders/${orderId}`).then(r => r.json())
})
</script>
<template>
  <OrderDetails :items="orderDetails" />
</template>
```

## SQL Translation

| Access SQL | Standard SQL (Laravel) | Notes |
|------------|------------------------|-------|
| `Yes/No` field = True | `column = 1` or `= true` | Boolean handling |
| `*` wildcard | `%` wildcard | In LIKE clauses |
| `#date#` | `'date'` | Date literals |
| `&` concatenation | `CONCAT()` | String concatenation |
| `[Field Name]` | `` `field_name` `` or `"field_name"` | Quoted identifiers |
| `UPDATE ... INNER JOIN` | Use subquery | Access allows JOIN in UPDATE |

## File Operations

| VBA | Laravel | Notes |
|-----|---------|-------|
| Attachment field | Storage facade | File storage |
| OLE Object | Binary field (avoid) | Store files separately |
| Export to Excel | Laravel Excel | `maatwebsite/excel` package |
| Export to PDF | DomPDF | `barryvdh/laravel-dompdf` |

## Migration Priority Guide

### High Priority (Core Functionality)
1. Database tables and relationships
2. CRUD operations
3. User authentication
4. Core business logic

### Medium Priority (Features)
1. Reports and exports
2. Search and filtering
3. Data validation
4. File uploads

### Low Priority (Nice to Have)
1. Advanced formatting
2. Complex reports
3. Non-essential features
4. UI polish

## Quick Reference Card

**Most Common Translations:**

```
VBA Recordset → Laravel Collection
Forms → Vue Components
DoCmd → Vue Router
DLookup → Model::where()->value()
MsgBox → Toast notification
CurrentDb → DB:: facade
Form_Load → onMounted()
AfterUpdate → @change or API callback
```

---

**See Also:**
- [VBA Fundamentals](./01-VBA-FUNDAMENTALS.md)
- [Laravel Migration Guide](./05-LARAVEL-MIGRATION.md)
- [Vue UI Migration Guide](./06-VUE-UI-MIGRATION.md)
