# Revit API Patterns for VOR Validation Scripts

Common patterns for working with the Revit API in IronPython 2.7 validation scripts.

## Element Collection

### By Category

```python
# Collect all instances of a category
walls = (DB.FilteredElementCollector(doc)
         .OfCategory(DB.BuiltInCategory.OST_Walls)
         .WhereElementIsNotElementType()
         .ToElements())
```

Common categories:
- `OST_Walls` — walls
- `OST_Doors` — doors
- `OST_Windows` — windows
- `OST_Floors` — floors
- `OST_Roofs` — roofs
- `OST_Columns` — columns
- `OST_Stairs` — stairs
- `OST_Ramps` — ramps
- `OST_Railing` — railings
- `OST_Furniture` — furniture
- `OST_Parking` — parking
- `OST_StructuralFraming` — beams/braces
- `OST_StructuralColumns` — structural columns
- `OST_StructuralFoundations` — foundations
- `OST_PipeCurves` — pipes
- `OST_DuctCurves` — ducts
- `OST_CableTray` — cable trays
- `OST_Conduit` — conduits

### By Class

```python
# Collect by .NET type
levels = (DB.FilteredElementCollector(doc)
          .OfClass(DB.Level)
          .ToElements())

viewports = (DB.FilteredElementCollector(doc)
             .OfClass(DB.Viewport)
             .ToElements())
```

### By Parameter Value (FilteredElementCollector)

```python
# Filter by a parameter value
param_prov = DB.ParameterValueProvider(
    DB.ElementId(DB.BuiltInParameter.WALL_BASE_OFFSET)
)
evaluator = DB.FilterNumericEquals()
rule = DB.FilterDoubleRule(param_prov, evaluator, 0.0, 0.001)
param_filter = DB.ElementParameterFilter(rule)

walls_at_zero = (DB.FilteredElementCollector(doc)
                 .OfCategory(DB.BuiltInCategory.OST_Walls)
                 .WhereElementIsNotElementType()
                 .WherePasses(param_filter)
                 .ToElements())
```

### Sheets

```python
sheets = (DB.FilteredElementCollector(doc)
          .OfClass(DB.ViewSheet)
          .WhereElementIsNotElementType()
          .ToElements())

for sheet in sheets:
    number = sheet.SheetNumber  # str, e.g. "A1"
    name = sheet.Name           # str, e.g. "Floor Plan"
```

### Views on a Sheet

```python
# Get all viewports placed on a sheet
viewports = sheet.GetAllViewports()  # list of ElementId
for vp_id in viewports:
    viewport = doc.GetElement(vp_id)
    view = doc.GetElement(viewport.ViewId)
    # view is a DB.View
```

## Parameter Access

### By BuiltInParameter (Preferred, Fast)

```python
# Built-in parameters are constants — no string matching needed
param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)
```

Common built-in parameters:
- `ALL_MODEL_INSTANCE_COMMENTS` — Comments (instance)
- `ALL_MODEL_INSTANCE_MARK` — Mark (instance)
- `WALL_BASE_OFFSET` — Base Offset
- `WALL_TOP_OFFSET` — Top Offset
- `WALL_BASE_HEIGHT` — Unconnected Height
- `WALL_STRUCTURAL_SIGNIFICANT` — Structural
- `DOOR_NUMBER` — Door Number (actually MARK)
- `ROOM_NUMBER` — Room Number
- `ROOM_NAME` — Room Name
- `ROOM_AREA` — Room Area (read-only)
- `ELEM_FAMILY_PARAM` — Family (read-only)
- `ELEM_TYPE_PARAM` — Type (read-only)
- `VIEW_SCALE` — View Scale
- `SHEET_NUMBER` — Sheet Number
- `SHEET_NAME` — Sheet Name
- `SHEET_ISSUE_DATE` — Issue Date
- `SHEET_CHECKED_BY` — Checked By
- `SHEET_DESIGNED_BY` — Designed By

### By Name (For Shared/Custom Parameters)

```python
def get_param_by_name(element, name):
    """Find parameter by name (case-sensitive)."""
    for p in element.Parameters:
        if p.Definition.Name == name:
            return p
    return None
```

### Read Parameter Values

```python
# String
val = param.AsString()

# Double (internal Revit units — feet for length)
val = param.AsDouble()

# Integer
val = param.AsInteger()

# ElementId (reference to another element)
val = param.AsElementId()

# Display string (formatted with units)
val = param.AsValueString()
```

### Write Parameter Values (Must Be in Transaction)

```python
# String
param.Set("new value")

# Double
param.Set(1.5)  # in internal units (feet)

# Integer
param.Set(42)
```

### Check Parameter Storage type

```python
if param.StorageType == DB.StorageType.String:
    val = param.AsString()
elif param.StorageType == DB.StorageType.Double:
    val = param.AsDouble()
elif param.StorageType == DB.StorageType.Integer:
    val = param.AsInteger()
elif param.StorageType == DB.StorageType.ElementId:
    val = param.AsElementId()
```

## Transactions

### Basic Pattern

```python
t = DB.Transaction(doc, "Script: description")
t.Start()
try:
    # ... modify elements ...
    t.Commit()
except Exception:
    t.RollBack()
    raise
```

### SubTransaction (Nested)

```python
t = DB.Transaction(doc, "Main")
t.Start()
try:
    # group of changes 1
    for elem in batch1:
        elem.get_Parameter(bip).Set("val")

    # optional sub-transaction for partial rollback
    st = DB.SubTransaction(doc)
    st.Start()
    try:
        # risky change
        st.Commit()
    except:
        st.RollBack()

    t.Commit()
except:
    t.RollBack()
```

### TransactionGroup (For Multiple Transactions)

```python
tg = DB.TransactionGroup(doc, "Batch operation")
tg.Start()
try:
    for item in items:
        t = DB.Transaction(doc, "Process item")
        t.Start()
        # ... modify ...
        t.Commit()
    tg.Assimilate()  # merges all into single undo
except:
    tg.RollBack()
```

## Type vs Instance Parameters

```python
# Instance parameter — on the element itself
inst_param = elem.get_Parameter(DB.BuiltInParameter.ALL_MODEL_INSTANCE_COMMENTS)

# Type parameter — on the element's type
elem_type = doc.GetElement(elem.GetTypeId())
type_param = elem_type.get_Parameter(DB.BuiltInParameter.ALL_MODEL_TYPE_COMMENTS)
```

## Geometry and Location

```python
# Get element location (point or curve)
loc = elem.Location
if isinstance(loc, DB.LocationPoint):
    point = loc.Point  # XYZ
elif isinstance(loc, DB.LocationCurve):
    curve = loc.Curve  # Curve (Line, Arc, etc.)
    start = curve.GetEndPoint(0)
    end = curve.GetEndPoint(1)
    length = curve.Length  # in feet
```

## Level and Reference

```python
# Get element's level
level_id = elem.LevelId
level = doc.GetElement(level_id)
level_name = level.Name

# Get all levels
levels = (DB.FilteredElementCollector(doc)
          .OfClass(DB.Level)
          .ToElements())
```

## Category Checking

```python
cat = elem.Category
if cat.Id.IntegerValue == int(DB.BuiltInCategory.OST_Walls):
    pass  # it's a wall

# Or check directly
if elem.Category.Name == "Walls":
    pass
```

## Error Handling Pattern for Scripts

```python
def run(doc, section, project, settings=None):
    try:
        # Phase 1: Collect (read-only, no transaction needed)
        elements = collect_elements(doc, settings)
        if not elements:
            return ValidationResult(check_name=SCRIPT_NAME, passed=True,
                                    message="No elements to check.")

        # Phase 2: Analyze (read-only)
        problems = analyze(elements)

        # Phase 3: Report (validation-only, no transaction)
        return ValidationResult(
            check_name=SCRIPT_NAME,
            passed=len(problems) == 0,
            message="Checked: {}. Problems: {}".format(len(elements), len(problems)),
            elements=[p.Id for p in problems]
        )

    except Exception as e:
        return ValidationResult(check_name=SCRIPT_NAME, passed=False,
                                message="Error: {}".format(str(e)))
```

## Unit Conversion

Revit internally uses feet for length. Common conversions:

```python
# mm to feet
MM_TO_FEET = 1.0 / 304.8

# feet to mm
feet_to_mm = 304.8

# Convert user-facing mm value to Revit internal
internal_value = user_mm * MM_TO_FEET

# Convert Revit internal to display mm
display_mm = internal_value * feet_to_mm
```
