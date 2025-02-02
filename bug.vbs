Late Binding: VBScript's flexibility with late binding can lead to runtime errors if an object or method doesn't exist.  This is especially problematic when dealing with COM objects or external libraries where versioning inconsistencies can occur.

Example:

```vbscript
Set objExcel = CreateObject("Excel.Application")
' ... later in the code ...
Set objChart = objExcel.Charts(1).Shapes(1) 'Error if the chart or shape doesn't exist
```

Early binding (declaring object types explicitly) can prevent some of these issues, but it sacrifices VBScript's dynamic nature.

Implict Type Conversion: VBScript's loose typing can result in unexpected type coercion.  Operations on mismatched data types may lead to runtime errors or produce incorrect results.

Example:

```vbscript
Dim x, y
x = "10"
y = 20
Dim z = x + y 'Type mismatch error. 'x' is a string, 'y' is a number.
```