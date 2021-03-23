ExcelEnt 
=========

## Overview
**ExcelEnt** is a simple library to make working with excel easier. It helps bind excel files to entities and create files based on entities.

ExcelEnt can be useful for these scenarios:
- bind excel file to entities
- generate excel list file
- generate excel list file with columns titles
- generate excel list file based on excel template
- replace excel template shortcodes by value
- add styles for all columns, for specific column, for row by condition

Usage
------

## Binding
```csharp
class XLSXBinder<T>
```
Create binding rules.
```csharp
binder.AddRule(0, e => e.Int, BindMappers.Int);
```
Get specific excel rows.
```csharp
binder.Skip(1);
binder.Take(2);
```
Bind excel to entities
```csharp
var entities = binder.Bind("test.xlsx", 0);
```

#### Binding funcs
**BindMappers** class contains some basic mapping funcs:
- BindMappers.Int
- BindMappers.NullInt
- BindMappers.Double
- BindMappers.NullDouble
- BindMappers.Bool
- BindMappers.NullBool
- BindMappers.StringBool *(get bool from string by string true value)*
- BindMappers.Date
- BindMappers.NullDate
- BindMappers.Enum *(get enum by string description attr)*
- BindMappers.String

Custom mapper example: func that take excel cell and return new value.
```csharp
/// Func<ICell, object>
(cell) => cell.ToString().Split('#')[1]
```
#### Example
```csharp
/// Bind excel 0, 1, 3 column, skip 1 row and take next 10 rows.
var items = new XLSXBinder<TestItem>()
    .AddRule(0, e => e.Int, BindMappers.Int)
    .AddRule(1, e => e.NullDouble, BindMappers.NullDouble)
    .AddRule(3, e => e.Enum, BindMappers.Enum<TestEnum>)
    .Skip(1)
    .Take(10)
    .Bind("data.xlsx", 0);
```

## Generation
```csharp
class XLSXWriter<T>
```
Create write rules.
```csharp
generator.AddRule(e => e.Int, 0);
```
Add columns titles
```csharp
generator.AddColumnsTitles(new string[]
{
    "first", 
    "second"
});
```
Generate excel based on existing template.
```csharp
// start index for entities inserting
var insertFromInd = 3;
// move rows below inserting index on new entity add
var moveFooter = true
generator.UseTemplating(t => t
	.FromTemplate("excelTemplate.xlsx", insertFromInd, moveFooter)
	// replace shortcodes in current template
	.ReplaceShortCode("shortcode", "new value"));
```
Styling
```csharp
.UseStyling(s => s
	// add style to workbook
    .AddStyle(CreateRedStyle, "red")
    .AddStyle(CreateTableStyle, "table")
    .AddStyle(CreateBlueStyle, "blue")
    // set style for 0-index column cells
    .AddCellsStyle(blueBGStyle, 0)
    // set default cells style
    .AddCellsStyle(tableStyle)
    // add row condition style
    .AddConditionRowStyle((entity, ind) => e.NullInt == null ? redBGStyle : null))
```
Style initialization action for AddStyle method
```csharp
void CreateStyle(ICellStyle style)
{
    style.BorderBottom = BorderStyle.Thin;
    ...
}
```
Generate file
```csharp
XSSFWorkbook wb = generator.Generate(items);
// save workbook using extention method
// generator.Generate(items).SaveTo("newExcel.xlsx")
wb.SaveTo("newExcel.xlsx");
```
#### Examples
```csharp
// simple excel with only data columns
XSSFWorkbook wb = new XLSXWriter<TestItem>()
	.AddRule(e => e.Int, 0)
    .AddRule(e => e.String, 1)
    .Generate(items);
```
```csharp
// excel with columns titles
XSSFWorkbook wb = new XLSXWriter<TestItem>()
    .AddRule(e => e.Int, 0)
    .AddRule(e => e.NullInt, 1)
    .AddColumnsTitles(new string[]
    {
        "First", 
        "Second"
    })
    .Generate(items);
```
```csharp
// excel based on template with some styling
XSSFWorkbook wb = new XLSXWriter<TestItem>()
    .AddRule(e => e.Int, 0)
    .AddRule(e => e.NullInt, 1)
    .UseTemplating(t => t
        .FromTemplate(templatePath, 3, true)
        .ReplaceShortCode("header", "TempTest")
    .UseStyling(s => s
        .AddStyle(CreateRedStyle, redBGStyle)
        .AddStyle(CreateTableStyle, tableStyle)
        .AddCellsStyle(tableStyle)
        .AddConditionRowStyle((e, i) => e.NullInt == null ? redBGStyle : null))
    .Generate(items);
```
## Task list
- [x] Add excel binder
- [x] Add simple generator
- [x] Add title cols generator
- [x] Add template based generator
- [x] Add cells styling
- [x] Add condition row styling
- [ ] Add default styles collection
