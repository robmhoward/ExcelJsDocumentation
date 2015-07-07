# Workbook

Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc.

## Properties
None

## Relationships
| Relationship | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|application|[Application](application.md)|Represents Excel application instance that contains this workbook. Read-only.||
|bindings|[BindingCollection](bindingcollection.md)|Represents a collection of bindings that are part of the workbook. Read-only.||
|names|[NamedItemCollection](nameditemcollection.md)|Represents a collection of workbook scoped named items (named ranges and constants). Read-only.||
|tables|[TableCollection](tablecollection.md)|Represents a collection of tables associated with the workbook. Read-only.||
|worksheets|[WorksheetCollection](worksheetcollection.md)|Represents a collection of worksheets associated with the workbook. Read-only.||

## Methods

| Method           | Return Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|[getSelectedRange()](#getselectedrange)|[Range](range.md)|Gets the currently selected range from the workbook.||

## API Specification

### getSelectedRange()
Gets the currently selected range from the workbook.

#### Syntax
```js
workbookObject.getSelectedRange();
```

#### Parameters
None

#### Returns
[Range](range.md)

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
ctx.executeAsync().then(function () {
		Console.log(selectedRange.address);
});
```
[Back](#methods)

