# Workbook
Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc. It can be used to list related references. 

## Properties

None

## Relationships
The Workbook has the following relationships defined:

| Relationship    | Type    |Description|Notes |
|:----------------|:--------|:----------|:-----|
| application  | [Application](application.md)| Returns an object that represents the Excel application managing the workbook. |
| names       | [NamedItem collection](nameditemCollection.md)| Collection of Named Ranges associated with the workbook  |Workbook.Names      |
| tables       | [Table collection](tableCollection.md)        | Collection of Tables associated with the workbook        |Workbook.ListObjects|
| worksheets   | [Worksheet collection](worksheetCollection.md)| Collection of Worksheets associated with the workbook    |Workbook.Worksheets |
| bindings   | [Binding collection](bindingCollection.md)| A collection of all the Binding objects that are part of the workbook.    | |

## Methods

The Worksheet has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[getActiveWorksheet()](#getactiveworksheet)| [Worksheet](worksheet.md) object |Get the currently active worksheet in the workbook.| |
|[getSelectedRange()](#getselectedrange)| [Range](range.md) object |Get the currently selected Range from the Workbook. | |  

## API Specification 

### getActiveWorksheet()

Get the currently active worksheet in the workbook.

#### Syntax
```js
context.workbook.getActiveWorksheet();
```
#### Parameters

None

#### Returns

[Worksheet](worksheet.md) object.

#### Examples 

```js
var ctx = new Excel.ExcelClientContext();
var activeWorksheet = ctx.workbook.getActiveWorksheet();
ctx.load(activeWorksheet);
ctx.executeAsync().then(function () {
		Console.log(activeWorksheet.name);
});
```
[Back](#methods)


### getSelectedRange()

Get the currently selected Range from the Workbook. 

#### Syntax
```js
context.workbook.getSelectedRange();
```
#### Parameters
None

#### Returns

[Range](range.md) object.

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var selectedRange = ctx.workbook.getSelectedRange();
ctx.executeAsync().then(function () {
		Console.log(selectedRange.address);
});
```
[Back](#methods)
