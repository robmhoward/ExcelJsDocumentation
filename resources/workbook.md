# Workbook
Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc. It can be used to list related references. 

## Properties

None

## Relationships
The Workbook resource has the following relationships defined:

| Relationship    | Type    |Description|Notes |
|:----------------|:--------|:----------|:-----|
| [application](#get-application)    | [Application](application.md)| Returns an object that represents epresents the Excel application which is managing the workbook. |
| [names](#list-names)         | [NamedItem](nameditem.md) collection| Collection of Named Ranges associated with the workbook  |Workbook.Names      |
| [tables](#list-tables)        | [Table](table.md) collection        | Collection of Tables associated with the workbook        |Workbook.ListObjects|
| [worksheets](#list-worksheets)    | [Worksheet](worksheet.md) collection| Collection of Worksheets associated with the workbook    |Workbook.Worksheets |

## Methods

The Worksheet resource has the following methods defined:

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
[Back](#workbook)


### Get Application

Get properties of workbook's application object. 

```js
context.workbook.application;
```
#### Returns

[Application](application.md) object.

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var application = ctx.workbook.application;
ctx.load(application);
ctx.executeAsync().then(function() {
	Console.log(application.calculationMode);
});

```
[Back](#relationships)

### List Names

Get Names collection that contains each of the Name objects contained in the Workbook. Each item contains the following properties. 
** Note: This API currently supports only the Workbook scoped items. **
#### Syntax
```js
context.workbook.tables;
```
#### Returns

[Named-Item](nameditem.md) collection.


#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var names = ctx.workbook.names;
ctx.load(names);
ctx.executeAsync().then(function () {
	Console.log("Names: Count= " + names.count);
	for (var i = 0; i < names.items.length; i++)
	{
		Console.log(names.items[i].name);
	}
});
```
[Back](#relationships)

### List Tables

Get Table collection contained in workbook. Each item contains the following properties. 

#### Syntax
```js
context.workbook.tables;
```
#### Returns

[Table](table.md) collection.

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
ctx.load(tables);
ctx.executeAsync().then(function () {
	for (var i = 0; i < tables.items.length; i++)
	{
		Console.log(tables.items[i].name);
	}
});
```
##### Getting the number of tables

```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("Tables: Count= " + tables.count);
});

```
[Back](#relationships)

### List Worksheets

The Worksheet collection contains each of the worksheets defined as part of the workbook. Note: This does not contain chart sheets.

#### Syntax
```js
context.workbook.worksheets;
```
#### Returns

[Worksheet](worksheet.md) collection. 


#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var worksheets = ctx.workbook.worksheets;
ctx.load(worksheets);
ctx.executeAsync().then(function () {
	for (var i = 0; i < worksheets.items.length; i++)
	{
		Console.log(worksheets.items[i].name);
		Console.log(worksheets.items[i].index);
	}
});
```

##### Getting the number of worksheets

```js
var ctx = new Excel.ExcelClientContext();
var worksheets = ctx.workbook.worksheets;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("Worksheets: Count= " + worksheets.count);
});

```
[Back](#relationships)