# Worksheet
The Worksheet object is a member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.

## [Properties](#get-worksheet)

| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|`index`| Number |The zero-based index of the worksheet within the workbook|Worksheet.Index|
|`name` | String |The user-visible name of the worksheet|Worksheet.Name    |


## Relationships
The Worksheet resource has the following relationships defined:

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|         
|charts | [Chart collection](chartCollection.md) |Collection of charts in the worksheet|Worksheet.ChartObject  |       
|tables | [Table collection](tableCollection.md) |Collection of Tables in the worksheet|Worksheet.ListObjects  |       

## Methods

The Worksheet resource has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[activate()][activate-link]| void     |Activates the worksheet |   |
|[deleteObject()][deleteobject-link]| void     |Deletes the worksheet ||
|[getCell(row: number, column: number)][getcell-link]| [Range](range.md) object |Returns a range containing the single cell specified by the zero-indexed row and column numbers          
|[getEntireWorksheetRange()][getentireworksheetrange-link]| [Range](range.md) object |Returns the range containing all cells in the worksheet| |
|[getRange(address: string)][getrange-link]| [Range](range.md) object |Returns the range specified by the address| |
|[getUsedRange()][getusedrange-link]| [Range](range.md) object |Returns the used range of the worksheet| |  

## API Specification 

### activate()

Make the worksheeet active in the Excel UI.

#### Syntax

```js
worksheetObject.activate();
```
#### Parameters
None

#### Returns

Nothing

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var wSheetName = 'Sheet1';
var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
worksheet.activate();
ctx.executeAsync().then();
```
[Back](#methods)

### deleteObject()

Delete a worksheet from the workbook. 

#### Syntax
```js
worksheetObject.deleteObject();
```
#### Parameters
None

#### Returns

Nothing

#### Examples

```js
var wSheetName = 'Sheet1';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
worksheet.deleteObject();
ctx.executeAsync().then();
```
[Back](#methods)


### getCell(row: number, column: number)
Get the Cell (as a Range object) object based on row and column address relative to a top of worksheet. 

#### Syntax

```js
worksheetObject.getCell(row, column);
```

#### Parameters 

Parameter      | Type   | Description
-------------- | ------ | ------------
`row`          | Number | Required. Row number of the cell to be retrieved. Zero indexed. 
`col`          | Number | Required. Column number of the cell to be retrieved. Zero indexed.

#### Returns

[Range](range.md) object.

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var cell = worksheet.getCell(0,0);
ctx.load(cell);
ctx.executeAsync().then(function() {
	Console.log(cell.address);
});
```
[Back](#methods)

### getEntireWorksheetRange()

Get the entire Range associated with the worksheet.

#### Syntax
```js
worksheetObject.getEntireSheetRange();
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. Worksheet name. 


#### Returns

[Range](range.md) object.


**Note: the grid properties of the Range (values, numberFormat, formula) contains `null` since the Range in question is unbounded.**

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var wSheetName = 'Sheet1';
var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
var entireRange = worksheet.getEntireSheetRange();
ctx.load(entireRange);
ctx.executeAsync().then(function () {
		Console.log(entireRange.address);
});
```
[Back](#methods)

### getRange(address: string)

Get a Range object that represents a single cell or a range of cells. 

#### Syntax

```js
worksheetObject.getRange(address);
```
#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `address`| String | Required. Address of the Range. 

#### Returns

[Range](range.md) object.

#### Examples

Below example uses range address to get the range object.

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.cellCount);
});
```

Below example uses a named-range to get the range object.
```js
var rangeName = 'MyRange';
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.names.getItem(rangeName).range;
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.cellCount);
});
```
[Back](#methods)

### getUsedRange()

Get the used-range of a worksheet. 

#### Syntax
```js
worksheetObject.getUsedRange();
```
#### Parameters

None

#### Returns

[Range](r.md) object.

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var wSheetName = 'Sheet1';
var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
var usedRange = worksheet.getUsedRange();
ctx.load(usedRange);
ctx.executeAsync().then(function () {
		Console.log(usedRange.address);
});
```
[Back](#methods)

### Get Worksheet

Get Worksheet object properties based on name.

#### Syntax
```js
worksheetCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. Worksheet name. 

#### Returns

[Worksheet](worksheet.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var wSheetName = 'Sheet1';
var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
ctx.executeAsync().then(function () {
		Console.log(worksheet.index);
});
```
[Back](#properties)



[activate-link]: #activate
[deleteobject-link]: #deleteobject
[getcell-link]: #getcellrow-number-column-number
[getentireworksheetrange-link]: #getentireworksheetrange
[getrange-link]: #getrangeaddress-string
[getusedrange-link]: #getusedrange
