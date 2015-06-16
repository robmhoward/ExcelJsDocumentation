# Worksheet
An Excel worksheet is a collection cells organized by rows and columns. It can contain data, tables, charts, etc. 

## [Properties](#get-worksheet)

| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|`id`   | String | A unique Id that identifies the Worksheet object in a Workbook. For a given worksheet, the Id remains constant through changes such as renames or moves.|        |
|`position`| Number |The zero-based position of the worksheet within the workbook.|Worksheet.Index|
|`name` | String |The display name of the worksheet. |Worksheet.Name    |


## Relationships

| Relationship     | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|         
|charts | [Chart collection](chartcollection.md)|Collection of charts that are part of the worksheet.|Worksheet.ChartObject| 
|tables | [Table collection](tablecollection.md)|Collection of tables that are part of the worksheet.|Worksheet.ListObjects|       

## Methods

The Worksheet has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[activate()][activate-link]| void       |Makes the worksheet the active in the Excel application. |   |
|[delete()][deleteobject-link]| void     |Deletes the worksheet and its associated data. ||
|[getCell(row: number, column: number)][getcell-link]| [Range](range.md) object |Returns a range object based on the the zero-indexed row and column numbers.||          
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

### delete()

Delete a worksheet from the workbook. 

#### Syntax
```js
worksheetObject.delete();
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
worksheet.delete();
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
`row`          | Number | Required. Row number of the cell to be retrieved. Zero-indexed. 
`column`          | Number | Required. Column number of the cell to be retrieved. Zero-indexed.

#### Returns

[Range](range.md) object.

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var cell = worksheet.getCell(0,0);
ctx.load(cell);
ctx.executeAsync().then(function() {
	Console.log(cell.address);
});
```
[Back](#methods)

### getRange(address: string)

Get a Range object that represents a single cell or a range of cells. This API can also be used to obtain the entire range object associated with the worksheet. 

#### Syntax

```js
worksheetObject.getRange(address);
```
#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `address`| String | Optional. Address or the name of the Range. If not specified, the entire worksheet range is returned. 

#### Returns

[Range](range.md) object.
**Note: If the entire worksheet range is returned, the grid properties of the Range (values, numberFormat, formula) will contain `null` since the Range in question is unbounded.**

#### Examples

Below example uses range address to get the range object.

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:F8";
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
var sheetName = "Sheet1";
var rangeName = 'MyRange';
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeName);
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.address);
});
```

Below example get the entire worksheeet range.
**Note: If the entire worksheet range is returned, the grid properties of the Range (values, numberFormat, formula) will contain `null` since the Range in question is unbounded.**

```js
var rangeName = 'MyRange';
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange();
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.address);
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
worksheetCollection.getItem(param);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `param`| String | Required. Worksheet name or id. 

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
[deleteobject-link]: #delete
[getcell-link]: #getcellrow-number-column-number
[getentireworksheetrange-link]: #getentireworksheetrange
[getrange-link]: #getrangeaddress-string
[getusedrange-link]: #getusedrange
