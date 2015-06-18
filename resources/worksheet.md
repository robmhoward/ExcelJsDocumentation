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
|[activate()][activate-link]| void       | Activate the worksheet in the Excel UI. |   |
|[delete()][deleteobject-link]| void     |Deletes the worksheet from the workbook. ||
|[getCell(row: number, column: number)][getcell-link]| [Range](range.md) object |Gets the range object containing the single cell specified by the zero-indexed row and column numbers. |          
|[getRange(address: string)][getrange-link]| [Range](range.md) object |Gets the range object specified by the address or name. | |
|[getRow(address: string)][getrange-link]| [Range](range.md) object |Gets the range object specified by the address or name. | |
|[getUsedRange()][getusedrange-link]| [Range](range.md) object |Gets the used range of the worksheet.| |  

GetContainedRange

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
Gets the range object containing the single cell specified by the zero-indexed row and column numbers. 

#### Syntax

```js
worksheetObject.getCell(row, column);
```

#### Parameters 

Parameter      | Type   | Description
-------------- | ------ | ------------
`row`          | Number | Required. The row number of the cell to be retrieved. Zero-indexed. 
`column`          | Number | Required. the column number of the cell to be retrieved. Zero-indexed.

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

Gets the range object specified by the address or name.

#### Syntax

```js
worksheetObject.getRange(address);
```
#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `address`| String | Optional. The address or the name of the range. If not specified, the entire worksheet range is returned. 

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

Gets the used range of the worksheet.

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

Gets a Worksheet object based on name.

#### Syntax
```js
worksheetCollection.getItem(param);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `param`| String | Required. The name or id of the worksheet. 

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
