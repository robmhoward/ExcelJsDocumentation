# Excel JavaScript APIs

## Objects 

* [Workbook](#workbook): Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc. It can be used to list related references. 
* [Worksheet](#worksheet): The Worksheet object is a member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.
* [Range](#range): Range represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells.  
* [Table](#table): Represents collection of organized cells designed to make management of the data easy. 
* [Chart](#chart): Represents a chart object in a workbook, which is a visual representation of underlying data.   
* [Named-Item](#named-item): Represents a defined name for a range of cells or a value. Names can be primitive named objects (as seen in the type below), range object, etc.

Also see: 

* [Error Messages](#error-messages): Provide important programming details related to Excel APIs.
* [Programming Notes](#programming-notes): Provide important programming details related to Excel APIs.

## Workbook
The [Workbook](resources/workbook.md) is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc. It can be used to list related references. 

Following are some of the methods supported for this resource:

| Task                             | Description                                |
|:----------------------------------|:-------------------------------------------|
| [Get-Application](#get-application) | Get metadata properties of a Excel Application managing the Workbook|    
| [List-Worksheets](#list-worksheets)         | Retrieve collection worksheets that are part of the workbook     |
| [List-Tables](#list-tables)             | Retrieve collection Tables that are part of the workbook         |
| [List-Names](#list-names)              | Retrieve collection Name Objects that are part of the workbook   |
| [Get-Selected-Range](#get-selected-range)              | Retrieve Range object that is currently selected    |
| [Calculate](#calculate) | Performs calculation on the Workbook or Application.   | 


### Get-Application

Get properties of workbook's application object. 

```js
context.workbook.application;
```
#### Returns

[Application](resources/application.md) object.

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var application = ctx.workbook.application;
ctx.load(application);
ctx.executeAsync().then(function() {
	Console.log(application.calculationMode);
});

```
[Back](#workbook)

### List-Worksheets

The Worksheet collection contains each of the worksheets defined as part of the workbook. Note: This does not contain chart sheets.

#### Syntax
```js
context.workbook.worksheets;
```
#### Returns

[Worksheet](resources/worksheet.md) collection. 


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

##### Getting the number of tables

```js
var ctx = new Excel.ExcelClientContext();
var worksheets = ctx.workbook.worksheets;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("Worksheets: Count= " + worksheets.count);
});

```
[Back](#workbook)

### List-Tables

Get Table collection contained in workbook. Each item contains the following properties. 

#### Syntax
```js
context.workbook.tables;
```
#### Returns

[Table](resources/table.md) collection.

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
[Back](#workbook)

### List-Names

Get Names collection that contains each of the Name objects contained in the Workbook. Each item contains the following properties. 
** Note: This API currently supports only the Workbook scoped items. **
#### Syntax
```js
context.workbook.tables;
```
#### Returns

[Named-Item](resources/nameditem.md) collection.


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
[Back](#workbook)

### Get-Selected-Range

Get the currently selected Range from the Workbook. 

#### Syntax
```js
context.workbook.getSelectedRange();
```
#### Parameters
None

#### Returns

[Range](resources/range.md) object.

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var selectedRange = ctx.workbook.getSelectedRange();
ctx.executeAsync().then(function () {
		Console.log(selectedRange.address);
});
```
[Back](#workbook)

### Calculate

Performs calculation on the Workbook or Application. 

#### Syntax
```js
context.workbook.calculate(calculationType)
```
#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
`calculationType` | String | Optional. `ReCalculate`: does normal calculation, `Full`: forces a full calculation of the data, `FullRebuild`: forces a full calculation of the data and rebuilds the dependencies (this is similar to re-entering all formulas). Note: if request body is not provided then calculation of the type `ReCalculation` is performed.

#### Returns

Nothing

#### Examples 

```js
var ctx = new Excel.ExcelClientContext();
ctx.workbook.application.calculate('Full');
ctx.executeAsync().then();
```
[Back](#workbook)

### Named-Item
[Named-Item](resources/nameditem.md) represents a defined named object. Names can be primitive named objects or reference to a range. This can be used to obtain Range object associated with names.

Following are the methods supported for this resource:

| Task                               | Description                                |
|:------------------------------------|:-------------------------------------------|
| [Get-Named-Item](#get-named-item)   | Retrieve a Named Object                                  |

### Get-Named-Item

Get a Named object. 

** Note: This API currently supports only the Workbook scoped items. **
#### Syntax
```js
namesCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. Name of the item.

#### Returns

[Named-Item](resources/nameditem.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var names = ctx.workbook.names;
var namedItem = names.getItem('MyRange');
ctx.load(namedItem);
ctx.executeAsync().then(function () {
		Console.log(namedItem.type);
});
```
[Back](#named-item)


## Worksheet 
The [Worksheet](resources/worksheet.md) object is a member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.

Following are the methods supported for this resource:

| Task                               | Description                                |  
|:------------------------------------|:-------------------------------------------| 
| [Get-Worksheet](#get-worksheet)     | Retrieve properties of a specific Worksheet                  |
| [Get-Used-Range](#get-used-range)   | Retrieve used-range of a specific worksheet                  |
| [Get-Entire-Worksheet-Range](#get-entire-worksheet-range) | Returns a Range object that represents the entire worksheet.  |
| [Add-Worksheet](#add-worksheet)     | Add a new Worksheet to the Workbook                          |
| [Delete-Worksheet](#delete-worksheet)            | Delete Worksheet from the Workbook              |
| [List-Charts](#list-charts)                      | Retrieve list of Charts in a Worksheet          |  
| [Get-Active-Worksheet](#get-active-worksheet)    | Get Worksheet that is currently active workbook |
| [Get-Cell](#get-cell)                            | Get Cell properties |


### Get-Worksheet

Get Worksheet object properties based on name.

#### Syntax
```js
worksheetsCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. Worksheet name. 

#### Returns

[Worksheet](resources/worksheet.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var wSheetName = 'Sheet1';
var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
ctx.executeAsync().then(function () {
		Console.log(worksheet.index);
});
```
[Back](#worksheet)

### Get-Used-Range

Get the used-range of a worksheet. 

#### Syntax
```js
worksheetObject.getUsedRange();
```
#### Parameters

None

#### Returns

[Range](resources/r.md) object.


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
[Back](#worksheet)

### Get-Entire-Worksheet-Range

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

[Range](resources/range.md) object.


** Note: the grid properties of the Range (values, numberFormat, formula) contains `null` since the Range in question is unbounded. **

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
[Back](#worksheet)


### Add-Worksheet

Add a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets.

#### Syntax
```js
worksheetsCollection.add(name);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`name`  | String| Optional. String value representing the name of the sheet to be added. If not specified, Excel determines the name of the new worksheet being added. 

#### Returns
[Worksheet](resources/worksheet.md) object.

#### Examples

```js
var wSheetName = 'Sample Name';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.add(wSheetName);
ctx.load(worksheet);
ctx.executeAsync().then(function () {
	Console.log(worksheet.name);
});
```
[Back](#worksheet)


### Delete-Worksheet

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
[Back](#worksheet)


### List-Charts 

Get Charts collection that contains each of the chart objects contained in the worksheet. Each item contains the following properties. 

#### Syntax
```js
worksheetObject.charts;
```

#### Returns

[Chart](resources/chart.md) collection.

#### Examples

```js
var wSheetName = 'Sheet1';
var ctx = new Excel.ExcelClientContext();
var charts = ctx.workbook.worksheets.getItem(wSheetName).charts;
ctx.load(charts);
ctx.executeAsync().then(function () {
	for (var i = 0; i < charts.items.length; i++)
	{
		Console.log(charts.items[i].name);
	}
});
```
[Back](#worksheet)


### Get-Active-Worksheet

Get the currently active worksheet in the workbook.

#### Syntax
```js
worksheetsCollection.getActiveWorksheet();
```
#### Parameters

None

#### Returns

[Worksheet](resources/worksheet.md) object.

#### Examples 

```js
var ctx = new Excel.ExcelClientContext();
var activeWorksheet = ctx.workbook.worksheets.getActiveWorksheet();
ctx.load(activeWorksheet);
ctx.executeAsync().then(function () {
		Console.log(activeWorksheet.name);
});
```
[Back](#worksheet)

### Get-Cell
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

[Range](resources/range.md) object.

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var cell = worksheet.cell(0,0);
ctx.load(cell);
ctx.executeAsync().then(function() {
	Console.log(cell.address);
});
```
[Back](#worksheet)

## Range

[Range](resources/range.md)  represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells. 

Following are the methods supported for this resource:

| Task                               | Description                                 |  
|:------------------------------------|:-------------------------------------------| 
| [Get-Range](#get-range)       | Retrieve Range properties and values based on Range Name or Address              |
| [Get-Range-Format](#get-range-format)       | Retrieve Range Formatting information such as font, border, background details.  |
| [Get-Range-Cell](#get-range-cell)         | Retrieve Cell properties                                                         |
| [Get-Used-Range-of-Range](#get-used-range-of-range)   | Retrieve used-range of within a Range                                            |
| [Insert-Range](#insert-range)     | Inserts a cell or a range of cells into the worksheet and shifts other cells away to make space. |
| [Update-Range](#update-range)           | Update Range values, formula, copy cells, change font properties, change background. |
| [Set-Range-Format](#set-range-format)       | Set format properties of a Range (font, background, wrap setting, etc.)              |
| [Set-Range-Border](#set-range-border)       | Add or update Range Borders                                                          |
| [Delete-Range](#delete-range)           | Remove the Range and its associated data/format                                      |
| [Clear-Range](#clear-range)            | Clear Range values, format, background, border, etc.                                 |
| [Entire-Row-Range](#entire-row-range)       | Returns a Range object that represents the entire row that contains the specified range.|
| [Entire-Column-Range](#entire-column-range)    | Returns a Range object that represents the entire column that contains the specified range.|


### Get-Range

Get a Range object that represents a single cell or a range of cells. 

#### Syntax

```js
worksheetObject.getRange(rangeAddress);
```

#### Returns

[Range](resources/range.md) object.

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

[Back](#range) 
### Get-Range-Format 

Get Range's format and styling details such as font, border, background information. This information is obtained by navigating to the font, background or borders property. 

#### Syntax

```js
rangeObject.format;
rangeObject.format.background;
rangeObject.format.font;
rangeObject.format.borders;
```

#### Returns

[Range Format](resources/format.md) object.
[Range Background](resources/background.md) object.
[Range Font](resources/font.md) object.
[Range Border Collection](resources/border.md) object.

Note: Depending on the need, you can select one or more of the format objects.

#### Examples

Below example selects all of the Range's format properties. 

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
ctx.load(range, select: ["format", "format/background", "format/borders", "format/font"] );
ctx.executeAsync().then(function() {
	Console.log(range.format.wrapText);
	Console.log(range.format.background.color);
	Console.log(range.format.font.name);
	Console.log(range.format.borders.getItem('InsideHorizontal').lineStyle;	
});
```

[Back](#range) 

### Get-Range-Cell

Get the Cell (as a Range object) object based on row and column address relative to a Range. 

Note that the returned object is a Range representing the single cell requested. The `address`, `columnIndex`, `rowIndex`, etc. property values of returned Range is relative to the worksheet. 

#### Syntax

```js
rangeObject.getCell(row, column);
```

#### Parameters 

Parameter      | Type   | Description
-------------- | ------ | ------------
`row`          | Number | Required. Row number of the cell to be retrieved. Zero indexed. 
`col`          | Number | Required. Column number of the cell to be retrieved. Zero indexed.

#### Returns

[Range](resources/range.md) object.

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var cell = range.cell(0,0);
ctx.load(cell);
ctx.executeAsync().then(function() {
	Console.log(cell.address);
});
```

[Back](#range) 
### Get-Used-Range-of-Range 

Get used-range portion within the requested Range object. 

#### Syntax

```js
rangeObject.getUsedRange();
```
##### Parameters

None

#### Returns

[Range](resources/range.md) object.

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D:F";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
var rangeUR = range.getUsedRange();
ctx.load(rangeUR);
ctx.executeAsync().then(function() {
	Console.log(rangeUR.address);
});
```

[Back](#range) 
### Insert-Range 

Inserts a cell or a range of cells into the worksheet and shifts other cells away to make space.

#### Syntax
```js
rangeObject.insert(shift);
```
#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`shift`| String | Optional. Specifies which way to shift the cells. Can be one of the following: `Right` or `Down`. If this argument is omitted, Microsoft Excel decides based on the shape of the range.

#### Returns
Nothing

#### Example

```js
var sheetName = "Sheet1";
var rangeAddress = "F5:F10";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.insert();
ctx.executeAsync().then();
```

[Back](#range) 
### Update-Range 

Set Range values, formula, number format.

#### Syntax
```js
rangeObject.property = value;
```
Where, property is one of the following Range properties that can be set. 

#### Properties

|Property          | Type          | Description                                           |
|----------------- | -------------- | ----------------------------------------------------- |
|`values`		   | Array [][] (string) or (number)    | Unformatted value of the specified range.	 		        |
|`numberFormat`    | Array [][] (String) | Typethat represents the format code for the object. |
|`formula`         | Array [][] (String) | Represents the object's formula notation.             |
|`formulaLocal`    | Array [][] (String) | Formula for the object, in the language of the user.  |

#### Returns

[Range](resources/range.md) object.

#### Example
The example below sets number-format, values and formulas on a grid that contains 2x3 grid.

```js
var sheetName = "Sheet1";
var rangeAddress = "F5:G7";
var numberFormat = [[null, "d-mmm"], [null, "d-mmm"], [null, null]]
var values = [["Today", 42147], ["Tomorrow", "5/24"], ["Difference in days", null]];
var formula = [[null,null], [null,null], [null,"=G6-G5"]];
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.numberFormat = numberFormat;
range.values = values;
range.formula = formula;
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.text);
});
```

[Back](#range) 
### Set-Range-Format 

Set relevant format objects to update the Range Font, Background, alignment, and Wrap settings.

#### Syntax
```js
rangeObject.format.property = value;
```
Where, property is one of the following Range's Format properties that can be set. 

#### Properties

[Range Format](resources/format.md)

| Property         | Type    |Description|
|:-----------------|:--------|:----------| 
|`horizontalAlignment`    | String  |Optional. Represents the horizontal alignment for the specified object. The value of this property can be to one of the following constants: `Center`, `Distributed`, `Justify`, `Left`, `Right`. `null` indicates that the entire range doesn't have uniform horizontal alignment.|Range.HorizontalAlignment|
|`verticalAlignment`    | String  |Optional. Represents the vertical alignment for the specified object. The value of this property can be to one of the following constants: `Bottom`, `Center`, `Distributed`, `Justify`, `Top`. `null` indicates that the entire range doesn't have uniform vertical alignment.|Range.VerticalAlignment|
|`wrapText`    | Boolean  |Optional. Indicates if Excel wraps the text in the object. `null` indicates that the entire range doesn't have uniform wrap setting|Range.WrapText|    

[Range Font](resources/font.md)

| Property         | Type    |Description| 
|:-----------------|:--------|:----------|
|`name`|String|Font name (e.g., "Calibri")| 
|`size`|Integer|Size of the font (e.g., 11)|
|`color`|String|HTML color code representation of the text color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |
|`italic`|Boolean| Represents the bold status of italic. `true` if the font style is italic|
|`bold`|Boolean| Represents the bold status of font. `true` if the font is bold. |
|`strikethrough`|Boolean| `true` if the font is struck through with a horizontal line. `false` by default.| 
|`subscript`|Boolean| `true` if the font is formatted as subscript. `false` by default.| 
|`superscript`|Boolean| `true` if the font is formatted as superscript; `false` by default.|
|`underlineStyle`|String|Type of underline applied to the font. Can be one of the following constants. Possible Values: `None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting`)|

[Range Background](resources/background.md)


| Property         | Type    |Description|
|:-----------------|:--------|:----------| 
|`color`|String|HTML color code representation of the Background color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |

#### Example
The example below sets font name, background color and wraps text. 

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.wrapText = true;
range.format.font.name = 'Times New Roman';
range.format.background.color = '0000FF';
ctx.executeAsync().then();
```

[Back](#range) 
### Set-Range-Border 

Sets border to a range and sets the Color, LineStyle, and Weight properties for the new border.

#### Syntax
```js
rangeObject.format.borders(sideIndex).property = value;
```
Where, property is one of the following Range's border properties that can be set. 

#### Properties

Property       | Type   | Description
--------------- | ------ | ------------
`lineStyle`| String | One of the constants of LineStyle specifying the line style for the border. Options are: `Continuous`: Continuous line, `Dash`: Dashed line, `DashDot`: Alternating dashes and dots, `DashDotDot`: Dash followed by two dots, `Dot`: Dotted line, `Double`: Double line, `None`: No line, `SlantDashDot`: Slanted dashes.|Border.LineStyle
`weight`| String | BorderWeight value that specifies the weight of the border around a range. Options are: `Hairline`: Hairline (thinnest border), `Medium`: Medium, `Thick`: Thick (widest border), `Thin`: Thin.|Border.Weight
`color`| String | HTML color code representing the color of the border line|Border.Color's representation in HTML color code.


** sideIndex values: **

`sideIndex` values | Type  | Description
--------------- | ------ | ------------
`DiagonalDown`  |String | Border running from the upper left-hand corner to the lower right of each cell in the range. 
`DiagonalUp`    |String |Border running from the lower left-hand corner to the upper right of each cell in the range.
`EdgeBottom`    |String |Border at the bottom of the range.
`EdgeLeft`      |String |Border at the left-hand edge of the range.
`EdgeRight`     |String |Border at the right-hand edge of the range.
`EdgeTop`       |String |Border at the top of the range.
`InsideHorizontal` |String|Horizontal borders for all cells in the range except borders on the outside of the range.
`InsideVertical`|String |Vertical borders for all the cells in the range except borders on the outside of the range.

#### Example
The example below adds grid border around the range.

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
range.format.borders('InsideVertical').lineStyle = 'Continuous';
range.format.borders('EdgeBottom').lineStyle = 'Continuous';
range.format.borders('EdgeLeft').lineStyle = 'Continuous';
range.format.borders('EdgeRight').lineStyle = 'Continuous';
range.format.borders('EdgeTop').lineStyle = 'Continuous';
ctx.executeAsync().then();
```

[Back](#range) 
### Delete-Range

Delete the Range data and clear the format and shift the cells.

#### Syntax

```js
rangeObject.delete(shift);
```
##### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
shift| String | Specifies which way to shift the cells. Can be one of the following: `Left` or `Up`. If this argument is omitted, Microsoft Excel decides based on the shape of the range.

#### Returns

Nothing

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D:F";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.delete();
ctx.executeAsync().then();
```

[Back](#range) 
### Clear-Range

Clear Range values, format, background, border, etc.

#### Syntax

```js
rangeObject.clear(applyTo);
```

##### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`applyTo` | String | Optional. `All`, `Format`, `Content`. If this option is not provided then the content and format of the range will be cleared. 

#### Returns

Nothing

#### Examples
Below example clears format and contents of the range. 

```js
var sheetName = "Sheet1";
var rangeAddress = "D:F";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.clear();
ctx.executeAsync().then();
```

[Back](#range) 
### Entire-Row-Range

Get an object that represents the entire row of the Range. This API is valid only if the subject range object is a single cell or a row of cells.

#### Syntax

```js
rangeObject.getEntireRow();
```
##### Parameters

None

#### Returns

[Range](resources/range.md) object.
** Note: the grid properties of the Range (values, numberFormat, formula) contains `null` since the Range in question is unbounded. **
#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D:F";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
var rangeER = range.getEntireRow();
ctx.load(rangeER);
ctx.executeAsync().then(function() {
	Console.log(rangeER.address);
});
```

[Back](#range) 
### Entire-Column-Range

Get an object that represents the entire column of the Range. This API is valid only if the subject range object is a single cell or a column of cells.

#### Syntax

```js
rangeObject.getEntireColumn();
```
##### Parameters

None

#### Returns

[Range](resources/range.md) object.
** Note: the grid properties of the Range (values, numberFormat, formula) contains `null` since the Range in question is unbounded. **
#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D:F";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
var rangeEC = range.getEntireColumn();
ctx.load(rangeEC);
ctx.executeAsync().then(function() {
	Console.log(rangeEC.address);
});
```

## Table

[Table](resources/table.md) represents collection of organized cells designed to make management of the data easy. Table consists of collection of [Table Row](resources/tableRow.md)and [TableColumn](resources/tableCoulmn.md) objects, which represents rows and Columns in a table. 

Following are the methods supported for this resource:

| Task                               | Description                                | 
|:------------------------------------|:-------------------------------------------|
| [Get-Table](#get-table)  | Retrieve Table properties                               |
| [Get-Table-Range](#get-table-range])     | Retrieve the Range object associated with the Table     |
| [Get-Header-Row](#get-header-row)        | Retrieve the Header Row Range object associated with the Table      |
| [Get-Data-Body](#get-data-body)          | Retrieve the Range object associated with Data Body of the Table.   |
| [Get-Totals-Row](#get-totals-row)        | Retrieve the Range object associated with Totals row  of the Table. |
| [Add-Table](#add-table)                  | Create a new Table Object                                       |
| [Resize-Table](#resize-table)            | Resize the table over a new Range.                              |
| [Delete-Table](#delete-table)            | Deletes Table and clears the cell data from the Table.          |
| [Update-Table](#update-table)            | Update Table properties such as Name, show totals, change table style, etc.         |
| [List-Table-Rows](#list-table-rows)     | Retrieve List of Rows of a Table                        |
| [Get-Row](#get-row)             | Retrieve Table Row's data and properties                |
| [Get-Row Range](#get-row-range)       | Retrieve Range Object associated with Table Row         |
| [Add-Row](#add-row)             | Add a Row in the Table                                  |
| [Update-Row](#update-row)          | Update values of row of data in the Table               |
| [Delete-Row](#delete-row)          | Deletes the cells of the Table Row and shifts upward any remaining cells below the deleted row.   |
| [List-Table-Columns](#list-table-columns)    | Retrieve List of Columns of a Table                     |
| [Get-Column](#get-column)         | Retrieve Table Column's data and properties                |
| [Get-Header-Row Range](#get-header-row)   | Retrieve the header row's Range for a Table Column Object|
| [Get-Total-Range](#get-total-range)    | Retrieve the total row's Range for a Table Column Object     |
| [Get-Data-Body-Range](#get-data-body-range)    | Retrieve the Range object that is the data portion of a Table Column |
| [Add-Column](#add-column)         | Add a Column in the Table  |
| [Update-Column](#update-column)   | Update values of column of data in the Table. |
| [Delete-Column](#delete-column)   | Deletes the column of data in the Table. |
 

### Get-Table

Get Table object properties based on name. 

#### Syntax

```js
tablesCollection.getItem(name);
```

#### Parameters

Parameter        | Type   | Description
---------------  | ------ | ------------
 `name`| String  | Required. Table name. 

#### Syntax
```js
tablesCollection.getItemAt(index);
```

#### Parameters

Parameter        | Type   | Description
---------------  | ------ | ------------
 `index`| Number | Required. Table index. Zero indexed.

#### Returns

[Table](resources/table.md) object. 

#### Examples

##### Getting a table by name

```js
var ctx = new Excel.ExcelClientContext();
var tableName = 'Table1';
var table = ctx.workbook.tables.getItem(tableName);
ctx.executeAsync().then(function () {
		Console.log(table.index);
});
```
##### Getting a table by index

```js
var ctx = new Excel.ExcelClientContext();
var index = 0;
var table = ctx.workbook.tables.getItemAt(0);
ctx.executeAsync().then(function () {
		Console.log(table.name);
});
```

[Back](#table)

### Get-Table-Range

Get Range object associated with the Table.

#### Syntax
```js
tableObject.getRange();
```

#### Parameters

None

#### Returns

[Range](resources/range.md) object.


#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableRange = table.getRange();
ctx.executeAsync().then(function () {
		Console.log(tableRange.address);
});
```

[Back](#table)

### Get-Header-Row

Get Header Range object associated with the Table.

#### Syntax
```js
tableObject.getHeaderRowRange();
```

#### Parameters

None

#### Returns


[Range](resources/range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableHeaderRange = table.getHeaderRowRange();
ctx.executeAsync().then(function () {
		Console.log(tableHeaderRange.address);
});
```
[Back](#table)

### Get-Data-Body

Get Data Body Range object associated with the Table.

#### Syntax
```js
tableObject.getDataBodyRange();
```

#### Parameters

None

#### Returns

[Range](resources/range.md) object.


#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableDataRange = table.getDataBodyRange();
ctx.executeAsync().then(function () {
		Console.log(tableDataRange.address);
});
```
[Back](#table)

### Get-Totals-Row

Get Totals Range object associated with the Table.

#### Syntax
```js
tableObject.getTotalRowRange();
```

#### Parameters

None

#### Returns

[Range](resources/range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableTotalsRange = table.getTotalRowRange();
ctx.executeAsync().then(function () {
		Console.log(tableTotalsRange.address);
});
```
[Back](#table)

### Add-Table
 
Create a New Table object.

#### Syntax
```js
tablesCollection.add(name, rangeSource, containsHeader, showTotals, tableStyle )
```
#### Parameters 

|Parameter       | Type   | Description
|--------------- | ------ | ------------
|`name`  | String | Optional. String value representing the name of the Table.
|`rangeSource`| String | Required. Address or name of the Range object representing the data source.
|`containsHeader` | Boolean | Optional. Boolean value that indicates whether the data being imported has column labels. If the Source does not contain headers (i.e,. when this property set to `false`), Excel will automatically generate headers. If this property value is not set, Excel will determine the header row on its own.
|`showTotals` | Boolean| Optional. Boolean to indicate whether the Total row is visible. This value can be set to show or remove the total row. By default this will be set to `false` 
|`tableStyle` | String | Optional. Constant that represents the Table style. Possible values include: `Light1` thru `Light21`, `Medium1` thru `Medium28`, `Dark1` thru `Dark11`. Excel determines the default style if one is not specified. 

#### Returns
[Table](resources/table.md) object.


#### Example
```js
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.add('MyTable', 'Sheet1!A1:E7', true, false, null);
ctx.load(table);
ctx.executeAsync().then(function () {
	Console.log(table.name);
});

```
[Back](#table)

### Resize-Table

Resize the table over a new Range. The top of the table must remain in the same row and the resulting table must overlap the original table.   

#### Syntax
```js
tableObject.resize(rangeSource);
```
#### Returns
[Table](resources/table.md) object.

#### Parameters 

Parameter       | Type   | Description
--------------- | ------ | ------------
`rangeSource`| String | Required. Address or name of the Range object representing the new data source.

#### Example 

```js
var tableName = 'Table1';
var newRangeSource = 'Sheet1!A2:D10';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
var rTable = table.resize(newRangeSource);
ctx.load(rTable);
ctx.executeAsync().then(function () {
	Console.log(rTable.name);
});
```
[Back](#table)

### Delete-Table

Deletes Table and clears the cell data from the Table.

#### Syntax
```js
tableObject.delete();
```

#### Parameters 
None

#### Returns
Nothing

#### Example 

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
table.deleteObject();
ctx.executeAsync().then();
```
[Back](#table)

### Update-Table 

This API allows setting of Table properties such as name and show totals. In order to update the table content, use the update table row or column API.

Deletes Table and clears the cell data from the Table.

#### Syntax
```js
tableObject.property = 'new-value';
```

#### Properties 

Following properties can be updated directly. 

|Property      | Type   | Description      |
|-------------- | ------ | -----------------|
| `name`        | String | String value that represents the name of the Table object   | 
| `showTotals`  | Boolean| Boolean to indicate whether the Total row is visible. This value can be set to show or remove the total row| 
| `tableStyle`  | String | Constant that represents the Table style. Possible values include: `TableStyleLight1` thru `TableStyleLight21`, `TableStyleMedium1` thru `TableStyleMedium28`, `TableStyleDark1` thru `TableStyleDark11`|

#### Example 

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
table.name = 'Table1-Renamed';
table.showTotals = false;
table.tableStyle = 'TableStyleMedium2';
ctx.load(table);
ctx.executeAsync().then(function () {
		Console.log(table.tableStyle);
});
```
[Back](#table)

### List-Table-Rows 

Get a list of Rows of a Table.   

#### Syntax
```js
tableObject.tableRows;
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. Table name. 

#### Returns

[Table Row](resources/tableRow.md) colelction

#### Examples

```js
var tableName = 'Table1'
var ctx = new Excel.ExcelClientContext();
var tableRows = context.workbook.tables.getItem(name).tableRows;
ctx.load(tableRows);
ctx.executeAsync().then(function () {
	for (var i = 0; i < tableRows.items.length; i++)
	{
		Console.log(tableRows.items[i].index);
	}
});
```
[Back](#table)

### Get-Row 

Get Table Row's data and properties  

#### Syntax
```js
tableRowsCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Row index of the row that you wish to get. Zero indexed.

#### Returns

[Table Row](resources/tableRow.md) object.


#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(0);
ctx.load(row);
ctx.executeAsync().then(function () {
	Console.log(row.index);
});
```
[Back](#table)

### Get-Row Range 
Get Range object associated with the Row.

#### Syntax
```js
tableRowObject.getRange();
```

#### Parameters

None

#### Returns

[Range](resources/range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(0);
var rowRange = row.getRange();
ctx.load(rowRange);
ctx.executeAsync().then(function () {
	Console.log(rowRange.address);
});
```

[Back](#table)
### Add-Row 

Add a Row in the Table.

#### Syntax
```js
tableRowsCollection.add(index, values);
```

#### Parameters 
Parameter       | Type   | Description
--------------- | ------ | ------------
`index` |  Number |Optional. Specifies the relative position of the new row. If not specified, the addition happens at the end. The previous column at this position is shifted outward to the bottom. **Zero Indexed**
`values` | Collection (primitive) | 2-D array of unformatted values of the table row. 

#### Returns
[Table Row](resources/tableRow.md) object.

#### Example
```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
var values = [["Sample", "Values", "For", "New", "Row"]];
var row = tables.getItem("Table1").tablerows.add(null, values);
ctx.load(row);
ctx.executeAsync().then(function () {
	Console.log(row.index);
});
```


[Back](#table)
### Update-Row 

Update values of table row.

#### Syntax
```js
tableRowObject.values = new-values
```
New-values is a 2-D array values of the table row 

#### Example
```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
var newValues = [["New", "Values", "For", "New", "Row"]];
var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
row.values = newValues;
ctx.load(row);
ctx.executeAsync().then(function () {
	Console.log(row.values);
});
```

[Back](#table)
### Delete-Row  

Deletes Table Row and clears the cell data from Table row.

#### Syntax

```js
tableRowObject.delete();
```

#### Parameters 
None

#### Returns
Nothing

#### Example 

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var row = ctx.workbook.tables.getItem(tableName).tableRows.getItemAt(2);
row.deleteObject();
ctx.executeAsync().then();
```


[Back](#table)
### List-Table-Columns 
Get a list of columns of a Table.   

#### Syntax

```js
tableObject.tableColumns;
```

#### Returns

[Table Row](resources/tableRow.md) colelction

#### Examples

```js
var tableName = 'Table1'
var ctx = new Excel.ExcelClientContext();
var tableColumns = context.workbook.tables.getItem(name).tableColumns;
ctx.load(tableColumns);
ctx.executeAsync().then(function () {
	for (var i = 0; i < tableColumns.items.length; i++)
	{
		Console.log(tableColumns.items[i].index);
	}
});
```
[Back](#table)
### Get-Column 

Get Table Column's data and properties.  

#### Syntax
```js
tableColumnsCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Column index of the column that you wish to get. Zero indexed.

#### Returns

[Table Column](resources/tableColumn.md) object.


#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
ctx.load(column);
ctx.executeAsync().then(function () {
	Console.log(column.index);
});
```
[Back](#table)
### Get-Header-Row-Range 

Get Range object associated with the Column's header.

#### Syntax

```js
tableColumnObject.getHeaderRowRange();
```


#### Parameters

None

#### Returns

[Range](resources/range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var row = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var headerRowRange = row.getHeaderRowRange();
ctx.load(headerRowRange);
ctx.executeAsync().then(function () {
	Console.log(headerRowRange.address);
});
```


[Back](#table)
### Get-Total-Range 

Get Range object associated with the Column's total.

#### Syntax 

```js
tableColumnObject.getTotalRowRange();
```

#### Parameters

None

#### Returns

[Range](resources/range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var row = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var totalRowRange = row.getTotalRowRange();
ctx.load(totalRowRange);
ctx.executeAsync().then(function () {
	Console.log(totalRowRange.address);
});
```

[Back](#table)
### Get-Data-Body-Range 
Get Range object associated with the Column's data body.

```js
tableColumnObject.getDataBodyRange();
```

#### Parameters

None

#### Returns

[Range](resources/range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var row = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var dataBodyRange = row.getDataBodyRange();
ctx.load(dataBodyRange);
ctx.executeAsync().then(function () {
	Console.log(dataBodyRange.address);
});
```

[Back](#table)
### Add-Column 

Add a Column in the Table.  

#### Syntax
```js
tableColumnsCollection.add(index, values);
```

Parameter       | Type   | Description
--------------- | ------ | ------------
`index` |  Number |Optional. Specifies the relative position of the new column. The previous column at this position is shifted outward to the right. If not specified, the addition happens at the end. **Zero Indexed**. **Note: The index value should be equal to or less than the last column's index value. In other words, this API cannot be used to append a column at the end of the table **
`values` | Collection (primitive) | 2-D array of unformatted values of the table column.

#### Returns
[Range](resources/range.md) object.

#### Example
```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
var row = tables.getItem("Table1").tableColumns.add(null, values);
ctx.load(row);
ctx.executeAsync().then(function () {
	Console.log(row.name);
});
```
[Back](#table)
### Update-Column 


Update values of table column.

#### Syntax
```js
tableColumnObject.values = new-values
```
Where, new-values is a 2-D array values of the table column. 

#### Example

```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
column.values = newValues;
ctx.load(column);
ctx.executeAsync().then(function () {
	Console.log(column.values);
});
```

[Back](#table)
### Delete-Column 

Deletes Table Column and clears the cell data from the Table Column.

#### Syntax

```js
tableColumnObject.delete();
```

#### Parameters 
None

#### Returns
Nothing

#### Example 

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
column.deleteObject();
ctx.executeAsync().then();
```


## Chart

[Chart](resources/chart.md) represents a chart object on a worksheet. 

Following are the methods supported for this resource:

| Task                               | Description                                | 
|:------------------------------------|:-------------------------------------------|
| [Add-Chart](#add-chart)     | Inserts a chart directly onto the grid.  |
| [Get-Chart](#get-chart)   | Gets a chart by name. |
| [Delete-Chart](#delete-chart)     | Deletes a chart directly on the grid.  |
| [Update-Chart](#update-chart)   | Update a chart including renaming, positioning and resizing. |
| [Set-Chart-SourceData](#set-chart-sourcedata)   | Sets the sourceData and seriesBy of a Chart.|
| [Format-Chart](#format-chart)   | Format a chart.|
| [Get-Chart-Title](#get-chart-title)   | Get the title of a chart. |
| [Set-Chart-Title](#set-chart-title)   | Set the title of a chart, including `text`, `position` and `overlay`. |
| [Delete-Chart-Title](#delete-chart-title)   | Delete the title from a chart. |
| [Format-Chart-Title](#format-chart-title)   | Format the Chart Title. |
| [Set-Chart-Legend](#set-chart-legend)   | Hide/Show Chart Legent and set position. |
| [Set-Chart-DataLabels](#set-chart-datalabels)   | Set display content and position of DataLabels. |
| [Set-Chart-Axis](#set-chart-axis)   | Set the `maximum`, `minimum`, `majorunit`,`minorunit` and `visible`of an axis. |
| [Set-Chart-AxisTitle](#set-chart-axistitle)   | Change the Axis Title text and visibility. |
| [Add-Chart-Gridlines](#add-chart-gridlines)   | Show Gridlines on an Axis |
| [Format-Chart-Series](#format-chart-series)   | Change the Fill Color of a series |


 

### Add-Chart

Inserts a chart directly onto the grid.

#### Syntax

```js
chartsCollection.add(chartType, sourceData, seriesBy);
```

#### Parameters

| Parameter         | Value    |Description|
|:-----------------|:--------|:----------|
| `type` | String | A String value that represents the type of a chart.  |
| `sourceData`  | String | Sets an address or name of the Range object as the data source.|
| `seriesBy` | String | Sets the way columns or rows are used as data series on the chart. Can be `auto`, `Rows` or `Columns`.|


#### Returns

[Chart](resources/chart.md) object. 

#### Examples

##### Add a chart of `chartType` "ColumnClustered" on worksheet "Charts" with `sourceData` from Range "A1:B4" and `seriresBy` is set to be "auto".

```js
var sheetName = "Charts";
var sourceData = sheetName + "!" + "A1:B4";
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
ctx.executeAsync().then(function () {
		logComment("New Chart Added");
});
```
[Back](#chart)

### Get-Chart

Gets a chart by name.

#### Syntax
```js
chartsCollection.getItem(name);	
```

#### Parameters
None.

#### Returns

[Chart](resources/chart.md) object. 

#### Examples

##### Get the Chart named "Chart1"
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

ctx.load(chart);
ctx.executeAsync().then(function () {
		logComment("Chart1 Loaded");
});
```

[Back](#chart)

### Delete-Chart

Deletes a chart directly on the grid.

#### Syntax

```js
chartObject.deleteObject();
```

#### Parameters
None.

#### Returns

Nothing.

#### Examples

##### Delete the Chart named "Chart1"

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.deleteObject();
ctx.executeAsync().then(function () {
		logComment"Chart Deleted");
});
```
[Back](#chart)

### Update-Chart

Update a chart including renaming, positioning and resizing.

#### Syntax

```js
chartObject.name="New Name";
chartObject.top = 100;
chartObject.left = 100;
chartObject.height = 200;
chartObject.weight = 200;
```

#### Parameters
None.

#### Returns

[Chart](resources/chart.md) object. 

#### Examples

##### Rename the chart to new name, resize the chart to 200 points in both height and weight. Move Chart1 to 100 points to the top and left. 
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");

chart.name="New Name";	
chart.top = 100;
chart.left = 100;
chart.height =200;
chart.width =200;
ctx.executeAsync().then(function () {
		logComment("Chart Updated");
});
```
[Back](#chart)


### Set-Chart-SourceData

Sets the sourceData and seriesBy of a Chart.

#### Syntax

```js
chartObject.setData(sourceData, seriesBy);
```

#### Parameters
| Parameter         | Value    |Description|
|:-----------------|:--------|:----------|
| `sourceData`  | String|  Sets an address or name of the Range object as the data source.|
| `seriesBy`  | String |  Sets the way columns or rows are used as data series on the chart. Can be one of the following `Rows`, `Columns` or `Auto`.|

#### Returns

[Chart](resources/chart.md) object. 

#### Examples

##### Set the `sourceData` to be "A1:B4" and `seriesBy` to be "Columns"
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	
var sourceData = "A1:B4";

chart.setData(sourceData, "Columns");
ctx.executeAsync().then();
```
[Back](#chart)


### Format-Chart

Format a chart.

#### Syntax

```js
chartObject.fillFormat.SetSolidColor(color);
```

#### Parameters
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|color| String | HTML color code representing the color of the interior/background. |

#### Returns
Nothing.

#### Examples

##### Set "Chart1" background to be red.

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.fillFormat.SetSolidColor("#FF0000");
ctx.executeAsync().then(function () {
		logComment("Chart Color Changed ");
});
```
[Back](#chart)


### Get-Chart-Title

Get the title of a chart.

#### Syntax

```js
chartObject.title.text;
```

#### Parameters
None. 

#### Returns
[ChartTitle](resources/chartTitle.md) object. 

#### Examples

##### Get the `text` of Chart Title from Chart1
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

ctx.load(chart);
var title = chart.title.text;

ctx.executeAsync().then(function () {
		logComment(title);
});
```
[Back](#chart)


### Set-Chart-Title

Set the title of a chart, including `text`, `position` and `overlay`.

#### Syntax

```js
chartObject.title.text= "My Chart"; 
chartObject.title.position = "top";
chartObject.title.overlay=true;

```

#### Parameters
None. 

#### Returns
[ChartTitle](resources/chartTitle.md) object. 

#### Examples

##### Set the `text` of Chart Title to "My Chart" and Make it show on top of the chart without overlaying.
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.title.text= "My Chart"; 
chart.title.position = "top";
chart.title.overlay=true;

ctx.executeAsync().then(function () {
		logComment("Char Title Changed");
});
```
[Back](#chart)

### Delete-Chart-Title

Delete the title from a chart.

#### Syntax

```js
chartObject.title.visible = false; 
```

#### Parameters
None. 

#### Returns
None.

#### Examples

##### Hide the title of Chart1
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.title.visible = false; 
ctx.executeAsync().then(function () {
		logComment("Title Hidden");
});
```
[Back](#chart)

### Format-Chart-Title

Formats the title from a chart.

#### Syntax

```js
chartObject.title.font.bold = true; 
chartObject.title.font.color = "#FF0000";
```

#### Parameters
None. 

#### Returns
None.

#### Examples

##### Make the title of Chart1 to be bold and red

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.title.font.bold = true; 
chart.title.font.color = "#FF0000";

ctx.executeAsync().then(function () {
		logComment("Title Format Updated");
});
```
[Back](#chart)

### Set-Chart-Legend

 Hide/Show Chart Legent and set position. 

#### Syntax

```js
chartObject.legend.visible = true;
chartObject.legend.position = "top"; 
```

#### Parameters
None.

#### Returns
None.

#### Examples

##### Show Legend of Chart1 and make it on top of the chart.
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.legend.visible = true;
chart.legend.position = "top"; 
ctx.executeAsync().then(function () {
		logComment("Legend Shown ");
});
```
[Back](#chart)

### Set-Chart-DataLabels

Set display content and position of DataLabels.

#### Syntax

```js
chartObject.datalabels.visible = true;
chartObject.datalabels.position = "top";
chartObject.datalabels.ShowSeriesName = true;
```

#### Parameters
None.

#### Returns
None.

#### Examples

##### Make Series Name shown in Datalabels and set the `position` of datalabels to be "top";
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.datalabels.visible = true;
chart.datalabels.position = "top";
chart.datalabels.ShowSeriesName = true;

ctx.executeAsync().then(function () {
		logComment("Datalabels Shown");
});
```
[Back](#chart)


### Set-Chart-Axis

 Set the  `maximum` ,  `minimum` ,  `majorunit` , `minorunit`  and  `visible` of an axis. 

#### Syntax

```js
chartObject.axes.valueaxis.maximum = 5;
chartObject.axes.valueaxis.minimum = 0;
chartObject.axes.valueaxis.majorunit = 1;
chartObject.axes.valueaxis.minorunit = 0.2;
chartObject.axes.categoryaxis.visible = false;
```

#### Parameters
None.

#### Returns
None.

#### Examples

#####  Set the  `maximum`,  `minimum` ,  `majorunit` , `minorunit`  and  `visible` of valueaxis. 
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.axes.valueaxis.maximum = 5;
chart.axes.valueaxis.minimum = 0;
chart.axes.valueaxis.majorunit = 1;
chart.axes.valueaxis.minorunit = 0.2;
chart.axes.valueaxis.visible = true;


ctx.executeAsync().then(function () {
		logComment("Axis Settings Changed");
});
```
[Back](#chart)


### Set-Chart-AxisTitle

 Change the Axis Title text and visibility. 

#### Syntax

```js

chartObject.axes.valueaxis.title.text = "Catagory";

```

#### Parameters
None.

#### Returns
None.

#### Examples

##### Add Catagory as the title for the catagory Axis
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.axes.valueaxis.title.text = "Catagory";

ctx.executeAsync().then(function () {
		logComment("Axis Title Added ");
});
```
[Back](#chart)

### Add-Chart-Gridlines

Show Gridlines on an Axis. 

#### Syntax

```js
chartObject.axes.valueaxis.majorgridlines.visible = true;
```

#### Parameters
None.

#### Returns
None.

#### Examples

##### Show Major Gridlines on ValueAxis of Chart1

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.axes.valueaxis.majorgridlines.visible = true;

ctx.executeAsync().then(function () {
		logComment("Axis Title Added ");
});
```
[Back](#chart)

### Format-Chart-Series

Change the Fill Color of a series.

#### Syntax

```js
chartObject.series.GetItemAt(1).fillFormat.SetSolidColor("#FF0000");
```

#### Parameters
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|color| String | HTML color code representing the color of the interior/background. |

#### Returns
None.

#### Examples

##### Change the fill color of Series1 to be red
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.series.GetItemAt(1).fillFormat.SetSolidColor("#FF0000");

ctx.executeAsync().then(function () {
		logComment("Series Fill Color Changed ");
});
```
[Back](#chart)




## Error Messages

Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur. 

|error.code|error.message|
|---------:|---------:|
|InvalidArgument |The argument is invalid or missing or has an incorrect format.|
|InvalidRequest  |Cannot process the request.|
|InvalidReference|This reference is not valid for the current operation.|
|InvalidBinding  |This object binding is no longer valid due to previous updates.|
|InvalidSelection|The current selection is invalid for this operation.|
|Unauthenticated |Required authentication information is either missing or invalid.|
|AccessDenied    |You cannot perform the requested operation.|
|ItemNotFound    |The requested resource doesn't exist.|
|InvalidMethod   | The method in the request is not allowed on the resource. |
|EditConflict    |Request could not be processed because of conflict.|
|ActivityLimitReached|Activity limit has been reached.|
|GeneralException|There was an internal error while processing the request.|
|NotImplemented  |The requested feature isn't implemented.|
|ServiceNotAvailable|The service is unavailable.|

#### Examples

```js
ctx.executeAsync().then(
function () {
	Console.log("...");
    },
    function (error) {
	   some.log("ErrorCode =" + error.code); //"InvalidArgument"
	   some.log("ErrorMessage =" + error.message); //"The argument is invalid or missing or has an incorrect format."
	});

```
[top](#excel-javascript-apis)

## Programming Notes

Following sections provide important programming details related to Excel APIs.

* [Properties and Relations Selection](#properties-and-relations-selection)
* [Document Binding](#null-input)
* [Reference Binding](#null-input)
* [Null-Input](#null-input)
* [Null-Input](#null-input)
* [Null-Response](#null-response)
* [Blank Input and Output](#blank-input-and-output)
* [Unbounded-Range](#unbounded-range)
* [Large-Range](#large-range)
* [Single Input Copy](#single-input-copy)
* [Throttling](#throttling)

[top](#excel-javascript-apis)

### Properties and Relations Selection 

* By default load() selects all scalar/complex properties of the object which is being loaded. The relations are not loaded by default.  Exceptions:  any binary, XML, etc properties are not returned. 
* The select option specifies a subset of properties and/or relations to include in the response.
* Default Select behavior: 
	*	Does not select any property
	*	Need to specify every property that needs to be returned
	*	Relations/Navigation properties are also allowed to be included in the list. Use expand syntax to 
* The properties to be selected are provided during the load statement.
* Select will essentially get the users into optimized mode of handpicking what they want. 
* Property names are listed as a parameter to the select property. Support two kinds of inputs
	* Property names are separated by comma. 
	* Provide an array of property name strings

```js	
context.load (<object-var>, select: []);
context.load (<object-var>, select: "comma separated list of properties");
```

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.ExcelClientContext();
var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

//load statement below loads the address, values, numberFormat properties of the Range and then expands on the format, format/background, entireRow relations
 
ctx.load (myRange, select: ["address", "values", "numberFormat", format, format/background, entireRow ]);

ctx.executeAsync().then(function () {
		console.log (myRange.address); //ok
		console.log (myRange.cellCount); //not-ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.background.color); //ok
		console.log (myRange.format.font.color); //not-ok
		console.log (myRange.entireRow.address); //ok
		console.log (myRange.entireColumn.address); //not-ok
// . . . 

//load statement below loads all the properties of the Range and then expands on the format, format/background, entireRow relations. If the "*" is left out of the load, none of the Range’s direct properties will be included in the load statement.
 
ctx.load (myRange, select: ["*", "format", "format/background", "entireRow" ]);

ctx.executeAsync().then(function () {
		console.log (myRange.address); //ok
		console.log (myRange.cellCount); //ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.background.color); //ok
		console.log (myRange.format.font.color); //not-ok
		console.log (myRange.entireRow.address); //ok
		console.log (myRange.entireColumn.address); //not-ok

```

[Back](#programming-notes)
### Document Binding

[Back](#programming-notes)
### Reference Binding

[Back](#programming-notes)
### Null-Input

#### null input in 2-D Array

**`null` input inside 2 dimensional array (for values, number-format, formula) is ignored** in the update API. No update will take place to the intended target when `null` input is sent in values or number-format or formula grid of values.

Example: In order to only update specific parts of the Range such as some cell's Number Format and retain the existing Number Format on other parts of the Range, set desired Number Format where needed and send `null` for the other cells. 

In below set request, only some parts of the Range Number Format is set while retaining the existing Number Format on the remainig part (by passing nulls).

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
#### null input for a property

**`null` is not a valid single input for the entire property.** e.g., following is not valid as the entire values cannot be set to null or ignored. 

```
 range.values= null;

```

Following is not valid either as null is not a valid color value. 
```
 range.format.background.color =  null;
```
[Back](#programming-notes)
### Null-Response

Representation of formatting properties that consists of non-uniform values would result in `null` value to be returned in the response. 

Example: A Range can consist of one of more cells. In cases where the individual cells contained in the Range specified doesn't have uniform formatting values, the range level representation will be undefined. 

```
  "size" : null,
  "color" : null,
```





### Blank Input and Output

Blank values in update requests are treated as instruction to clear or reset the respective property. Blank value is represented by two double-quotes with no space in between. `""`

Example: 
* For `values`, the range value is cleared out. This is same as clearing the contents in the application.
* For `numberFormat`, the number format is set to `General`.
* For `formula` and `formulaLocale`, the formula values are clearned out. 

For read operations, expect to receive blank values if the contents of the cells are blanks. If the cell contains no data or value, then the API returns a blank value. Blank value is represented by two double-quotes with no space in between. `""`.

```
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```
  range.formula = [["", "", "=Rand()"]];
```
[Back](#programming-notes)
### Unbounded-Range

#### Read

Unbounded range address contains only column or row identifiers and unspecified row identifier or column identifiers (respectively), such as:

* `C:C`, `A:F`, `A:XFD` (contains unspecified rows)
* `2:2`, `1:4`, `1:1048546` (contains unspecified columns)

When the API makes a request to retrieve an unbounded Range (e.g., `getRange('C:C')`, the response returned contains `null` for cell level properties such as `values`, `text`, `numberFormat`, `formula`, etc.. Other Range properties such as `address`, `cellCount`, etc. will reflect the unbounded range.

#### Write

Setting cell level properties (such as values, numberFormat, etc.) on unbounded Range is **not allowed** as the input request might be too large to handle. 

Example: following is not a valid update request as the requested range is unbounded one. 
```js
var sheetName = 'Sheet1';
var rangeAddress = 'A:B';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
range.values = 'Due Date';
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.text);
});

```

When such a Range is update operation is attempted, the API returns the an error.

[Back](#programming-notes)
### Large-Range

Large Range implies a Range whose size is too large for a single API call. Many factors such as number of cells or values or number-formats, or formulas, etc. contained in the range can make the response large enough to be unsuitable for API interaction. 

The API makes best attempt to return or write-to the requested data. However, due to the large size involved, API might result in an error condition due to large resource utilization. 

In order to avoid such condition, it is recommended to read or write large Range in multiple smaller range sizes.

[Back](#programming-notes)
### Single Input Copy

To support updating a range with same values or number-format or applying same formula across a range, the following convention is used in the set API. In Excel, this behavior is similar to inputting values or formulas to a range in the CTRL+Enter mode. 

API will look for *single cell value* and and if the target range dimension doesn't match the input range dimension it will apply the update to the entire range in the CTRL+Enter model with the value or formula provided in the request.

#### Examples

Following request updates selected range with the a text of "Due Date". Note that Range has 20 cells whereas the provided input only has 1 cell value.

```js
var sheetName = 'Sheet1';
var rangeAddress = 'A1:A20';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
range.values = 'Due Date';
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.text);
});

```

Following request updates selected range with date of 3/11/2015".  

```js
var sheetName = 'Sheet1';
var rangeAddress = 'A1:A20';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
range.numberFormat = 'm/d/yyyy';
range.values = '3/11/2015';
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.text);
});

```
Following request updates selected range with a formula of that will be applied across in the CTRL+Enter mode.  

```js
var sheetName = 'Sheet1';
var rangeAddress = 'A1:A20';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
range.formula = '=DAYS(B15,42060)';
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.text);
});
```
[Back](#programming-notes)
### Throttling 

Excel Service uses throttling to maintain optimal performance and reliability of the service. Throttling limits the number of user actions or concurrent calls (by script or code) to prevent overuse of resources.

Though this is less common, certain pattern of API usage such as high frequency requests or high volume requests that increases CPU or memory utilization of the servers beyond limit would likely get you throttled.

When a user exceeds usage limits, Excel service throttles any further requests from that user account for a short period. All user actions are throttled while the throttle is in effect.

API requests while the throttle is in effect will result in below error condition:

```js
ctx.executeAsync().then(
function () {
	Console.log("...");
    },
    function (error) {
	   some.log("ErrorCode =" + error.code); //"ActivityLimitReached"
	   some.log("ErrorMessage =" + error.message); //"Activity limit has been reached."
	});
```
[Back](#programming-notes)

[top](#excel-javascript-apis)