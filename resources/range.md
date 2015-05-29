# Range
Range represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells. 

## [Properties](#get-range)
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`address`         |String         |Returns a String value that represents the range reference in A1 Style. **Address value will contain the Sheet reference (e.g., `Sheet1!A1:B4`)**|Range.Address|
|`addressLocal`    |String         |Returns the range reference for the specified range in the language of the user.
|`cellCount`       | Number          |Number of cells in the range|Range.Count|
|`columnIndex`     | Number          |Returns the number of the first column in the first area in the specified range. This is adjusted to be zero indexed. Read-only|Range.Column|
|`columnCount`    | Number           |Returns the number of the first row of the first area in the range. This is adjusted to be zero indexed. Read-only|Range.Row|
|`formula`         |Array [][]|Represents the object's formula in A1 style notation|Range.formula|
|`formulaLocal`    |Array [][]|Formula for the object, in the language of the user in A1 style notation|Range.FormulaLocal|
|`numberFormat`    |Array [][]|Value that represents the format code for the object|Range.NumberFormat
|`rowcount`        | Number          |Returns the total number of columns in the Range selected. Read-only |Range.Column|
|`rowIndex`        | Number          |Returns the number of the first row of the first area in the range. This is adjusted to be zero indexed. Read-only|Range.Row|
|`text`            |Array [][]|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel client will not affect the value returned by the API. |Range.Text|
|`values`          |Array [][]|Unformatted values of the specified range|Range.Value2|

## Relationships
Range resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|[format](#get-range-format)          |[Format](format.md) Object  |Format object contains Range's Font, Background, Borders, Alignment, Style, etc. settings ||
|[worksheet](#get-range-worksheet) |[Worksheet](worksheet.md) Object  |The worksheet containing the current range. ||


## Methods

The Worksheet resource has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[clear(applyTo: string)](#clearapplyto-string)| void     |Clear Range values, format, background, border, etc. |   |
|[delete()](#delete)| void     |Deletes the worksheet ||
|[getCell(row: number, column: number)](#getcellrow-number-column-number)| [Range](range.md) object |Returns a range containing the single cell specified by the zero-indexed row and column numbers          
|[getEntireColumn()](#getentirecolumn)| [Range](range.md) object |Get an object that represents the entire column of the Range. This API is valid only if the subject range object is a single cell or a column of cells.| |
|[getEntireRow()](#getentirerow)| [Range](range.md) object |Get an object that represents the entire row of the Range. This API is valid only if the subject range object is a single cell or a row of cells.| |
|[getUsedRange()](#getusedrange)| [Range](range.md) object |Returns the used range of the Range.| |  
|[insert(shift: string)](#insertshift-string)|void| Inserts a cell or a range of cells into the worksheet and shifts other cells away to make space.| |
|[select()](#select)|void| Select the specified Range in the Excel UI.| |

## API Specification 

[Back](#methods)

### delete()

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
[Back](#methods) 

### getCell(row: number, column: number)

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

[Range](range.md) object.

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
[Back](#methods) 

### getEntireColumn()

Get an object that represents the entire column of the Range. This API is valid only if the subject range object is a single cell or a column of cells.

#### Syntax

```js
rangeObject.getEntireColumn();
```
##### Parameters

None

#### Returns

[Range](range.md) object.
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
[Back](#methods)

### getEntireRow()

Get an object that represents the entire row of the Range. This API is valid only if the subject range object is a single cell or a row of cells.

#### Syntax

```js
rangeObject.getEntireRow();
```
##### Parameters

None

#### Returns

[Range](range.md) object.
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
[Back](#methods)

### getUsedRange()
Get used-range portion within the requested Range object. 

#### Syntax

```js
rangeObject.getUsedRange();
```
##### Parameters

None

#### Returns

[Range](range.md) object.

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

[Back](#methods) 

### insert(shift: string)

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
[Back](#methods) 

### select

Select the specified Range in the Excel UI.

#### Syntax
```js
rangeObject.select();
```
#### Parameters
None

#### Returns
Nothing

#### Example

```js
var sheetName = "Sheet1";
var rangeAddress = "F5:F10";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.select();
ctx.executeAsync().then();
```
[Back](#methods) 


### Get Range

Get a Range object that represents a single cell or a range of cells. 

#### Syntax

```js
worksheetObject.getRange(rangeAddress);
```

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

[Back](#properties) 

### clear(applyTo: string)

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

### Get Range Format 

Get Range's format and styling details such as font, border, background information. This information is obtained by navigating to the font, background or borders property. 

#### Syntax

```js
rangeObject.format;
rangeObject.format.background;
rangeObject.format.font;
rangeObject.format.borders;
```

#### Returns

[Range Format](format.md) object.
[Range Background](background.md) object.
[Range Font](font.md) object.
[Range Border Collection](border.md) object.

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
[Back](#relationships)

### Set Range Format 

Set relevant format objects to update the Range Font, Background, alignment, and Wrap settings.

#### Syntax
```js
rangeObject.format.property = value;
```
Where, property is one of the following Range's Format properties that can be set. 

#### Properties

[Range Format](format.md)

| Property         | Type    |Description|
|:-----------------|:--------|:----------| 
|`horizontalAlignment`    | String  |Optional. Represents the horizontal alignment for the specified object. The value of this property can be to one of the following constants: `Center`, `Distributed`, `Justify`, `Left`, `Right`. `null` indicates that the entire range doesn't have uniform horizontal alignment.|Range.HorizontalAlignment|
|`verticalAlignment`    | String  |Optional. Represents the vertical alignment for the specified object. The value of this property can be to one of the following constants: `Bottom`, `Center`, `Distributed`, `Justify`, `Top`. `null` indicates that the entire range doesn't have uniform vertical alignment.|Range.VerticalAlignment|
|`wrapText`    | Boolean  |Optional. Indicates if Excel wraps the text in the object. `null` indicates that the entire range doesn't have uniform wrap setting|Range.WrapText|    

[Range Font](font.md)

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

[Range Background](background.md)


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

[Back](#relationships)
### Set Range Border 

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
[Back](#relationships)

### Get Range Worksheet

Get Worksheet object of the current Range.

#### Syntax
```js
rangeObject.worksheet;
```
#### Returns

[Worksheet](worksheet.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var names = ctx.workbook.names;
var namedItem = names.getItem('MyRange');
range = namedItem.range;
var rangeWorksheet = range.worksheet;
load(rangeWorksheet)
ctx.executeAsync().then(function () {
		Console.log(rangeWorksheet.name);
});
```
[Back](#relationships)