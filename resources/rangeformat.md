# Range Format

Format object represents format settings of a Range. This includes Font, fill, Borders, Alignment, Style, etc. 

## [Properties](#get-range-format)
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`font`            |[Range Font](rangefont.md) Object                |Returns the Font object defined on the overall Range selected|Range.Font|
|`fill`      |[Range Fill](rangefill.md) Object    |Returns the fill object defined on the overall Range selected|Range.Interior|
|`horizontalAlignment`    | String  |Represents the horizontal alignment for the specified object. The value of this property can be to one of the following constants: `General`, `Fill`, `CenterAcrossSelection`, `Center`, `Distributed`, `Justify`, `Left`, `Right`. `null` indicates that the entire range doesn't have uniform horizontal alignment.|Range.HorizontalAlignment|
|`verticalAlignment`    | String  |Represents the vertical alignment for the specified object. The value of this property can be to one of the following constants: `Bottom`, `Center`, `Distributed`, `Justify`, `Top`. `null` indicates that the entire range doesn't have uniform vertical alignment.|Range.VerticalAlignment|
|`wrapText`    | Boolean  |Indicates if Excel wraps the text in the object. `null` indicates that the entire range doesn't have uniform wrap setting|Range.WrapText|


## Relationships
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`borders`         |[Range Border collection](rangebordercollection.md)|Collection of border objects that apply to the overall Range selected|Range.Borders|

## Methods

None

### Get Range Format 

Get Range's format and styling details such as font, border, fill information. This information is obtained by navigating to the font, fill or borders property. 

#### Syntax

```js
rangeObject.format;
rangeObject.format.fill;
rangeObject.format.font;
rangeObject.format.borders;
```

#### Returns

* [Range Format](rangeformat.md) object.
* [Range Fill](rangefill.md) object.
* [Range Font](rangefont.md) object.
* [Range Border Collection](rangeborder.md) object.

Note: Depending on the need, you can select one or more of the format objects.

#### Examples

Below example selects all of the Range's format properties. 

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
ctx.load(range, select: ["format", "format/fill", "format/borders", "format/font"] );
ctx.executeAsync().then(function() {
	Console.log(range.format.wrapText);
	Console.log(range.format.fill.color);
	Console.log(range.format.font.name);
	Console.log(range.format.borders.getItem('InsideHorizontal').lineStyle;	
});
```
[Back](#properties)

### Set Range Format 

Set relevant format objects to update the Range Font, fill, alignment, and Wrap settings.

#### Syntax
```js
rangeObject.format.property = value;
```
Where, property is one of the following Range's Format properties that can be set. 

#### Properties

[Range Format](Format.md)

| Property         | Type    |Description|
|:-----------------|:--------|:----------| 
|`horizontalAlignment`    | String  |Optional. Represents the horizontal alignment for the specified object. The value of this property can be to one of the following constants: `Center`, `Distributed`, `Justify`, `Left`, `Right`. `null` indicates that the entire range doesn't have uniform horizontal alignment.|Range.HorizontalAlignment|
|`verticalAlignment`    | String  |Optional. Represents the vertical alignment for the specified object. The value of this property can be to one of the following constants: `Bottom`, `Center`, `Distributed`, `Justify`, `Top`. `null` indicates that the entire range doesn't have uniform vertical alignment.|Range.VerticalAlignment|
|`wrapText`    | Boolean  |Optional. Indicates if Excel wraps the text in the object. `null` indicates that the entire range doesn't have uniform wrap setting|Range.WrapText|    

[Range Font](Font.md)

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

[Range fill](rangefill.md)

| Property         | Type    |Description|
|:-----------------|:--------|:----------| 
|`color`|String|HTML color code representation of the fill color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |

#### Example
The example below sets font name, fill color and wraps text. 

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.wrapText = true;
range.format.font.name = 'Times New Roman';
range.format.fill.color = '0000FF';
ctx.executeAsync().then();
```

[Back](#properties)

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


**sideIndex values:**

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
[Back](#properties)
