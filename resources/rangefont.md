# Range Font

This object represents the font attributes (font name, font size, color, and so on) for an object. 

## [Properties](#get-range-font)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`name`|String|Font name (e.g., "Calibri")|Range.Font.Name|
|`size`|Integer|Size of the font (e.g., 11)|Range.Font.Size|
|`color`|String|HTML color code representation of the text color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |Conversion from Range.Font.Color value to html color string|
|`italic`|Boolean|Represents the bold status of italic. true if the font style is italic|Range.Font.Italic|
|`bold`|Boolean|Represents the bold status of font. true if the font is bold. |Range.Font.Bold|
|`strikethrough`|Boolean|true if the font is struck through with a horizontal line. false by default.|Range.Font.Strikethrough|
|`subscript`|Boolean|true if the font is formatted as subscript. false by default.|Range.Font.Subscript|
|`superscript`|Boolean|true if the font is formatted as superscript; false by default.|Range.Font.Superscript  |
|`underlineStyle`|String|Type of underline applied to the font. Can be one of the following constants. Possible Values: `None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting`|Range.Font.Underline|

## Relationships
None

## Methods

None

## API Specification

### Get Range Font 

Get Range's font information. This information is obtained by navigating to the font relation.

#### Syntax

```js
rangeObject.format.font;
```

#### Returns

* [Range Fill](rangefill.md) object.


#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var rangeFont = ramge.format.font;
ctx.load(rangeFont);
ctx.executeAsync().then(function() {
	Console.log(rangeFont.name);
});
```
[Back](#properties)

### Set Range Font 

Set range's font properties.

#### Syntax
```js
rangeObject.format.font.property = value;
```
Where, property is one of the following properties that can be set. 

#### Properties

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

#### Example
The example below sets font name. 

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.font.name = 'Times New Roman';
ctx.executeAsync().then();
```
[Back](#properties)
