# Font

This object represents the font attributes (font name, font size, color, and so on) for an object. 

## [Properties](#set-font)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`name`|String|Font name (e.g., "Calibri")|Range.Font.Name|
|`size`|number|Size of the font (e.g., 11)|Range.Font.Size|
|`color`|String|HTML color code representation of the text color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |Conversion from Range.Font.Color value to html color string|
|`italic`|Boolean|Represents the bold status of italic. true if the font style is italic|Range.Font.Italic|
|`bold`|Boolean|Represents the bold status of font. true if the font is bold. |Range.Font.Bold|
|`underline`|Boolean|Type of underline applied to the font. Can be one of the following constants. Possible Values: `None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting`|Range.Font.Underline|

## Relationships
None

## Methods
None.

## API Specification 

### Set Font

Update a chart font formatting.

#### Syntax
Use chart title as an example.
```js
chartObject.title.format.font.name = "Calibri";
chartObject.title.format.font.size = 12;
chartObject.title.format.font.color = "#FF0000";
chartObject.title.format.font.italic =  false;
chartObject.title.format.font.bold = true;
chartObject.title.format.font.underline = false;

```

#### Properties
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`name`|String|Font name (e.g., "Calibri")|
|`size`|number|Size of the font (e.g., 11)|Range.Font.Size|
|`color`|String|HTML color code representation of the text color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |
|`italic`|Boolean|Represents the bold status of italic. true if the font style is italic|
|`bold`|Boolean|Represents the bold status of font. true if the font is bold. |
|`underline`|Boolean|Type of underline applied to the font. Can be one of the following constants. Possible Values: `None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting`|

#### Returns

[ChartFont](resources/chartFont.md) object. 

#### Examples

##### Set chart title to be Calbri, size 10, bold and in red. 
```js
var ctx = new Excel.ExcelClientContext();
var title = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").title;

title.format.font.name = "Calibri";
title.format.font.size = 12;
title.format.font.color = "#FF0000";
title.format.font.italic =  false;
title.format.font.bold = true;
title.format.font.underline = false;

ctx.executeAsync().then(function () {
		logComment("Chart Title Font Updated");
});
```
[Back](#properties)
