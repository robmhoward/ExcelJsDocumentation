# RangeFont

This object represents the font attributes (font name, font size, color, etc.) for an object.

## [Properties](#getter-and-setter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|bold|bool|Represents the bold status of font.||
|color|string|HTML color code representation of the text color. E.g. #FF0000 represents Red.||
|italic|bool|Represents the italic status of the font.||
|name|string|Font name (e.g. "Calibri")||
|size|double|Font size.||
|underline|string|Type of underline applied to the font. Possible values are: None, Single, Double, SingleAccountant, DoubleAccountant.||

## Relationships
None


## Methods
None


## API Specification

#### Getter and Setter Examples

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
The example below sets font name. 

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.font.name = 'Times New Roman';
ctx.executeAsync().then();
```

[Back](#properties)
