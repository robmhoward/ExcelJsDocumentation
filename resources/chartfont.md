# ChartFont

This object represents the font attributes (font name, font size, color, etc.) for a chart object.

## [Properties](#setter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|bold|bool|Represents the bold status of font.||
|color|string|HTML color code representation of the text color. E.g. #FF0000 represents Red.||
|italic|bool|Represents the italic status of font.||
|name|string|Font name (e.g. "Calibri")||
|size|double|Size of the font (e.g. 11)||
|underline|string|Type of underline applied to the font. Possible values are: None, Single.||

## Relationships
None


## Methods
None


## API Specification

#### Setter Examples

Use chart title as an example.

```js
chartObject.title.format.font.name = "Calibri";
chartObject.title.format.font.size = 12;

[Back](#properties)
