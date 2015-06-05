# Range Fill

Represents the interior of an object, which includes fill formating information. 

## [Properties](#get-range-fill)
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`color`|String|HTML color code representation of the fill color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |Conversion from Range.Interior.Color value to html color string|

## Programming notes about `color` property: 

A `color` hex code is a way of specifying color using hexadecimal values. The code itself is a hex triplet, which represents three separate values that specify the levels of the component colors. The code starts with a pound sign (#) and is followed by six hex values or three hex value pairs (for example, #AFD645). 

Of the 6 Hex values, first two characters represent the values 0 through 255 for red in hex; the middle two for green and the last two for blue (#RRGGBB). For example, FF is equal to 255. Therefore, the purest white obtainable is the highest intensity of red, green and blue, which is #FFFFFF (red=255, green=255 and blue=255). Black is the lack of all RGB (#0000000).

When `color` value is updated, the input value needs to follow the appropriate formatting as mentioned above. The Alpha characters of the hex color code can be lower or upper case. 


## Relationships
None

## Methods
None

## API Specification

### Get Range Fill

Get Range's background/fill information. This information is obtained by navigating to the fill relation.

#### Syntax

```js
rangeObject.format.fill;
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
var rangeFill = ramge.format.fill;
ctx.load(rangeFill);
ctx.executeAsync().then(function() {
	Console.log(rangeFill.color);
});
```
[Back](#properties)

### Set Range Fill 

Set range background/fill properties.

#### Syntax
```js
rangeObject.format.fill.property = value;
```
Where, property is one of the following properties that can be set. 

#### Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------| 
|`color`|String|HTML color code representation of the fill color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |

#### Example
The example below sets fill color. 

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.fill.color = '0000FF';
ctx.executeAsync().then();
```

[Back](#properties)