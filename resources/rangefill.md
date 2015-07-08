# RangeFill

Represents the background of a range object.

## [Properties](#getter-and-setter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|color|string|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange")||

## Relationships
None


## Methods

| Method           | Return Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|[clear()](#clear)|void|Resets the range background.||

## API Specification

### clear()
Resets the range background.

#### Syntax
```js
rangeFillObject.clear();
```

#### Parameters
None

#### Returns
void

#### Examples

This example resets the range background.

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var rangeFill = ramge.format.fill;
rangeFill.clear();
ctx.executeAsync();
```


[Back](#methods)

#### Getter and Setter Examples
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
The example below sets fill color. 

```js
var sheetName = "Sheet1";
var rangeAddress = "F:G";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.fill.color = '0000FF';
ctx.executeAsync();
```
[Back](#properties)
