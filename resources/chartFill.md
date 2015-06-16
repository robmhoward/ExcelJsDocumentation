# Chart Fill Format
Represents the fill formatting for a chart element.

## Properties
None.

## Relationships
None

## Methods

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[setSolidColor(color: string)](#setsolidcolorcolor-string)|void |Sets the fill formatting of a chart element to a uniform color.
|[clear()](#clear)|void |Clear the fill format of a chart element.



### setSolidColor(color: string)

Sets the fill formatting of a chart element to a uniform color.

#### Syntax
```js
ChartObject.format.fill.setSolidColor("#FF0000");	
```

#### Parameters
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`color`|String|HTML color code representation of the text color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |


#### Returns
None.

#### Examples

##### Set BackGround Color of Chart1 to be red.
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.format.fill.setSolidColor("#FF0000");

ctx.executeAsync().then(function () {
		logComment("Chart1 Background Color Changed.");
});
```
[Back](#methods)

### clear()

Clear the fill format of a chart element.

#### Syntax
Use chart major gridlines on value axis as an example.
```js
GridlinesObject.format.line.clear();
```

#### Parameters
None.

#### Returns

Nothing.

#### Examples

Clear the line format of the major Gridlines on value axis of the Chart named "Chart1"

```js
var ctx = new Excel.ExcelClientContext();
var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;	

gridlines.format.line.clear();
ctx.executeAsync().then(function () {
		logComment"Chart Major Gridlines Format Cleared");
});
```
[Back](#methods)