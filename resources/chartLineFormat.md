# ChartLineFormat
Provides access to the Office Art formatting for chart line elements.

## [Properties](#set-line)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|color| String | HTML color code representing the color of lines in the chart. |Chart.ChartFormat.LineFormat.Color's representation in HTML color code.|

## Relationships
None.

## Methods
| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[clear()](#clear)|void |Clear the line format of a chart element.

## API Specification 
### clear()

Clear the line format of a chart eleme

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

##### Clear the line format of the major Gridlines on value axis of the Chart named "Chart1"

```js
var ctx = new Excel.ExcelClientContext();
var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").axes.valueaxis.majorGridlines;	

gridlines.format.line.clear();
ctx.executeAsync().then(function () {
		logComment"Chart Major Gridlines Format Cleared");
});
```
[Back](#methods)

### Set Line

Update a chart line formatting.

#### Syntax
Use chart major gridlines on value axis as an example.
```js
GridlinesObject.format.line.color = "#FF0000";

```

#### Properties
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`color`|String|HTML color code representation of the text color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |


#### Returns

[ChartLineFormat](chartLineFormat.md) object. 

#### Examples

##### Set chart major gridlines on value axis to be red.
```js
var ctx = new Excel.ExcelClientContext();
var gridlines = ctx.workbook.worksheets.getItem("Sheet1").charts.axes.valueaxis.majorGridlines;

gridlines.format.line.color = "#FF0000";


ctx.executeAsync().then(function () {
		logComment("Chart Gridlines Color Updated");
});
```
[Back](#properties)