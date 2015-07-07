# ChartLegend

Represents the legend in a chart.

## [Properties](#getter-and-setter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|overlay|bool|Boolean value for whether the chart legend should overlap with the main body of the chart.||
|position|string|Represents the position of the legend on the chart. Possible values are: Top, Bottom, Left, Right, Corner, Custom.||
|visible|bool|A boolean value the represents the visibility of a ChartLegend object.||

## Relationships
| Relationship | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|format|[ChartLegendFormat](chartlegendformat.md)|Represents the formatting of a chart legend, which includes fill and font formatting. Read-only.||

## Methods
None


## API Specification

#### Getter and Setter Examples

Get the `position` of Chart Legend from Chart1

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var legend = chart.legend;
ctx.load(legend);
ctx.executeAsync().then(function () {
		Console.log(legend.position);
});
```

Set to show legend of Chart1 and make it on top of the chart.

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.legend.visible = true;
chart.legend.position = "top"; 
chart.legend.overlay = false; 
ctx.executeAsync().then(function () {
		Console.log("Legend Shown ");
});
``` 
[Back](#properties)
