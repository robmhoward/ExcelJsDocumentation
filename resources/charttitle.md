# ChartTitle

Represents a chart title object of a chart.

## [Properties](#getter-and-setter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|overlay|bool|Boolean value representing if the chart title will overlay the chart or not.||
|text|string|Represents the title text of a chart.||
|visible|bool|A boolean value the represents the visibility of a chart title object.||

## Relationships
| Relationship | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|format|[ChartTitleFormat](charttitleformat.md)|Represents the formatting of a chart title, which includes fill and font formatting. Read-only.||

## Methods
None


## API Specification

#### Getter and Setter Examples

Get the `text` of Chart Title from Chart1.

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var title = chart.title;
ctx.load(title);
ctx.executeAsync().then(function () {
		Console.log(title.text);
});
```

Set the `text` of Chart Title to "My Chart" and Make it show on top of the chart without overlaying.

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

ctx.executeAsync().then(function () {
		Console.log("Char Title Changed");
});
```

[Back](#properties)
