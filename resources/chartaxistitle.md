# ChartAxisTitle

Represents the title of a chart axis.

## [Properties](#getter-and-setter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|text|string|Represents the axis title.||
|visible|bool|A boolean that specifies the visibility of an axis title.||

## Relationships
| Relationship | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|format|[ChartAxisTitleFormat](chartaxistitleformat.md)|Represents the formatting of chart axis title. Read-only.||

## Methods
None


## API Specification

#### Getter and Setter Examples
Get the `text` of Chart Axis Title from the value axis of Chart1.

```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var title = chart.axes.valueaxis.title;
ctx.load(title);
ctx.executeAsync().then(function () {
		Console.log(title.text);
});

Add "Values" as the title for the value Axis
```js
var ctx = new Excel.RequestContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.axes.valueaxis.title.text = "Values";

ctx.executeAsync().then(function () {
		Console.log("Axis Title Added ");
});
```

[Back](#properties)
