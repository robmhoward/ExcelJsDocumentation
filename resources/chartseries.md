# ChartSeries

Represents a series in a chart.

## [Properties](#setter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|name|string|Represents the name of a series in a chart.||

## Relationships
| Relationship | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|format|[ChartSeriesFormat](chartseriesformat.md)|Represents the formatting of a chart series, which includes fill and line formatting. Read-only.||
|points|[ChartPointsCollection](chartpointscollection.md)|Represents a collection of all points in the series. Read-only.||

## Methods
None


## API Specification

#### Setter Examples

Rename the 1st series of Chart1 to "New Series Name"

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.series.getItemAt(0).name = "New Series Name";

ctx.executeAsync().then(function () {
		Console.log("Series1 Renamed");
});
```

[Back](#properties)
