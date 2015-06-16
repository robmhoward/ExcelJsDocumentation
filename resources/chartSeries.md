# Chart Series
Represents a series in a chart.

## [Properties](#set-chart-series)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`name`          |String|Returns or sets the name of a series in a chart. ||

## Relationships

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `points`          |[ChartPoint Collection](chartPointsCollection.md) | Represents an object that returns a single point or a collection of all the points in the series. 
| `format`          |[ChartSeriesFormat](chartSeriesFormat.md) Object |  Represents the formatting of a chart series, which includes fill(interior/background) and line formatting.

## Methods
None.

## API Specification

### Set Chart Series
Set the properties of a series in a chart.

#### Syntax

```js
chartObject.series.getItemAt(0);
```

#### Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`name`          |String|A String value that represents a Series object |

#### Returns
[ChartSeries](chartSeries.md) object. 

#### Examples

##### Rename the 1st series of Chart1 to "New Series Name"

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.series.getItemAt(0).name = "New Series Name";

ctx.executeAsync().then(function () {
		logComment("Series1 Renamed");
});
```
[Back](#properties)