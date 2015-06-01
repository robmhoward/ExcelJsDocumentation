# Chart Series
Represents a series in a chart.

## [Properties](#set-chart-series)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`name`          |String|A String value that represents a Series object ||

## Relationships
The ChartSeries resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `points`          |[ChartPoints Collection](chartPointsCollection.md) | Represents Points in a series in a chart.
| `format`          |[ChartSeriesFormat](chartSeriesFormat.md) Object |  Represents the format of chart series, which includes fill(interior/background) and line formatting.

## Methods
None.

## API Specification
### Set Chart Series
Set properties of ChartSeries.

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