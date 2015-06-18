# Chart Point
Represents a point of a series in a chart.

## Properties(#get-chart-point)
| Properties    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `value`          | object | Returns the value of a chart point. Read-only.

## Relationships
The ChartPoint has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `format`          |[Chart Point Format](chartPointFormat.md) object | Represents the formatting of a chart point, which includes fill (interior/background) and line formatting.

## Methods
None.


## API Specification

### Get Chart Point
Get the properties of a chart, like value.

#### Syntax

```js
chartPointObject.value;
```

#### Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`value`          |object| Returns the value of a chart point. Read-only. |

#### Returns
object

#### Examples

##### Get the value of the the 1st Point in Seires1.

```js
var ctx = new Excel.ExcelClientContext();
var point = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points.getItemAt(0);	

ctx.load(point);

ctx.executeAsync().then(function () {
		logComment(point.value);
});
```
[Back](#properties)