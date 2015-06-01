# Axix
Represents a single axis in a chart.

## [Properties](#get-chart-axis)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `minimum` | Object |Returns or sets the minimum value on the value axis. Auto if left empty.  | Axis.MinimumScale|
| `maximum` | Object |Returns or sets the maximum value on the value axis. Auto if left empty. | Axis.MaximumScale|
| `majorunit` | Object |Returns or sets the interval between two major tick marks. Auto if left empty.  | Axis.majorunit|
| `minorunit` | Object | Returns or sets the interval between two minor tick marks. Auto if left empty. | Axis.minorunit|


## Relationships
The Chart resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `title`          |[ChartAxisTitle](chartAxisTitle.md) Object | Represents the title of a specified axis. | Axis.AxisTitle
| `majorGridlines` | [ChartGridlines](chartGridlines.md) Object   | Returns a Gridlines object that represents the major gridlines for the specified axis.   | Axis.MajorGridlines|
| `minorGridlines` | [ChartGridlines](chartGridlines.md) Object   | Returns a Gridlines object that represents the minor gridlines for the specified axis.  | Axis.MinorGridlines|
| `format`          |[ChartAxisFormat](chartAxisFormat.md) Object | Represents the format of a chart object, which includes line/border and font formatting.

## Methods
None.

## API Specification 

### Get Chart Axis

Gets a ChartAxis object.

#### Syntax
Use value axis as an example here.

```js
chartObject.axes.valueaxis;
```

#### Properties
| Property         | Value    |Description|
|:-----------------|:--------|:----------|
| `minimum` | Object |Returns or sets the minimum value on the value axis. Auto if left empty.  | 
| `maximum` | Object |Returns or sets the maximum value on the value axis. Auto if left empty. | 
| `majorunit` | Object |Returns or sets the interval between two major tick marks. Auto if left empty.  | 
| `minorunit` | Object |eturns or sets the interval between two minor tick marks.  Auto if left empty. | 

#### Returns

[ChartAxis](resources/chartAxis.md) object. 

#### Examples

##### Get the `maximum` of Chart Axis from Chart1
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var axis = chart.axes.valueaxis;
ctx.load(axis);
ctx.executeAsync().then(function () {
		logComment(axis.maximum);
});
```

[Back](#properties)


### Set Chart Axis

 Set the  `maximum` ,  `minimum` ,  `majorunit` , `minorunit` of an axis. 

#### Syntax

```js
chartObject.axes.valueaxis.maximum = 5;
chartObject.axes.valueaxis.minimum = 0;
chartObject.axes.valueaxis.majorunit = 1;
chartObject.axes.valueaxis.minorunit = 0.2;
```

#### Properties
| Property         | Value    |Description|
|:-----------------|:--------|:----------|
| `minimum` | Object |Returns or sets the minimum value on the value axis. Auto if left empty.  | 
| `maximum` | Object |Returns or sets the maximum value on the value axis. Auto if left empty. | 
| `majorunit` | Object |Returns or sets the interval between two major tick marks. Auto if left empty.  | 
| `minorunit` | Object |eturns or sets the interval between two minor tick marks.  Auto if left empty. | 

#### Returns
[ChartAxis](resources/chartAxis.md) object. 

#### Examples

#####  Set the  `maximum`,  `minimum` ,  `majorunit` , `minorunit` of valueaxis. 
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.axes.valueaxis.maximum = 5;
chart.axes.valueaxis.minimum = 0;
chart.axes.valueaxis.majorunit = 1;
chart.axes.valueaxis.minorunit = 0.2;

ctx.executeAsync().then(function () {
		logComment("Axis Settings Changed");
});
```
[Back](#properties)