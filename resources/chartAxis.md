# Chart Axis
Represents a single axis in a chart.

## [Properties](#get-chart-axis)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `minimum` | object |Returns or sets the minimum value on the value axis. Auto if left empty.  | Axis.MinimumScale|
| `maximum` | object |Returns or sets the maximum value on the value axis. Auto if left empty. | Axis.MaximumScale|
| `majorunit` | object |Returns or sets the interval between two major tick marks. Auto if left empty.  | Axis.majorunit|
| `minorunit` | object | Returns or sets the interval between two minor tick marks. Auto if left empty. | Axis.minorunit|


## Relationships
The Chart has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `title`          |[ChartAxisTitle](chartAxisTitle.md) object | Represents the title of a specified axis. | Axis.AxisTitle
| `majorGridlines` | [ChartGridlines](chartGridlines.md) object   | Returns a Gridlines object that represents the major gridlines for the specified axis.   | Axis.MajorGridlines|
| `minorGridlines` | [ChartGridlines](chartGridlines.md) object   | Returns a Gridlines object that represents the minor gridlines for the specified axis.  | Axis.MinorGridlines|
| `format`          |[ChartAxisFormat](chartAxisFormat.md) object | Represents the format of a chart object, which includes line/border and font formatting.

## Methods
None.

## API Specification 

### Get Chart Axis

Gets a ChartAxis object.

#### Syntax
Use value axis as an example here.

```js
chartObject.axes.axisTypeObject;
```

Where, axisTypeObject could be one of the following: 

| axis Type    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `categoryAxis` |[ChartAxis](chartAxis.md) object | Represents the category axis in a chart. | 
| `valueAxis` | [ChartAxis](chartAxis.md) object   | Represents the value axis in a chart.  | |
| `seriesAxis` | [ChartAxis](chartAxis.md) object   |Represents the series axis in a 3D chart. | |
     
#### Properties
| Property         | Value    |Description|
|:-----------------|:--------|:----------|
| `minimum` | object |Returns or sets the minimum value on the value axis. Auto if left empty.  | 
| `maximum` | object |Returns or sets the maximum value on the value axis. Auto if left empty. | 
| `majorunit` | object |Returns or sets the interval between two major tick marks. Auto if left empty.  | 
| `minorunit` | object |eturns or sets the interval between two minor tick marks.  Auto if left empty. | 

#### Returns

[ChartAxis](chartAxis.md) object. 

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

 Set the  `maximum`,  `minimum`,  `majorunit`, and `minorunit` of an axis. 

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
| `minimum` | object |Returns or sets the minimum value on the value axis. Auto if left empty.  | 
| `maximum` | object |Returns or sets the maximum value on the value axis. Auto if left empty. | 
| `majorunit` | object |Returns or sets the interval between two major tick marks. Auto if left empty.  | 
| `minorunit` | object |eturns or sets the interval between two minor tick marks.  Auto if left empty. | 

#### Returns
[ChartAxis](chartAxis.md) object. 

#### Examples

#####  Set the  `maximum`,  `minimum`,  `majorunit`, `minorunit` of valueaxis. 
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