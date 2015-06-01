# Chart
Represents a chart object in a workbook.

## [Properties](#get-chart)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `name`  | String | A String value that represents the name of a Chart object.   | Chart.Name      |
| `height`| Number | A Number value that represents the height, in points, of the chart object. | ChartArea.Height|
| `width` | Number | A Number value that represents the width, in points, of the chart object. | ChartArea.Width |
| `top` | Number |a Number value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).| ChartArea.Top |
| `left` | Number | a Number value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart).| ChartArea.Left |


## Relationships
The Chart resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `title`          |[ChartTitle](chartTitle.md) Object | Returns a ChartTitle object that represents the title of the specified chart, including the text, visibility, position and formating of the title.
| `series`          |[ChartSeries](chartSeries.md) Object |Represents a series in a chart.
| `axes`          |[ChartAxes](chartAxes.md) Object |Represents a collection of Axes in the Chart.
| `dataLabels`          |[ChartDataLabels](chartDataLabels.md) Object | Represents the datalabels on the chart.
| `legend`          |[ChartLegend](chartLegend.md) Object | Returns a Legend object that represents the legend for the chart. 
| `format`          |[ChartFormat](chartFormat.md) Object | Represents the format of a chart object, which includes fill(interior/background), line/border and font formatting.

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[delete()](#delete)| void     |Deletes the Chart ||
|[setData(sourceData: string, seriesBy: string)](#setdatasourcedata-string-seriesby-string)| [Chart](Chart.md)  object |Sets the sourceData and seriesBy of a Chart.          

## API Specification 

### delete()

Deletes a chart directly on the grid.

#### Syntax

```js
chart.delete();
```

#### Parameters
None.

#### Returns

Nothing.

#### Examples

##### Delete the Chart named "Chart1"

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.delete();
ctx.executeAsync().then(function () {
		logComment"Chart Deleted");
});
```
[Back](#methods)


### setData(sourceData: string, seriesBy: string)

Sets the sourceData and seriesBy of a Chart.

#### Syntax

```js
chartObject.setData(sourceData, seriesBy);
```

#### Parameters
| Parameter         | Value    |Description|
|:-----------------|:--------|:----------|
| `sourceData`  | String|  Sets an address or name of the Range object as the data source.|
| `seriesBy`  | String |  Sets the way columns or rows are used as data series on the chart. Can be one of the following `Rows`, `Columns` or `Auto`.|

#### Returns

[Chart](chart.md) object. 

#### Examples

##### Set the `sourceData` to be "A1:B4" and `seriesBy` to be "Columns"
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	
var sourceData = "A1:B4";

chart.setData(sourceData, "Columns");
ctx.executeAsync().then();
```
[Back](#methods)

### Get Chart

Gets a chart object by name.

#### Syntax
```js
chartsCollection.getItem(name);	
```

#### Parameters
None.

#### Returns

[Chart](chart.md) object. 

#### Examples

##### Get the Chart named "Chart1"
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

ctx.load(chart);
ctx.executeAsync().then(function () {
		logComment("Chart1 Loaded");
});
```

[Back](#properties)

### Set Chart

Update a chart including renaming, positioning and resizing.

#### Syntax

```js
chartObject.name="New Name";
chartObject.top = 100;
chartObject.left = 100;
chartObject.height = 200;
chartObject.weight = 200;
```

#### Properties
| Property         | Value    |Description|
|:-----------------|:--------|:----------|
| `name`  | String|A String value that represents the name of a Chart object                              |
| `height`|  Number |Returns or sets a Number value that represents the height, in points, of the object |
| `width` |  Number |Returns or sets a Number value that represents the width, in points, of the object. | 
| `top` |  Number |Returns or sets a Number value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
| `left` |  Number |Returns or sets a Number value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart). | 

#### Returns

[Chart](chart.md) object. 

#### Examples

##### Rename the chart to new name, resize the chart to 200 points in both height and weight. Move Chart1 to 100 points to the top and left. 
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");

chart.name="New Name";	
chart.top = 100;
chart.left = 100;
chart.height =200;
chart.width =200;
ctx.executeAsync().then(function () {
		logComment("Chart Updated");
});
```
[Back](#properties)
