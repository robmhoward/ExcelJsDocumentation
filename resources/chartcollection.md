# Chart Collection
A collection of all the chart objects on a worksheet. 

## [Properties](#get-chart-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Returns the number of charts in the collection.|charts.count|
|`items`| [Chart](chart.md) Array| Returns a collection of all the charts in a workbook. |[charts.item] |

## Relationships

None

## Methods

The chart has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(type: string, sourceData: string, seriesBy: string)](#addtype-string-sourcedata-string-seriesby-string)| [Chart](chart.md) object              |Creates a new chart. ||
|[getItem(name: string)](#getitemname-id)| [Chart](chart.md) object     |Gets a chart using its name.||
|[getItemAt(index: number)](#getitematindex-number)| [Chart](chart.md) object    |Gets a chart based on its position in the collection.||


## API Specification 

### Get Chart Collection

Get the chart collection. 

#### Syntax
```js
worksheetObject.charts;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|charts.count|
|`items`| object[] | A collection of all the chart objects that are part of the workbook|[charts.item] |


#### Returns

[chart](chart.md) collection. 

#### Examples


```js
var ctx = new Excel.ExcelClientContext();
var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
ctx.load(charts);
ctx.executeAsync().then(function () {
	for (var i = 0; i < charts.items.length; i++)
	{
		Console.log(charts.items[i].name);
		Console.log(charts.items[i].index);
	}
});
```

##### Getting the number of charts

```js
var ctx = new Excel.ExcelClientContext();
var charts = ctx.workbook.worksheets.getItem("Sheet1").charts;
ctx.load(charts);
ctx.executeAsync().then(function () {
	Console.log("charts: Count= " + charts.count);
});

```
[Back](#properties)

### add(type: string, sourceData: string, seriesBy: string)

Add a new chart to the workbook. 

#### Syntax
```js
chartsCollection.add(type, sourceData, seriesBy);
```

#### Parameters

| Parameter         | Value    |Description|
|:-----------------|:--------|:----------|
| `type` | String | Represents the type of a chart.  |
| `sourceData`  | String | The address or name of the range that contains the source data.|
| `seriesBy` | String |  Specifies the way columns or rows are used as data series on the chart. Can be one of the following: `Rows`, `Columns` or `Auto`.|

#### Returns
[chart](chart.md) object.

#### Examples

##### Add a chart of `chartType` "ColumnClustered" on worksheet "Charts" with `sourceData` from Range "A1:B4" and `seriresBy` is set to be "auto".

```js
var sheetName = "Sheet1";
var sourceData = sheetName + "!" + "A1:B4";
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
ctx.executeAsync().then(function () {
		logComment("New Chart Added");
});
```
[Back](#methods)

### getItem(name: string)

Gets a chart object based on its name.

#### Syntax
```js
chartsCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. If of the chart to be retrieved. 

#### Returns

[chart](chart.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var chartname = 'Chart1';
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem(chartname);
ctx.executeAsync().then(function () {
		Console.log(chart.index);
});
```
[Back](#methods)


### getItemAt(index: number)

Gets a chart object based on its position in the collection. 

#### Syntax
```js
chartsCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index value of the object to be retrieved. Zero-indexed.

#### Returns

[chart](chart.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var lastPosition = ctx.workbook.worksheets.getItem("Sheet1").charts.count - 1;
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItemAt(lastPosition);
ctx.executeAsync().then(function () {
		Console.log(chart.name);
});
```
[Back](#methods)
