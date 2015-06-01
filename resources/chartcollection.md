# Chart Collection
A collection of all the chart objects on a worksheet. 

## [Properties](#get-chart-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|charts.count|
|`items`| [Chart](chart.md) Array| A collection of all the chart objects that are part of the workbook|[charts.item] |

## Relationships

None

## Methods

The chart has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(type: string, sourceData: any, seriesBy: string)](#addtype-string-sourceData-any-seriesBy-string)| [chart](chart.md)              |Creates a new chart. The new chart becomes the active workbook. ||
|[getItem(name: string)](#getitemname-string)| [chart](chart.md)      |Retrieve a chart object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [chart](chart.md)     |Retrieve a chart based on its position in the items[] array.||


## API Specification 

### Get chart Collection

Get the chart collection. 

#### Syntax
```js
context.workbook.worksheets.getItem("Sheet1").charts;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|charts.count|
|`items`| Object[] | A collection of all the chart objects that are part of the workbook|[charts.item] |


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

### add(type: string, sourceData: any, seriesBy: string)

Add a new chart to the workbook. The chart will be added at the end of existing charts.

#### Syntax
```js
chartsCollection.add(type, sourceData, seriesBy);
```

#### Parameters

| Parameter         | Value    |Description|
|:-----------------|:--------|:----------|
| `type` | String | A String value that represents the type of a chart.  |
| `sourceData`  | String | A String that represents an address or name of the Range object as the data source.|
| `seriesBy` | String |  A String that represents the way columns or rows are used as data series on the chart. Can be `auto`, `Rows` or `Columns`.|

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

Get chart object properties based on name.

#### Syntax
```js
chartsCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. chart name. 

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

Get chart object properties based on its position in the items[] array. 

#### Syntax
```js
chartsCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index or position in the items[]. Zero indexed.

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
