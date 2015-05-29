# chart Collection
A collection of all the chart objects that are part of the workbook. 

## [Properties](#get-chart-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|charts.count|
|`items`| Object[] | A collection of all the chart objects that are part of the workbook|[charts.item] |

## Relationships

None

## Methods

The chart resource has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(name: string)](#addname-string)| [chart](chart.md)              |Creates a new chart. The new chart becomes the active workbook. ||
|[getItem(name: string)](#getitemname-string)| [chart](chart.md)      |Retrieve a chart object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [chart](chart.md)     |Retrieve a chart based on its position in the items[] array.||


## API Specification 

### Get chart Collection

Get properties of the chart collection. 

#### Syntax
```js
context.workbook.charts;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|charts.count|
|`items`| Object[] | A collection of all the chart objects that are part of the workbook|[charts.item] |


#### Returns

[chart](chart.md) collection. 

#### Examples


#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var charts = ctx.workbook.charts;
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
var charts = ctx.workbook.charts;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("charts: Count= " + charts.count);
});

```
[Back](#properties)

### add(name: string)

Add a new chart to the workbook. The chart will be added at the end of existing charts.

#### Syntax
```js
chartsCollection.add(name);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`name`  | String| Optional. String value representing the name of the sheet to be added. If not specified, Excel determines the name of the new chart being added. 

#### Returns
[chart](../resources/chart.md) object.

#### Examples

```js
var wSheetName = 'Sample Name';
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.charts.add(wSheetName);
ctx.load(chart);
ctx.executeAsync().then(function () {
	Console.log(chart.name);
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

[chart](../resources/chart.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var wSheetName = 'Sheet1';
var chart = ctx.workbook.charts.getItem(wSheetName);
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

[chart](../resources/chart.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var lastPosition = ctx.workbook.charts.count - 1;
var chart = ctx.workbook.charts.getItemAt(lastPosition);
ctx.executeAsync().then(function () {
		Console.log(chart.name);
});
```
[Back](#methods)
