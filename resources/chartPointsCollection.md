# Chart Points Collection
A collection of all the ChartPoints objects of a chart. 

## [Properties](#get-chartpoints-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.||
|`items`| [Chart Points](chartPoints.md) Array | A collection of all the chart objects that are part of the workbook| |

## Relationships

None

## Methods

The chart has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[getItemAt(index: number)](#getitematindex-number)| [chartPoints](chartPoints.md)     |Retrieve a ChartPoints Object based on its position in the items[] array.||


## API Specification 

### Get ChartPoints Collection

Get the ChartPoints collection. 

#### Syntax
```js
context.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;	
```

#### Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`count`| Number   | Number of objects in the collection.|
|`items`| Object[] | A collection of all the chart objects that are part of the workbook|

#### Returns

[ChartPoints](chartPoints.md) collection. 

#### Examples

##### Getting the names of points in the points collection
```js
var ctx = new Excel.ExcelClientContext();
var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
ctx.load(pointsCollection);
ctx.executeAsync().then(function () {
	Console.log("Points Collection loaded");
});
```

##### Getting the number of points

```js
var ctx = new Excel.ExcelClientContext();
var pointsCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").points;
ctx.load(pointsCollection);
ctx.executeAsync().then(function () {
	Console.log("points: Count= " + pointsCollection.count);
});

```
[Back](#properties)


### getItemAt(index: number)

Get chartPoints object properties based on its position in the items[] array. 

#### Syntax
```js
ChartPointsCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index or position in the items[]. Zero indexed.

#### Returns

[chartPoints](../resources/chartPoints.md) object.

#### Examples

##### set the border color for the first points in the points collection
```js
var ctx = new Excel.ExcelClientContext();
var point = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series.getItemAt(0).points;
points.getItemAt(0).format.fill.setSolidColor("8FBC8F");
ctx.executeAsync().then(function () {
	Console.log("Point Border Color Changed");
});
```
[Back](#methods)
