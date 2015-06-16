# Chart Point Collection
A collection of all the ChartPoint objects of a chart. 

## [Properties](#get-chartpoint-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.||
|`items`| [Chart Point](chartPoint.md) Array | A collection of all the chart objects that are part of the workbook| |

## Relationships

None

## Methods

The chart has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[getItemAt(index: number)](#getitematindex-number)| [ChartPoint](chartPoint.md)     |Retrieve a ChartPoint object based on its position in the collection..||


## API Specification 

### Get ChartPoint Collection

Get the ChartPoint Collection. 

#### Syntax
```js
chartObject.points;	
```

#### Properties

| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`count`| Number   | Number of objects in the collection.|
|`items`| object[] | A collection of all the chart objects that are part of the workbook|

#### Returns

[ChartPoint](chartPoint.md) collection. 

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

Get chartPoint object properties based on its position in the collection.. 

#### Syntax
```js
ChartPointsCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index value of the object to be retrieved.. Zero indexed.

#### Returns

[ChartPoint](chartPoint.md) object.

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
