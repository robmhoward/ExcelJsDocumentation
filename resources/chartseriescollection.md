# ChartSeriesCollection

Represents a collection of chart series.

## [Properties](#getter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|count|int|Returns the number of series in the collection. Read-only.||
|items|[ChartSeriesCollection](chartseriescollection.md)|A collection of chartSeries objects. Read-only.||

## Relationships
None


## Methods

| Method           | Return Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|[getItemAt(index: number)](#getitematindex-number)|[ChartSeries](chartseries.md)|Retrieves a series based on its position in the collection||

## API Specification

### getItemAt(index: number)
Retrieves a series based on its position in the collection

#### Syntax
```js
chartSeriesCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[ChartSeries](chartseries.md)

#### Examples

Get the name of the first series in the series collection.
```js
var ctx = new Excel.RequestContext();
var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
ctx.load(seriesCollection);
ctx.executeAsync().then(function () {
	Console.log(seriesCollection.items[0].name);
});
```


[Back](#methods)

#### Getter Examples
Getting the names of series in the series collection.

```js
var ctx = new Excel.RequestContext();
var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
ctx.load(seriesCollection);
ctx.executeAsync().then(function () {
	for (var i = 0; i < seriesCollection.items.length; i++)
	{
		Console.log(seriesCollection.items[i].name);
	}
});
```

Get the number of chart series in collection.

```js
var ctx = new Excel.RequestContext();
var seriesCollection = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1").series;
ctx.load(seriesCollection);
ctx.executeAsync().then(function () {
	Console.log("series: Count= " + seriesCollection.count);
});

```


[Back](#properties)
