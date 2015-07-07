# RangeBorderCollection

Represents the border objects that make up range border.

## [Properties](#getter-and-setter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|count|int|Number of border objects in the collection. Read-only.||
|items|[RangeBorderCollection](rangebordercollection.md)|A collection of rangeBorder objects. Read-only.||

## Relationships
None


## Methods

| Method           | Return Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|[getItem(index: string)](#getitemindex-string)|[RangeBorder](rangeborder.md)|Gets a border object using its name||
|[getItemAt(index: number)](#getitematindex-number)|[RangeBorder](rangeborder.md)|Gets a border object using its index||

## API Specification

### getItem(index: string)
Gets a border object using its name

#### Syntax
```js
rangeBorderCollectionObject.getItem(index);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|index|string|Index value of the border object to be retrieved.  Possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight, InsideVertical, InsideHorizontal, DiagonalDown, DiagonalUp|

#### Returns
[RangeBorder](rangeborder.md)

#### Examples
```js
var sheetName = "Sheet1";
var rangeAddress = "A1:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var borderName = 'EdgeTop';
var border = range.format.borders.getItem(borderName);
ctx.load(border);
ctx.executeAsync().then(function () {
		Console.log(border.style);
});
```


[Back](#methods)

### getItemAt(index: number)
Gets a border object using its index

#### Syntax
```js
rangeBorderCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[RangeBorder](rangeborder.md)

#### Examples
```js
var sheetName = "Sheet1";
var rangeAddress = "A1:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var border = ctx.workbook.borders.getItemAt(0);
ctx.load(border);
ctx.executeAsync().then(function () {
		Console.log(border.sideIndex);
});
```


[Back](#methods)

#### Getter and Setter Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var borders = range.format.borders;
ctx.load(borders);
ctx.executeAsync().then(function () {
	Console.log(borders.count);
	for (var i = 0; i < borders.items.length; i++)
	{
		Console.log(borders.items[i].sideIndex);
	}
});
```
The example below adds grid border around the range.

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:F8";
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
range.format.borders('InsideHorizontal').lineStyle = 'Continuous';
range.format.borders('InsideVertical').lineStyle = 'Continuous';
range.format.borders('EdgeBottom').lineStyle = 'Continuous';
range.format.borders('EdgeLeft').lineStyle = 'Continuous';
range.format.borders('EdgeRight').lineStyle = 'Continuous';
range.format.borders('EdgeTop').lineStyle = 'Continuous';
ctx.executeAsync().then();
```
[Back](#properties)
