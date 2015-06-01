# Worksheet Collection
A collection of all the worksheet objects that are part of the workbook. 

## [Properties](#get-worksheet-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|Worksheets.count|
|`items`| [Worksheet](worksheet.md) Array | A collection of all the worksheet objects that are part of the workbook|[Worksheets.item] |

## Relationships

None

## Methods

The Worksheet collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(name: string)](#addname-string)| [Worksheet](worksheet.md) Object              |Creates a new worksheet. The new worksheet becomes the active workbook. ||
|[getItem(name: string)](#getitemname-string)| [Worksheet](worksheet.md) Object      |Retrieve a worksheet object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [Worksheet](worksheet.md) Object     |Retrieve a worksheet based on its position in the items[] array.||


## API Specification 

### Get Worksheet Collection

Get properties of the worksheet collection. 

#### Syntax
```js
context.workbook.worksheets.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|Worksheets.count|
|`items`| Object[] | A collection of all the worksheet objects that are part of the workbook|[Worksheets.item] |


#### Returns

[Worksheet](worksheet.md) collection. 

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var worksheets = ctx.workbook.worksheets;
ctx.load(worksheets);
ctx.executeAsync().then(function () {
	for (var i = 0; i < worksheets.items.length; i++)
	{
		Console.log(worksheets.items[i].name);
		Console.log(worksheets.items[i].index);
	}
});
```

##### Getting the number of worksheets

```js
var ctx = new Excel.ExcelClientContext();
var worksheets = ctx.workbook.worksheets;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("Worksheets: Count= " + worksheets.count);
});

```
[Back](#properties)

### add(name: string)

Add a new worksheet to the workbook. The worksheet will be added at the end of existing worksheets.

#### Syntax
```js
worksheetCollection.add(name);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`name`  | String| Optional. String value representing the name of the sheet to be added. If not specified, Excel determines the name of the new worksheet being added. 

#### Returns
[Worksheet](worksheet.md) object.

#### Examples

```js
var wSheetName = 'Sample Name';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.add(wSheetName);
ctx.load(worksheet);
ctx.executeAsync().then(function () {
	Console.log(worksheet.name);
});
```
[Back](#methods)

### getItem(name: string)

Get Worksheet object properties based on name.

#### Syntax
```js
worksheetCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. Worksheet name. 

#### Returns

[Worksheet](worksheet.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var wSheetName = 'Sheet1';
var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
ctx.executeAsync().then(function () {
		Console.log(worksheet.index);
});
```
[Back](#methods)


### getItemAt(index: number)

Get Worksheet object properties based on its position in the items[] array. 

#### Syntax
```js
worksheetCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index or position in the items[]. Zero indexed.

#### Returns

[Worksheet](worksheet.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var lastPosition = ctx.workbook.worksheets.count - 1;
var worksheet = ctx.workbook.worksheets.getItemAt(lastPosition);
ctx.executeAsync().then(function () {
		Console.log(worksheet.name);
});
```
[Back](#methods)
