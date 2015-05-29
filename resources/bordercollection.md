# Border Collection 

Represents the border objects that make up Range border. 


## [Properties](#get-border-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|range.borders.count|
|`items`| Object[] | A collection of all the border objects that are part of the workbook|ListObjects |

## Relationships

None

## Methods

The border collection resource has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[getItem(name: string)](#getitemname-string)| [border](border.md)      |Retrieve a border object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [border](border.md)     |Retrieve a border based on its position in the items[] array.||


## API Specification 

### Get border Collection

Get properties of the border collection. 

#### Syntax
```js
rangeObject.borders.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|borders.count|
|`items`| Object[] | A collection of all the border objects that are part of the workbook|[borders.item] |


#### Returns

[border](border.md) collection. 

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var borders = range.format.borders;
ctx.load(borders);
ctx.executeAsync().then(function () {
	for (var i = 0; i < borders.items.length; i++)
	{
		Console.log(borders.items[i].sideIndex);
	}
});
```

##### Getting the number of borders

```js
var sheetName = "Sheet1";
var rangeAddress = "D5:F8";
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
var borders = range.format.borders;
ctx.load(borders);
ctx.executeAsync().then(function () {
	Console.log(borders.count);
});
```
[Back](#properties)

### add(name: string)

Add a new border to the workbook. The border will be added at the end of existing borders.

#### Syntax
```js
bordersCollection.add(name);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`name`  | String| Optional. String value representing the name of the sheet to be added. If not specified, Excel determines the name of the new border being added. 

#### Returns
[border](../resources/border.md) object.

#### Examples

```js
var wSheetName = 'Sample Name';
var ctx = new Excel.ExcelClientContext();
var border = ctx.workbook.borders.add(wSheetName);
ctx.load(border);
ctx.executeAsync().then(function () {
	Console.log(border.name);
});
```
[Back](#methods)

### getItem(name: string)

Get border object properties based on name.

#### Syntax
```js
bordersCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. border name. 

#### Returns

[border](../resources/border.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var borderName = 'border1';
var border = ctx.workbook.borders.getItem(borderName);
ctx.executeAsync().then(function () {
		Console.log(border.index);
});
```
[Back](#methods)


### getItemAt(index: number)

Get border object properties based on its position in the items[] array. 

#### Syntax
```js
bordersCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index or position in the items[]. Zero indexed.

#### Returns

[border](../resources/border.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var border = ctx.workbook.borders.getItemAt(0);
ctx.executeAsync().then(function () {
		Console.log(border.name);
});
```
[Back](#methods)
