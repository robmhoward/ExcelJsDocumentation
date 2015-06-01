# tablerow Collection
A collection of all the tablerow objects that are part of the table. 

## [Properties](#get-tablerow-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tablerows.count|
|`items`| Object[] | A collection of all the tablerow objects that are part of the table|[tablerows.item] |

## Relationships

None

## Methods

The tablerow collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(name: string)](#addname-string)| [tablerow](tablerow.md)              |Creates a new tablerow. ||
|[getItem(name: string)](#getitemname-string)| [tablerow](tablerow.md)      |Retrieve a tablerow object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [tablerow](tablerow.md)     |Retrieve a tablerow based on its position in the items[] array.||


## API Specification 

### Get tablerow Collection

Get properties of the tablerow collection. 

#### Syntax
```js
context.workbook.tablerows.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tablerows.count|
|`items`| Object[] | A collection of all the tablerow objects that are part of the table|[tablerows.item] |


#### Returns

[tablerow](tablerow.md) collection. 

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var tablerows = ctx.workbook.tablerows;
ctx.load(tablerows);
ctx.executeAsync().then(function () {
	for (var i = 0; i < tablerows.items.length; i++)
	{
		Console.log(tablerows.items[i].name);
		Console.log(tablerows.items[i].index);
	}
});
```

##### Getting the number of tablerows

```js
var ctx = new Excel.ExcelClientContext();
var tablerows = ctx.workbook.tablerows;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("tablerows: Count= " + tablerows.count);
});

```
[Back](#properties)

### add(name: string)

Add a new tablerow to the workbook. The tablerow will be added at the end of existing tablerows.

#### Syntax
```js
tablerowsCollection.add(name);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`name`  | String| Optional. String value representing the name of the sheet to be added. If not specified, Excel determines the name of the new tablerow being added. 

#### Returns
[tablerow](tablerow.md) object.

#### Examples

```js
var wSheetName = 'Sample Name';
var ctx = new Excel.ExcelClientContext();
var tablerow = ctx.workbook.tablerows.add(wSheetName);
ctx.load(tablerow);
ctx.executeAsync().then(function () {
	Console.log(tablerow.name);
});
```
[Back](#methods)

### getItem(name: string)

Get tablerow object properties based on name.

#### Syntax
```js
tablerowsCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. tablerow name. 

#### Returns

[tablerow](tablerow.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var wSheetName = 'Sheet1';
var tablerow = ctx.workbook.tablerows.getItem(wSheetName);
ctx.executeAsync().then(function () {
		Console.log(tablerow.index);
});
```
[Back](#methods)


### getItemAt(index: number)

Get tablerow object properties based on its position in the items[] array. 

#### Syntax
```js
tablerowsCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index or position in the items[]. Zero indexed.

#### Returns

[tablerow](tablerow.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var lastPosition = ctx.workbook.tablerows.count - 1;
var tablerow = ctx.workbook.tablerows.getItemAt(lastPosition);
ctx.executeAsync().then(function () {
		Console.log(tablerow.name);
});
```
[Back](#methods)
