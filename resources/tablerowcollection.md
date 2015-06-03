# tablerow Collection
A collection of all the tablerow objects that are part of the table. 

## [Properties](#get-tablerow-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tablerows.count|
|`items`| [Table Row](tablerow.md) Array | A collection of all the tablerow objects that are part of the table|[tablerows.item] |

## Relationships

None

## Methods

The tablerow collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(index: number, values: array[][])](#addindex-number-values-array)| [Table Row](tablerow.md) Object  |Creates a new tablerow. ||
|[getItem(name: string)](#getitemname-string)| [Table Row](tablerow.md) Object |Retrieve a tablerow object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [Table Row](tablerow.md) Object |Retrieve a tablerow based on its position in the items[] array.||


## API Specification 


### add(index: number, values: array[][])

Add a new row to the table. 

#### Syntax
```js
tableRowCollection.add(index, values);
```
#### Parameters 
Parameter       | Type   | Description
--------------- | ------ | ------------
`index` |  Number |Optional. Specifies the relative position of the new row. If not specified, the addition happens at the end. The previous column at this position is shifted outward to the bottom. **Zero Indexed**
`values` | Collection (primitive) | 2-D array of unformatted values of the table row. 

#### Returns
[Table Row](tableRow.md) object.

#### Example
```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
var values = [["Sample", "Values", "For", "New", "Row"]];
var row = tables.getItem("Table1").tablerows.add(null, values);
ctx.load(row);
ctx.executeAsync().then(function () {
	Console.log(row.index);
});
```
[Back](#methods)

### getItem(name: string)

Get tablerow object properties based on name.

#### Syntax
```js
tableRowCollection.getItem(name);
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
tableRowCollection.getItemAt(index);
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

### Get tablerow Collection

Get properties of the tablerow collection. 

#### Syntax
```js
tableRowCollection.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tablerows.count|
|`items`| [Table Row](tablerow.md) Array  | A collection of all the tablerow objects that are part of the table|[tablerows.item] |


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
