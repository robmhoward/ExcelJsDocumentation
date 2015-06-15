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
|[add(index: number, values: any[][])](#index-number-values-any)| [Table Row](tablerow.md) Object  |Creates a new tablerow. ||
|[getItemAt(index: number)](#getitematindex-number)| [Table Row](tablerow.md) Object |Retrieve a tablerow based on its position in the collection..||

## API Specification 


### add(index: number, values: any[][])

Add a new row to the table. 

#### Syntax
```js
tableRowCollection.add(index, values);
```
#### Parameters 
Parameter       | Type   | Description
--------------- | ------ | ------------
`index` |  Number |Optional. Specifies the relative position of the new row. If not specified, the addition happens at the end. The previous column at this position is shifted outward to the bottom. **Zero Indexed**
`values` | any[][] | 2-D array of unformatted values of the table row. 


#### Returns
[Table Row](tableRow.md) object.

#### Example
```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
var values = [["Sample", "Values", "For", "New", "Row"]];
var row = tables.getItem("Table1").rows.add(null, values);
ctx.load(row);
ctx.executeAsync().then(function () {
	Console.log(row.index);
});
```
[Back](#methods)

### getItemAt(index: number)

Get tablerow object properties based on its position in the collection.. 

#### Syntax
```js
tableRowCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index value of the object to be retrieved.. Zero indexed.

#### Returns

[tablerow](tablerow.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var tablerow = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0);
ctx.load(tablerow);
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
var tablerows = ctx.workbook.tables.getItem('Table1').rows;
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
var tablerow = ctx.workbook.tables.getItem('Table1').rows.getItemAt(0);
ctx.load(tablerows);
ctx.executeAsync().then(function () {
	Console.log("tablerows: Count= " + tablerows.count);
});

```
[Back](#properties)
