# Table Column Collection
Represents a collection of all the columns that are part of the table.  

## [Properties](#get-tablecolumn-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Returns the number of columns in the table.|tablecolumns.count|
|`items`| [Table Column](tablecolumn.md) Array | Returns a collection of all the column objects that are part of the table.|[tablecolumns.item] |

## Relationships

None

## Methods

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(index: number, values: any[][])](#index-number-values-any)| [Table Column](tablecolumn.md) object             |Adds a new column to the table.  ||
|[getItem(param: string or number)](#getitemparam-string-or-number)| [Table Column](tablecolumn.md) object     |Gets a column object by name or id.||
|[getItemAt(index: number)](#getitematindex-number)| [tablecolumn](tablecolumn.md) object|Gets a column based on its position in the collection.||


## API Specification 

### add(index: number, values: any[][])

Add a new column to the table. 

#### Syntax
```js
tableColumnCollection.add(index, values);
```

Parameter       | Type   | Description
--------------- | ------ | ------------
`index` |  Number | Optional. Specifies the relative position of the new column. The previous column at this position is shifted outward to the right. If not specified, the addition happens at the end.  Note: The index value should be equal to or less than the last column's index value. In other words, this API cannot be used to append a column at the end of the table. **Zero-indexed**.
`values` | any[][] | Required. A 2-dimensional array of unformatted values of the table column.


#### Returns
[Range](range.md) object.

#### Example
```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
var column = tables.getItem("Table1").columns.add(null, values);
ctx.load(column);
ctx.executeAsync().then(function () {
	Console.log(column.name);
});
```
[Back](#methods)

### getItem(param: string or number)

Get tablecolumn object properties based on name.

#### Syntax
```js
tableColumnCollection.getItem(param);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `param`| String | Required. tablecolumn name or id. 

#### Returns

[tablecolumn](tablecolumn.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItem(0);
ctx.load(tablecolumn)
ctx.executeAsync().then(function () {
		Console.log(tablecolumn.name);
});
```
[Back](#methods)

### getItemAt(index: number)

Get tablecolumn object properties based on its position in the collection. 

#### Syntax
```js
tableColumnCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index value of the object to be retrieved. Zero-indexed.

#### Returns

[tablecolumn](tablecolumn.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var tablecolumn = ctx.workbook.tables.getItem['Table1'].columns.getItemAt(0);
ctx.load(tablecolumn)
ctx.executeAsync().then(function () {
		Console.log(tablecolumn.name);
});
```
[Back](#methods)

### Get tablecolumn Collection

Get properties of the tablecolumn collection. 

#### Syntax
```js
tableColumnCollection.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tablecolumns.count|
|`items`| [Table Column](tablecolumn.md) Array | A collection of all the tablecolumn objects that are part of the table|[tablecolumns.item] |

#### Returns

[tablecolumn](tablecolumn.md) collection. 

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
ctx.load(tablecolumns);
ctx.executeAsync().then(function () {
	for (var i = 0; i < tablecolumns.items.length; i++)
	{
		Console.log(tablecolumns.items[i].name);
	}
});
```

##### Getting the number of tablecolumns

```js
var ctx = new Excel.ExcelClientContext();
var tablecolumns = ctx.workbook.tables.getItem['Table1'].columns;
ctx.load(tablecolumns);
ctx.executeAsync().then(function () {
	Console.log("tablecolumns: Count= " + tablecolumns.count);
});

```
[Back](#properties)

