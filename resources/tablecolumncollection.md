# tablecolumn Collection
A collection of all the tablecolumn objects that are part of the table. 

## [Properties](#get-tablecolumn-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tablecolumns.count|
|`items`| [Table Column](tablecolumn.md) Array | A collection of all the tablecolumn objects that are part of the table|[tablecolumns.item] |

## Relationships

None

## Methods

The tablecolumn collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(index: number, values: array[][])](#addindex-number-values-array)| [Table Column](tablecolumn.md) Object             |Creates a new tablecolumn.  ||
|[getItem(param: string or number)](#getitemparam-string-or-number)| [Table Column](tablecolumn.md) Object     |Retrieve a tablecolumn object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [tablecolumn](tablecolumn.md) Object|Retrieve a tablecolumn based on its position in the items[] array.||


## API Specification 

### add(index: number, values: array[][])

Add a new column to the table. 

#### Syntax
```js
tableColumnCollection.add(index, values);
```

Parameter       | Type   | Description
--------------- | ------ | ------------
`values` | array[][] | Required. 2-D array of unformatted values of the table column.
`index` |  Number | Optional. Specifies the relative position of the new column. The previous column at this position is shifted outward to the right. If not specified, the addition happens at the end.  Note: The index value should be equal to or less than the last column's index value. In other words, this API cannot be used to append a column at the end of the table. **Zero Indexed**.

#### Returns
[Range](range.md) object.

#### Example
```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
var values = [["Sample"], ["Values"], ["For"], ["New"], ["Column"]];
var row = tables.getItem("Table1").tableColumns.add(null, values);
ctx.load(row);
ctx.executeAsync().then(function () {
	Console.log(row.name);
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
var wSheetName = 'Sheet1';
var tablecolumn = ctx.workbook.tablecolumns.getItem(wSheetName);
ctx.executeAsync().then(function () {
		Console.log(tablecolumn.index);
});
```
[Back](#methods)

### getItemAt(index: number)

Get tablecolumn object properties based on its position in the items[] array. 

#### Syntax
```js
tableColumnCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index or position in the items[]. Zero indexed.

#### Returns

[tablecolumn](tablecolumn.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var lastPosition = ctx.workbook.tablecolumns.count - 1;
var tablecolumn = ctx.workbook.tablecolumns.getItemAt(lastPosition);
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
var tablecolumns = ctx.workbook.tablecolumns;
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
var tablecolumns = ctx.workbook.tablecolumns;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("tablecolumns: Count= " + tablecolumns.count);
});

```
[Back](#properties)

