# Table Collection

A collection of all the table objects that are part of the workbook. 

## [Properties](#get-table-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|ListObjects.count|
|`items`| [Table](table.md) Array | A collection of all the table objects that are part of the workbook|[ListObjects.item] |

## Relationships

None

## Methods

The table collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(name: string)](#addname-string)| [Table](table.md) Object              |Creates a new table. ||
|[getItem(name: string)](#getitemname-string)| [Table](table.md) Object      |Retrieve a table object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [Table](table.md) Object     |Retrieve a table based on its position in the items[] array.||


## API Specification 

### Get table Collection

Get properties of the table collection. 

#### Syntax
```js
context.workbook.tables.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tables.count|
|`items`| Object[] | A collection of all the table objects that are part of the workbook|[tables.item] |


#### Returns

[table](table.md) collection. 

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
ctx.load(tables);
ctx.executeAsync().then(function () {
	for (var i = 0; i < tables.items.length; i++)
	{
		Console.log(tables.items[i].name);
	}
});
```

##### Getting the number of tables

```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("tables: Count= " + tables.count);
});

```
[Back](#properties)

### add(name: string)

Add a new table to the workbook. The table will be added at the end of existing tables.

#### Syntax
```js
tablesCollection.add(name);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`name`  | String| Optional. String value representing the name of the sheet to be added. If not specified, Excel determines the name of the new table being added. 

#### Returns
[table](table.md) object.

#### Examples

```js
var wSheetName = 'Sample Name';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.add(wSheetName);
ctx.load(table);
ctx.executeAsync().then(function () {
	Console.log(table.name);
});
```
[Back](#methods)

### getItem(name: string)

Get table object properties based on name.

#### Syntax
```js
tablesCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. table name. 

#### Returns

[table](table.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var tableName = 'Table1';
var table = ctx.workbook.tables.getItem(tableName);
ctx.executeAsync().then(function () {
		Console.log(table.index);
});
```
[Back](#methods)


### getItemAt(index: number)

Get table object properties based on its position in the items[] array. 

#### Syntax
```js
tablesCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index or position in the items[]. Zero indexed.

#### Returns

[table](table.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItemAt(0);
ctx.executeAsync().then(function () {
		Console.log(table.name);
});
```
[Back](#methods)
