# Table Collection

Represents a collection of all the tables that are part of the workbook. 

## [Properties](#get-table-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Returns the number of tables in the workbook.|ListObjects.count|
|`items`| [Table](table.md) Array | Returns a collection of all the table objects that are part of the workbook.|[ListObjects.item] |

## Relationships

None

## Methods

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(address: string, hasHeaders: bool)](#addaddress-string-hasheaders-bool)| [Table](table.md) object | Create a new table. The range source address determines the worksheet under which the table will be added. ||
|[getItem(id: object)](#getitemname-object)| [Table](table.md) object      |Gets a table by name or id.||
|[getItemAt(index: number)](#getitematindex-number)| [Table](table.md) object     |Gets a table based on its position in the collection.||

## API Specification 

### add(address: string, hasHeaders: bool)

Create a New Table object. The range source address determines the worksheet under which the table will be added. 

#### Syntax
```js
tableCollection.add(rangeAddress, hasHeader);
```
#### Parameters 

Parameter       | Type   | Description
--------------- | ------ | ------------
|`address`| String | Required. Address or name of the Range object representing the data source.
|`hasHeaders` | Boolean | Optional. Boolean value that indicates whether the data being imported has column labels. If the Source does not contain headers (i.e,. when this property set to `false`), Excel will automatically generate headers. If this property value is not set, Excel will determine the header row on its own.

#### Returns
[Table](table.md) object.

#### Example
```js
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.add('Sheet1!A1:E7', true);
ctx.load(table);
ctx.executeAsync().then(function () {
	Console.log(table.name);
});

```


[Back](#methods)

### getItem(id: object)

Gets a table object based on name or id.

#### Syntax
```js
tableCollection.getItem(id);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `id`| String | Required. Name or id of the table to be retrieved.

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

Gets a table object based on its position in the collection. 

#### Syntax
```js
tableCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index value of the object to be retrieved. Zero-indexed.

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

### Get table Collection

Get properties of the table collection. 

#### Syntax
```js
tableCollection.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tables.count|
|`items`| object[] | A collection of all the table objects that are part of the workbook|[tables.item] |


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