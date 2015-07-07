# TableRowCollection

Represents a collection of all the rows that are part of the table.

## [Properties](#getter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|count|int|Returns the number of rows in the table. Read-only.||
|items|[TableRowCollection](tablerowcollection.md)|A collection of tableRow objects. Read-only.||

## Relationships
None


## Methods

| Method           | Return Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|[add(index: number, values: object[][])](#addindex-number-values-object)|[TableRow](tablerow.md)|Adds a new row to the table.||
|[getItemAt(index: number)](#getitematindex-number)|[TableRow](tablerow.md)|Gets a row based on its position in the collection.||

## API Specification

### add(index: number, values: object[][])
Adds a new row to the table.

#### Syntax
```js
tableRowCollectionObject.add(index, values);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|Optional. Specifies the relative position of the new row. If null, the addition happens at the end. Any rows below the inserted row are shifted downwards. Zero-indexed.|
|values|object[][]|Optional. A 2-dimensional array of unformatted values of the table row.|

#### Returns
[TableRow](tablerow.md)

#### Examples

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
Gets a row based on its position in the collection.

#### Syntax
```js
tableRowCollectionObject.getItemAt(index);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|index|number|Index value of the object to be retrieved. Zero-indexed.|

#### Returns
[TableRow](tablerow.md)

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

#### Getter Examples

```js
var ctx = new Excel.ExcelClientContext();
var tablerows = ctx.workbook.tables.getItem('Table1').rows;
ctx.load(tablerows);
ctx.executeAsync().then(function () {
	Console.log("tablerows Count: " + tablerows.count);
	for (var i = 0; i < tablerows.items.length; i++)
	{
		Console.log(tablerows.items[i].index);
	}
});
```
[Back](#properties)
