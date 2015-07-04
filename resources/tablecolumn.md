# Table Column
Represents a Column in a table. The TableColumn object is a member of the TableColumns collection.

## [Properties](#get-table-column)

|Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `id`     |  Number | Returns the unique key that identifies the column within the table.    |        |
|`index`          |  Number | Returns the index number of the column within the columns collection of the table. Zero-indexed.| ListColumn.Index|
|`name`           | String | Returns the name of the table column.| ListColumn.Name|
|`values`         | Array (Primitive)  | Returns unformatted values of the table Column. |Collection of ListColumn.Range.Value2|


## Relationships
None

## Methods

The TableColumn has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[delete()](#delete)| void     |Deletes the column from the table.  ||
|[getDataBodyRange()](#getdatabodyrange)| [Range](range.md) object     | Gets the range object associated with the data body of the column.||
|[getHeaderRowRange()](#getheaderrowrange)| [Range](range.md) object     | Gets the range object associated with the header row of the column.||
|[getRange()](#getrange)| [Range](range.md) object     | Gets the range object associated with the entire column.||
|[getTotalRowRange()](#gettotalrowrange)| [Range](range.md) object     | Gets the Range object associated with the totals row of the column.||

## API Specification 

### delete() 

Deletes the column from the table. 

#### Syntax

```js
tableColumnObject.delete();
```

#### Parameters 
None

#### Returns
Nothing

#### Example 

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
column.delete();
ctx.executeAsync().then();
```
[Back](#methods)

### getDataBodyRange() 
Get Range object associated with the Column's data body.

```js
tableColumnObject.getDataBodyRange();
```
#### Parameters

None

#### Returns

[Range](range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var dataBodyRange = column.getDataBodyRange();
ctx.load(dataBodyRange);
ctx.executeAsync().then(function () {
	Console.log(dataBodyRange.address);
});
```
[Back](#methods)

### getHeaderRowRange()

Get Range object associated with the Column's header.

#### Syntax

```js
tableColumnObject.getHeaderRowRange();
```

#### Parameters

None

#### Returns

[Range](range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var headerRowRange = columns.getHeaderRowRange();
ctx.load(headerRowRange);
ctx.executeAsync().then(function () {
	Console.log(headerRowRange.address);
});
```
[Back](#methods)

### getRange() 
Get Range object associated with the Column.

```js
tableColumnObject.getRange();
```
#### Parameters

None

#### Returns

[Range](range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var columnRange = columns.getRange();
ctx.load(range);
ctx.executeAsync().then(function () {
	Console.log(range.columnRange);
});
```
[Back](#methods)

### getTotalRowRange() 

Get Range object associated with the Column's total.

#### Syntax 

```js
tableColumnObject.getTotalRowRange();
```

#### Parameters

None

#### Returns

[Range](range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var columns = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(0);
var totalRowRange = columns.getTotalRowRange();
ctx.load(totalRowRange);
ctx.executeAsync().then(function () {
	Console.log(totalRowRange.address);
});
```

[Back](#methods)

### Get Table Column 

Get Table Column's data and properties.  

#### Syntax
```js
tableColumnsCollection.getItem(param);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `param`| String or Number | Column index (Zero-indexed) or column name of the column that you wish to get. 

#### Returns

[Table Column](tableColumn.md) object.


#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItem(0);
ctx.load(column);
ctx.executeAsync().then(function () {
	Console.log(column.index);
});
```
[Back](#properties)

### Update Table Column 

Update values of table column.

#### Syntax
```js
tableColumnObject.values = new-values
```
Where, new-values is a 2-dimensional array values of the table column. 

#### Example

```js
var ctx = new Excel.ExcelClientContext();
var tables = ctx.workbook.tables;
var newValues = [["New"], ["Values"], ["For"], ["New"], ["Column"]];
var column = ctx.workbook.tables.getItem(tableName).tableColumns.getItemAt(2);
column.values = newValues;
ctx.load(column);
ctx.executeAsync().then(function () {
	Console.log(column.values);
});
```
[Back](#properties)