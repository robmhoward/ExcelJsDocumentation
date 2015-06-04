# Table
Represents collection of organized cells designed to make management of the data easy.

## [Properties](#get-table)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `id`  |  Number | A unique key that identifies the Table object in a Workbook. Note: If the table gets deleted, Excel Workbook could re-use the id value for another Table.   |        |
| `name`       | String| String value that represents the name of the Table object   | ListObject.Name       |
| `showHeaders` | Boolean| Boolean to indicate whether the Header row is visible or not. This value can be set to show or remove the header row| ListObject.ShowHeaders|
| `showTotals` | Boolean| Boolean to indicate whether the Total row is visible or not. This value can be set to show or remove the total row| ListObject.ShowTotals|
| `tableStyle` | String | Constant that represents the Table style. Possible values include: `Light1` thru `Light21`, `Medium1` thru `Medium28`, `StyleDark1` thru `StyleDark11`|ListObject.TableStyle|

## Relationships
The Table has the following relationships defined:

| relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| tableColumns  | [TableColumn collection](tablecolumncollection.md)       |Collection of Table columns |ListObject.TableColumns  |          
| tableRows      | [TableRow collection](tablerowcollection.md)         |Collection of Table rows |ListObject.ListRows      |

## Methods

The Worksheet has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[deleteObject()](#delete)| void     |Deletes the worksheet ||
|[getHeaderRowRange()](#getheaderrowrange) | [Range](range.md) Object |Retrieve the Range object associated with Data Body of the Table.||
|[getDataBodyRange()](#getdatabodyrange) | [Range](range.md) Object |Retrieve the Header Row Range object associated with the Table  ||
|[getRange()](#getrange) | [Range](range.md) Object |Get Range object associated with the Table. ||
|[getTotalRowRange()](#gettotalrowrange) | [Range](range.md) Object |Get Totals Range object associated with the Table. ||

## API Specification 

### deleteObject()

Deletes Table and clears the cell data from the Table.

#### Syntax
```js
tableObject.delete();
```

#### Parameters 
None

#### Returns
Nothing

#### Example 

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
table.deleteObject();
ctx.executeAsync().then();
```
[Back](#methods)

### getHeaderRowRange()

Deletes Table and clears the cell data from the Table.

#### Syntax
```js
tableObject.delete();
```

#### Parameters 
None

#### Returns
Nothing

#### Example 

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
table.deleteObject();
ctx.executeAsync().then();
```
[Back](#methods)


### getDataBodyRange()

Get Data Body Range object associated with the Table.

#### Syntax
```js
tableObject.getDataBodyRange();
```

#### Parameters

None

#### Returns

[Range](range.md) object.


#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableDataRange = table.getDataBodyRange();
ctx.executeAsync().then(function () {
		Console.log(tableDataRange.address);
});
```
[Back](#methods)
### getHeaderRowRange()

Get Header Range object associated with the Table.

#### Syntax
```js
tableObject.getHeaderRowRange();
```

#### Parameters

None

#### Returns


[Range](range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableHeaderRange = table.getHeaderRowRange();
ctx.executeAsync().then(function () {
		Console.log(tableHeaderRange.address);
});
```
[Back](#methods)


### getRange()

Get Range object associated with the Table.

#### Syntax
```js
tableObject.getRange();
```

#### Parameters

None

#### Returns

[Range](range.md) object.


#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableRange = table.getRange();
ctx.executeAsync().then(function () {
		Console.log(tableRange.address);
});
```

[Back](#methods)
### getTotalRowRange()

Get Totals Range object associated with the Table.

#### Syntax
```js
tableObject.getTotalRowRange();
```

#### Parameters

None

#### Returns

[Range](range.md) object.

#### Examples

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
var tableTotalsRange = table.getTotalRowRange();
ctx.executeAsync().then(function () {
		Console.log(tableTotalsRange.address);
});
```
[Back](#methods)


### Get Table

Get Table object properties based on name. 

#### Syntax

```js
tableCollection.getItem(name);
```

#### Parameters

Parameter        | Type   | Description
---------------  | ------ | ------------
 `name`| String  | Required. Table name. 

#### Syntax
```js
tableCollection.getItemAt(index);
```

#### Parameters

Parameter        | Type   | Description
---------------  | ------ | ------------
 `index`| Number | Required. Table index. Zero indexed.

#### Returns

[Table](table.md) object. 

#### Examples

##### Getting a table by name

```js
var ctx = new Excel.ExcelClientContext();
var tableName = 'Table1';
var table = ctx.workbook.tables.getItem(tableName);
ctx.executeAsync().then(function () {
		Console.log(table.index);
});
```
##### Getting a table by index

```js
var ctx = new Excel.ExcelClientContext();
var index = 0;
var table = ctx.workbook.tables.getItemAt(0);
ctx.executeAsync().then(function () {
		Console.log(table.name);
});
```
[Back](#properties)

### Update Table

This API allows setting of Table properties such as name and show totals. In order to update the table content, use the update table row or column API.

#### Syntax
```js
tableObject.property = 'new-value';
```

#### Properties 

Following properties can be updated directly. 

|Property      | Type   | Description      |
|-------------- | ------ | -----------------|
| `name`        | String | String value that represents the name of the Table object   | 
| `showTotals`  | Boolean| Boolean to indicate whether the Total row is visible. This value can be set to show or remove the total row| 
| `tableStyle`  | String | Constant that represents the Table style. Possible values include: `TableStyleLight1` thru `TableStyleLight21`, `TableStyleMedium1` thru `TableStyleMedium28`, `TableStyleDark1` thru `TableStyleDark11`|

#### Example 

```js
var tableName = 'Table1';
var ctx = new Excel.ExcelClientContext();
var table = ctx.workbook.tables.getItem(tableName);
table.name = 'Table1-Renamed';
table.showTotals = false;
table.tableStyle = 'TableStyleMedium2';
ctx.load(table);
ctx.executeAsync().then(function () {
		Console.log(table.tableStyle);
});
```
[Back](#properties)

