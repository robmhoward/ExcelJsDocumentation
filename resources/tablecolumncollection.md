# tablecolumn Collection
A collection of all the tablecolumn objects that are part of the table. 

## [Properties](#get-tablecolumn-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tablecolumns.count|
|`items`| Object[] | A collection of all the tablecolumn objects that are part of the table|[tablecolumns.item] |

## Relationships

None

## Methods

The tablecolumn collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(name: string)](#addname-string)| [tablecolumn](tablecolumn.md)              |Creates a new tablecolumn.  ||
|[getItem(name: string)](#getitemname-string)| [tablecolumn](tablecolumn.md)      |Retrieve a tablecolumn object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [tablecolumn](tablecolumn.md)     |Retrieve a tablecolumn based on its position in the items[] array.||


## API Specification 

### Get tablecolumn Collection

Get properties of the tablecolumn collection. 

#### Syntax
```js
context.workbook.tablecolumns.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|tablecolumns.count|
|`items`| Object[] | A collection of all the tablecolumn objects that are part of the table|[tablecolumns.item] |


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
		Console.log(tablecolumns.items[i].index);
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

### add(name: string)

Add a new tablecolumn to the workbook. The tablecolumn will be added at the end of existing tablecolumns.

#### Syntax
```js
tablecolumnsCollection.add(name);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`name`  | String| Optional. String value representing the name of the sheet to be added. If not specified, Excel determines the name of the new tablecolumn being added. 

#### Returns
[tablecolumn](tablecolumn.md) object.

#### Examples

```js
var wSheetName = 'Sample Name';
var ctx = new Excel.ExcelClientContext();
var tablecolumn = ctx.workbook.tablecolumns.add(wSheetName);
ctx.load(tablecolumn);
ctx.executeAsync().then(function () {
	Console.log(tablecolumn.name);
});
```
[Back](#methods)

### getItem(name: string)

Get tablecolumn object properties based on name.

#### Syntax
```js
tablecolumnsCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. tablecolumn name. 

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
tablecolumnsCollection.getItemAt(index);
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
