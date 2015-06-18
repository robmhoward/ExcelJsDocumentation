# NamedItem Collection
A collection of all the nameditem objects that are part of the workbook. 

## [Properties](#get-nameditem-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|nameditems.count|
|`items`| [Named Item](nameditem.md) Array | A collection of all the nameditem objects that are part of the workbook|[nameditems.item] |

## Relationships

None

## Methods

The nameditem collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[getItem(name: string)](#getitemname-string)| [Named Item](nameditem.md) object      |Gets a nameditem object using its name||
|[getItemAt(index: number)](#getitematindex-number)| [Named Item](nameditem.md) object     |Gets a nameditem based on its position in the collection.||


## API Specification 

### Get Nameditem Collection

Get properties of the nameditem collection. 

#### Syntax
```js
nameditemCollection.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|nameditems.count|
|`items`| [Named Item](nameditem.md) Array | A collection of all the nameditem objects that are part of the workbook|[nameditems.item] |


#### Returns

[nameditem](nameditem.md) collection. 

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var nameditems = ctx.workbook.names;
ctx.load(nameditems);
ctx.executeAsync().then(function () {
	for (var i = 0; i < nameditems.items.length; i++)
	{
		Console.log(nameditems.items[i].name);
		Console.log(nameditems.items[i].index);
	}
});
```

##### Getting the number of nameditems

```js
var ctx = new Excel.ExcelClientContext();
var nameditems = ctx.workbook.names;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("nameditems: Count= " + nameditems.count);
});

```
[Back](#properties)

### getItem(name: string)

Gets a nameditem object based on its name.

#### Syntax
```js
nameditemCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. nameditem name. 

#### Returns

[nameditem](nameditem.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var nameditem = ctx.workbook.names.getItem(wSheetName);
ctx.executeAsync().then(function () {
		Console.log(nameditem.type);
});
```
[Back](#methods)


### getItemAt(index: number)

Gets nameditem object based on its position in the collection. 

#### Syntax
```js
nameditemCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index value of the object to be retrieved. Zero-indexed.

#### Returns

[nameditem](nameditem.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var nameditem = ctx.workbook.names.getItemAt(0);
ctx.executeAsync().then(function () {
		Console.log(nameditem.name);
});
```
[Back](#methods)
