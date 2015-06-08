# Binding Collection
A collection of all the [Binding](binding.md) objects that are part of the workbook. 

## [Properties](#get-binding-collection)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|Bindings.count|
|`items`| [Binding](binding.md) Array | A collection of all the Binding objects that are part of the workbook|[Bindings.item] |

## Relationships

None

## Methods

The Binding collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(id: string)](#addid-string)| [Binding](binding.md) Object              |Creates a new binding.  ||
|[getItem(param: string)](#getitemparam-string)| [Binding](binding.md) Object      |Retrieve a binding  object using its id||
|[getItemAt(index: number)](#getitematindex-number)| [Binding](binding.md) Object     |Retrieve a binding based on its position in the items[] array.||


## API Specification 

### add(id: string)

Add a new binding object. 

#### Syntax
```js
bindingCollection.add(id);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`id`  | String| String value representing the id of the binding to be added. If not specified, Excel determines the id of the new binding being added. 

#### Returns
[Binding](binding.md) object.

#### Examples

```js
var bID = 'Sample Name';
var ctx = new Excel.ExcelClientContext();
var binding = ctx.workbook.bindings.add(bID);
ctx.load(binding);
ctx.executeAsync().then(function () {
	Console.log(binding.id);
});
```
[Back](#methods)


### getItemAt(index: number)

Get Binding object properties based on its position in the items[] array. 

#### Syntax
```js
bindingCollection.getItemAt(index);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `index`| Number | Required. Index or position in the items[]. Zero indexed.

#### Returns

[Binding](binding.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var lastPosition = ctx.workbook.bindings.count - 1;
var binding = ctx.workbook.bindings.getItemAt(lastPosition);
ctx.executeAsync().then(function () {
		Console.log(binding.id);
});
```
[Back](#methods)


### Get Binding Collection

Get properties of the binding collection. 

#### Syntax
```js
workbookObject..bindings.property;
```

#### Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`count`| Number   | Number of objects in the collection.|Bindings.count|
|`items`| [Binding](binding.md) Array | A collection of all the binding objects that are part of the workbook|[Bindings.item] |


#### Returns

[Binding](binding.md) collection. 

#### Examples

```js
var ctx = new Excel.ExcelClientContext();
var bindings = ctx.workbook.bindings;
ctx.load(bindings);
ctx.executeAsync().then(function () {
	for (var i = 0; i < bindings.items.length; i++)
	{
		Console.log(bindings.items[i].id);
		Console.log(bindings.items[i].index);
	}
});
```

##### Getting the number of bindings

```js
var ctx = new Excel.ExcelClientContext();
var bindings = ctx.workbook.bindings;
ctx.load(bindings);
ctx.executeAsync().then(function () {
	Console.log("Bindings: Count= " + bindings.count);
});

```
[Back](#properties)