# Binding
The Binding object is a member of the Bindings collection. The Bindings collection contains all the Binding objects in a workbook.

## [Properties](#get-binding)

| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|`id`   | String | String |The user-visible name of the binding. Get only.|Binding.Name    |       
|`type`| String |Returns the type of the binding. Can be `Table`,`Range` or `Text`. Get only. |Binding.Type|


## Relationships
None.    

## Methods

The Binding has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[getRange()][getrange-link]| [Range](range.md) object |Returns the Range of the Binding| |
|[getTable()][gettable-link]| [Table](table.md) object |Returns the Table of the Binding| |  
|[getText()][gettext-link]| String |Returns the text of the Binding| |  

## API Specification 
### getRange()

Get a Range object that represents a single cell or a range of cells. 

#### Syntax

```js
bindingObject.getRange();
```
#### Parameters
None.

#### Returns

[Range](range.md) object.

#### Examples

Below example uses binding to get the range.

```js
var ctx = new Excel.ExcelClientContext();
var binding = ctx.workbook.bindings.getItemAt(0);
var range = binding.getRange();
ctx.load(range);
ctx.executeAsync().then(function() {
	Console.log(range.cellCount);
});
```

[Back](#methods)

### getTable()

Get the table of the binding. 

#### Syntax
```js
bindingObject.getTable();
```
#### Parameters

None

#### Returns

[Table](table.md) object.

#### Examples

```js
var ctx = new Excel.ExcelClientContext();

var binding = ctx.workbook.bindings.getItemAt(0);
var table = binding.getTable();
ctx.load(table);
ctx.executeAsync().then(function () {
		Console.log(table.name);
});
```
[Back](#methods)

### getText()

Get the text of a binding. 

#### Syntax

```js
bindingObject.getText();
```
#### Parameters
None.

#### Returns
String.

#### Examples

Below example uses binding to get the range.

```js
var ctx = new Excel.ExcelClientContext();
var binding = ctx.workbook.bindings.getItemAt(0);
var text = binding.getText();
ctx.load(text);
ctx.executeAsync().then(function() {
	Console.log(text);
});
```

[Back](#methods)

### Get Binding

Get a Binding object properties. 

#### Syntax

```js
bindingObject.type;
```
#### Properties
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|`id`   | String | String |The user-visible name of the binding|Binding.Name    |       
|`type`| String |Returns the type of the binding. |Binding.Type|


#### Returns

[Binding](binding.md) object.

#### Examples

Below example to get binding properties.

```js
var ctx = new Excel.ExcelClientContext();
var binding = ctx.workbook.bindings.getItemAt(0);
ctx.load(binding);
ctx.executeAsync().then(function() {
	Console.log(binding.type);
});
```

[Back](#properties)




[getrange-link]: #getrange
[gettable-link]: #gettable
[gettext-link]: #gettext
