# Named Item

Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range.
This object can be used to obtain Range object associated with names.

## [Properties](#get-named-item)

| Property         | Type    |Description|Notes  |
|:-----------------|:--------|:----------|:-----|
| `name`  | String|String value representing the name of the object.| Name.Name|
| `type` | String|Indicates what type of reference is associated with the name. Possible options are: `Range`, `String`, `Integer`, `Double`, `Boolean`. | Derived property |
| `value`| String |Represents the formula that the name is defined to refer to. e.g., `=Sheet14!$B$2:$H$12`, `=4.75`, etc. | Name.Value|
| `visibile` | Boolean |Boolean value that determines whether the object is visible. | Name.Visible |

## Relationships
None
     
## Methods

The Worksheet resource has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[getRange()](#getrange)| [Range](range.md) object |Returns the Range object that is associated with the name. `null` if the name is not of the type `Range`.| |

## API Specification 

### getRange()

Returns the Range object that is associated with the name. `null` if the name is not of the type `Range`. 

**Note: This API currently supports only the Workbook scoped items.**

#### Syntax
```js
namedItemObject.getRange(); 
```

#### Parameters
None

#### Returns

[Range](range.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var names = ctx.workbook.names;
var range = names.getItem('MyRange').getRange();
ctx.load(range);
ctx.executeAsync().then(function () {
		Console.log(range.address);
});
```
[Back](#methods)

### Get Named Item

Get a Named object. 

** Note: This API currently supports only the Workbook scoped items. **
#### Syntax
```js
namesCollection.getItem(name);
```

#### Parameters

Parameter       | Type  | Description
--------------- | ------ | ------------
 `name`| String | Required. Name of the item.

#### Returns

[Named-Item](nameditem.md) object.

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
var names = ctx.workbook.names;
var namedItem = names.getItem('MyRange');
ctx.load(namedItem);
ctx.executeAsync().then(function () {
		Console.log(namedItem.type);
});
```
[Back](#properties)
