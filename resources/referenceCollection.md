# Reference Collection
Reference collection allows add-ins add and remove temporary references on range.

## Properties
None.

## Relationships

None

## Methods

The Binding collection has the following methods defined:

| Method     | Return Type    |Description|Notes  |
|:-----------------|:--------|:----------|:------|
|[add(rangeObject: Range)](#addrangeobject-range)| Null             |Creates a new reference on a range.  ||
|[remove(rangeObject: Range)](#removerangeobject-range)| Null             |Remove a reference on the range.  ||


## API Specification 

### add(rangeObject: range)

Add a new binding to the workbook. The binding will be added at the end of existing bindings.

#### Syntax
```js
referenceCollection.add(rangeObject);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| The Range Object which needs to be added to the reference collection.

#### Returns
Null

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.references.add(range);
range.insert("Down");

ctx.load(range);
ctx.executeAsync().then(function () {
	Console.log(range.address); // Address should be updated to A3:B4
});
```
[Back](#methods)

### remove(rangeObject: range)

Add a new binding to the workbook. The binding will be added at the end of existing bindings.

#### Syntax
```js
referenceCollection.remove(rangeObject);
```

#### Parameters

Parameter       | Type   | Description
--------------- | ------ | ------------
`rangeObject`  | [Range](range.md)| The Range Object which needs to be removed to the reference collection.

#### Returns
Null

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.ExcelClientContext();
var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
ctx.references.add(range);
range.insert("Down");

ctx.load(range);
ctx.executeAsync().then(function () {
	Console.log(range.address); // Address should be updated to A3:B4

	ctx.references.remove(range);
	range.insert("Down");
	ctx.executeAsync().then(function () {
		Console.log(range.address); // Address will remain A3:B4 though the underlying range shifted down after another range was inserted.
	});
});
```
[Back](#methods)