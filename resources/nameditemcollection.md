# NamedItemCollection

A collection of all the nameditem objects that are part of the workbook.

## [Properties](#getter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|items|[NamedItemCollection](nameditemcollection.md)|A collection of namedItem objects. Read-only.||

## Relationships
None


## Methods

| Method           | Return Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|[getItem(name: string)](#getitemname-string)|[NamedItem](nameditem.md)|Gets a nameditem object using its name||

## API Specification

### getItem(name: string)
Gets a nameditem object using its name

#### Syntax
```js
namedItemCollectionObject.getItem(name);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|name|string|nameditem name.|

#### Returns
[NamedItem](nameditem.md)

#### Examples

```js
var ctx = new Excel.RequestContext();
var nameditem = ctx.workbook.names.getItem(wSheetName);
ctx.executeAsync().then(function () {
		Console.log(nameditem.type);
});
```

[Back](#methods)

#### Getter Examples

```js
var ctx = new Excel.RequestContext();
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

Get the number of nameditems.

```js
var ctx = new Excel.RequestContext();
var nameditems = ctx.workbook.names;
ctx.load(tables);
ctx.executeAsync().then(function () {
	Console.log("nameditems: Count= " + nameditems.count);
});

```


[Back](#properties)
