# RequestContext

The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, request context is required to get access to Excel and related objects such as worksheets, tables, etc. from the add-in. 

## Properties
None

## Methods

| Method         | Return Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|[load(param: object, expand: object)](loadparam-object)  |void     |Fills the Excel proxy object created in JavaScript layer with property and object values specified in the parameter.||
|[executeAsync()](load)  |Promise Object |Submits the request queue to Excel and returns a promise object, which can be used for chaining further actions.||

## API Specification

### load() 
Fills the Excel proxy object created in JavaScript layer with property and object values specified in the parameter

#### Syntax
```js
requestContextObject.load(<parameter>);
object.load(<parameter>);
```

#### Parameters
| Parameter       | Type    |Description|
|:----------------|:--------|:----------|
|select|Object|Optional. Specify the list of properties and relationships that needs to be loaded by Excel upon executeAsync() call. Also accepts an array containing the property/relationship names. By default all scalar/complex type properties of the object are loaded.|
|expand|Object|Optional. Specify the list of relationships that needs to be loaded by Excel upon executeAsync() call. Also accepts an array containing the relationship names.|

#### Returns
void

##### Examples

The following example shows how to read how to copy the values from Range A1:A2 to B1:B2.

```js
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
range.load ("address, values"); 
//or ctx.load(range, “address, values”);

ctx.executeAsync()
	.then(function () {
	var myvalues=range.values;
	ctx.workbook.worksheets. getActiveWorksheet().getRange("B1:B2").values= myvalues;
	ctx.executeAsync()
  		.then(function () {
			console.log(rang.address);
			console.log(rang.values);
		})
		.catch(function(error) {
			console. error(JSON.stringify(error));
		})
});
```
##### Example
The following example loads the name property of worksheet and names of tables that are part of the worksheet and their associated column names. It prints the worksheet name, table name and column names after executeAsync all. 

```js

var ctx = new Excel.RequestContext();
	var worksheets = ctx.workbook.worksheets;
	worksheets.load("name, items, tables\name, tables\column\name");
	ctx.executeAsync()
		.then(function () {
			for (var i = 0; i < worksheets.items.length; i++) {
				for (var j = 0; j < worksheets.items[i].tables.length ; j++) {
					for (var k = 0; k < worksheets.items[i].tables.items[j].columns.count; k++) {
						console.log(worksheets.items[i].name + worksheets.items[i].tables.items[j].name + worksheets.items[i].talbes.items[j].columns.items[k].name);
					}
				}
			}
		})
		.then(function () {
			console.log("Done");
		})
		.catch(function (error) {
		console.error(JSON.stringify(error));
		});
```

Following example uses expand to load the format relationship of the range.

```js
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
range.load(expand: "format") 
ctx.executeAsync()
	.then(function () {
			console.log(range.format.wrapText);
	.catch(function(error) {
			console. error(JSON.stringify(error));
		})
});

```

### executeAsync() 

#### Syntax
```js
requestContextObject.executeAsync();
```

#### Parameters
None

#### Returns
Promise object.

##### Examples


```js
	var ctx = new Excel.RequestContext();
	var sheet = ctx.workbook.worksheets.add();

	ctx.executeAsync()
		.then(function () {   			
			console.log("Done");
		 })
		.catch(function(error) {
			console. error(JSON.stringify(error));
		});
```
