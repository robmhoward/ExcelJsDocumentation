# LoadOption

An object that can be passed to the load method to specify the options such as select, expand parameters. 

## Properties
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|select|object|Provide comma delimited list or an array of parameter/relationship names to be loaded upon executeAsync call. e.g. "property1, relationship1", [ "property1", "relationship1"]. Optional.||
|expand|object|Provide comma delimited list or an array of relationship names to be loaded upon executeAsync call. e.g. "relationship1, relationship2", [ "relationship1", "relationship2"]. Optional.||
|top|int| Specify the number of items in the queried collection to be included in the result. Optional.||
|skip|int|Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the selection of result will start after skipping the specified number if items. Optional.||

### Example

The following example shows how to read and copy the values from Range A1:A2 to B1:B2.

```js
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
ctx.load(range, {"select": "address, values", "expand" : "range/format"});

ctx.executeAsync()
	.then(function () {
	var myvalues=range.values;
	ctx.workbook.worksheets. getActiveWorksheet().getRange("B1:B2").values= myvalues;
	ctx.executeAsync()
  		.then(function () {
			console.log(range.address);
			console.log(range.values);
			console.log(range.format.wrapText);
		})
		.catch(function(error) {
			console. error(JSON.stringify(error));
		})
});
```