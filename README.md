# Excel JavaScript APIs

## Objects 

* [Workbook](resources/workbook.md): Workbook is the top level object which contains related workbook objects such as worksheets, tables, ranges, etc. It can be used to list related references. 
* [Worksheet](resources/worksheet.md): The Worksheet object is a member of the Worksheets collection. The Worksheets collection contains all the Worksheet objects in a workbook.
	* [Worksheet Collection](resources/worksheetcollection.md): A collection of all the Workbook objects that are part of the workbook. 
* [Range](resources/range.md): Range represents a cell, a row, a column, a selection of cells containing one or more contiguous blocks of cells.  
* [Table](resources/table.md): Represents collection of organized cells designed to make management of the data easy. 
	* [Table Collection](resources/tablecollection.md): A collection of Tables in a workbook or worksheet. 
	* [TableColumn Collection](resources/tablecolumncollection.md): A collection of all the columns in a Table. 
	* [TableRow Collection](resources/tablerowcollection.md): A collection of all the rows in a Table. 
* [Chart](resources/chart.md): Represents a chart object in a workbook, which is a visual representation of underlying data.   
	* [Chart Collection](resources/chartcollection.md): A collection of charts in a workbook or a worksheet.	
* [NamedItem](resources/nameditem.md): Represents a defined name for a range of cells or a value. Names can be primitive named objects (as seen in the type below), range object, etc.
	* [NamedItem Collection](resources/nameditemcollection.md): a collection of named items of a workbook.
* [Binding](resources/binding.md): An abstract class that represents a binding to a section of the workbook.
	* [Binding Collection](resources/bindingcollection.md):A collection of all the Binding objects that are part of the workbook. 
* [Reference Collection](resources/referencecollection.md): Reference collection allows add-ins to add and remove temporary references on range.
* [Request Context](resources/requestcontext.md): The RequestContext object facilitates requests to the Excel application.

Also read the following programming notes: 

* [Programming Notes](#programming-notes): Provide important programming details related to Excel APIs.
* [Error Messages](#error-messages): Provide important programming details related to Excel APIs.

[top](#excel-javascript-apis)

## Programming Notes

Following sections provide important programming details related to Excel APIs.

* [The Basics](#the-basics)
* [Properties and Relations Selection](#properties-and-relations-selection)
* [Document Binding](#null-input)
* [Reference Binding](#null-input)
* [Null-Input](#null-input)
* [Null-Response](#null-response)
* [Blank Input and Output](#blank-input-and-output)
* [Unbounded-Range](#unbounded-range)
* [Large-Range](#large-range)
* [Single Input Copy](#single-input-copy)
* [Throttling](#throttling)

[top](#excel-javascript-apis)

### The Basics

This section introduces three key concepts to help get started with the Excel API. Namely, RequestContext, executeAsync and load statements.  

#### RequestContext
The RequestContext object facilitates requests to the Excel application. Since the Office add-in and the Excel application run in two different processes, request context is required to get access to Excel and related objects such as worksheets, tables, etc. from the add-in. 

A request context object is created as shown bellow: 

```js
var ctx = new Excel.RequestContext();
```
#### executeAsync()

The Excel JavaScript objects created in the add-ins are local proxy objects. Any method invocation or setting of properties queues commands in JavaScript, but does not submit them until executeAsync() is called. executeAsync submits the request queue to Excel and  returns a promise object, which can be used for chaining further actions. 

##### Example

The following example shows how to write values from an array to a range. First RequestContext() is created to get access to the workbook. Then a worksheet is added. Range A1:B2 on the sheet is retrieved afterwards. Finally we assign the values stored in the array to this range. All these commands are queued and will run when ctx.executeAsync() is called.  executeAsync() returns a promise which can be used to chain it with other operations.

```js
	var ctx = new Excel.RequestContext();
	var sheet = ctx.workbook.worksheets.add();
	var values = [
				 ["Type", "Estimate"],
				 ["Transportation", 1670]
				 ];
	var range = sheet.getRange("A1:B2");
	range.values = values;
	//statements queued above will not be executed until the executeAsync() is called. 
	ctx.executeAsync()
		.then(function () {   			
			console.log("Done");
		 })
		.catch(function(error) {
			console. error(JSON.stringify(error));
		});
```

#### load()
Load method is used to fill in the Excel proxy objects created in the add-in JavaScript layer. When trying to retrieve an object, say, a worksheet, a local proxy object is created first in the JavaScript layer. Such an object can be used to queue up setting of its properties and invoking methods. However, for reading object properties or relations, the load() method and executeAsync() needs to be invoked first. Load method takes in the parameters and relations that needs to be loaded when the executeAsync is called. 

##### Syntax

```js
object.load(param);
//or
object.load({loadOption});
```
Where, 

* param is the list of parameter/relationship names to be loaded specified as comma delimited strings or array of names. See .load() methods under each object for details.
* loadOption specifies selection, expansion, top/skip options. See loadOption object for details.

##### Example
The following example shows how to read how to copy the values from Range A1:A2 to B1:B2 by using load() method on the range object.

```js
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
range.load ("address, values, range/format"); 

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

The following example shows how to read how to copy the values from Range A1:A2 to B1:B2 by using load() method on the context object.

```js
var ctx = new Excel.RequestContext();
var range = ctx.workbook.worksheets.getActiveWorksheet().getRange("A1:A2");
range.load({select": "address, values", "expand" : "range/format"});

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

#### Summary
1.	Getting a RequestContext is the first step to interact with Excel.
2.	All JavaScript objects are local proxy objects.  Any method invocation or setting of properties queues commands in JavaScript, but does not submit them until executeAsync() is called. 
3.	Load is a special type of command for retrieval of properties. Properties can only be accesed after invoking executeAsync(). 
4.	For performance reasons, avoid loading objects without specifying individual properties that will be used.

[top](#programming-notes)

### Properties and Relations Selection 

* By default load() selects all scalar/complex properties of the object which is being loaded. The relations are not loaded by default.  Exceptions:  any binary, XML, etc properties are not returned. 
* The select option specifies a subset of properties and/or relations to include in the response.
* Default Select behavior: 
	*	Does not select any property
	*	Need to specify every property that needs to be returned
	*	Relations/Navigation properties are also allowed to be included in the list. Use expand syntax to 
* The properties to be selected are provided during the load statement.
* Select will essentially get the users into optimized mode of handpicking what they want. 
* Property names are listed as a parameter to the select property. Support two kinds of inputs
	* Property names are separated by comma. 
	* Provide an array of property name strings

```js	
context.load (<var1>,<relation1/var2>);
object.load  (<var1>,<relation1/var2>);

// Pass the parameter as an array.
context.load (["var1", "relation1/var2"]);
object.load (["var1", "relation1/var2"]);
```

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.ExcelClientContext();
var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

//load statement below loads the address, values, numberFormat properties of the Range and then expands on the format, format/background, entireRow relations
 
myRange.load (["address", "values", "numberFormat", "format", "format/background", "entireRow"]);

ctx.executeAsync().then(function () {
		console.log (myRange.address); //ok
		console.log (myRange.cellCount); //not-ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.background.color); //ok
		console.log (myRange.format.font.color); //not-ok
		console.log (myRange.entireRow.address); //ok
		console.log (myRange.entireColumn.address); //not-ok
// . . . 

//load statement below loads all the properties of the Range and then expands on the format, format/background, entireRow relations.  
 
myRange.load(["address", "format", "format/background", "entireRow" ]);

ctx.executeAsync().then(function () {
		console.log (myRange.address); //ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.background.color); //ok
		console.log (myRange.format.font.color); //not-ok
		console.log (myRange.entireRow.address); //ok
		console.log (myRange.entireColumn.address); //not-ok
```

[Back](#programming-notes)

### Null-Input

#### null input in 2-D Array

**`null` input inside 2 dimensional array (for values, number-format, formula) is ignored** in the update API. No update will take place to the intended target when `null` input is sent in values or number-format or formula grid of values.

Example: In order to only update specific parts of the Range such as some cell's Number Format and retain the existing Number Format on other parts of the Range, set desired Number Format where needed and send `null` for the other cells. 

In below set request, only some parts of the Range Number Format is set while retaining the existing Number Format on the remainig part (by passing nulls).

```js
  range.values = [["Eurasia", "29.96", "0.25", "15-Feb" ]];
  range.numberFormat = [[null, null, null, "m/d/yyyy;@"]];
```
#### null input for a property

**`null` is not a valid single input for the entire property.** e.g., following is not valid as the entire values cannot be set to null or ignored. 

```
 range.values= null;

```

Following is not valid either as null is not a valid color value. 
```
 range.format.background.color =  null;
```
[Back](#programming-notes)
### Null-Response

Representation of formatting properties that consists of non-uniform values would result in `null` value to be returned in the response. 

Example: A Range can consist of one of more cells. In cases where the individual cells contained in the Range specified doesn't have uniform formatting values, the range level representation will be undefined. 

```
  "size" : null,
  "color" : null,
```

### Blank Input and Output

Blank values in update requests are treated as instruction to clear or reset the respective property. Blank value is represented by two double-quotes with no space in between. `""`

Example: 
* For `values`, the range value is cleared out. This is same as clearing the contents in the application.
* For `numberFormat`, the number format is set to `General`.
* For `formula` and `formulaLocale`, the formula values are clearned out. 

For read operations, expect to receive blank values if the contents of the cells are blanks. If the cell contains no data or value, then the API returns a blank value. Blank value is represented by two double-quotes with no space in between. `""`.

```
  range.values = [["", "some", "data", "in", "other", "cells", ""]];
```

```
  range.formula = [["", "", "=Rand()"]];
```
[Back](#programming-notes)
### Unbounded-Range

#### Read

Unbounded range address contains only column or row identifiers and unspecified row identifier or column identifiers (respectively), such as:

* `C:C`, `A:F`, `A:XFD` (contains unspecified rows)
* `2:2`, `1:4`, `1:1048546` (contains unspecified columns)

When the API makes a request to retrieve an unbounded Range (e.g., `getRange('C:C')`, the response returned contains `null` for cell level properties such as `values`, `text`, `numberFormat`, `formula`, etc.. Other Range properties such as `address`, `cellCount`, etc. will reflect the unbounded range.

#### Write

Setting cell level properties (such as values, numberFormat, etc.) on unbounded Range is **not allowed** as the input request might be too large to handle. 

Example: following is not a valid update request as the requested range is unbounded one. 
```js
var sheetName = 'Sheet1';
var rangeAddress = 'A:B';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
range.values = 'Due Date';
range.load(text);
ctx.executeAsync().then(function() {
	Console.log(range.text);
});

```

When such a Range is update operation is attempted, the API returns the an error.

[Back](#programming-notes)
### Large-Range

Large Range implies a Range whose size is too large for a single API call. Many factors such as number of cells or values or number-formats, or formulas, etc. contained in the range can make the response large enough to be unsuitable for API interaction. 

The API makes best attempt to return or write-to the requested data. However, due to the large size involved, API might result in an error condition due to large resource utilization. 

In order to avoid such condition, it is recommended to read or write large Range in multiple smaller range sizes.

[Back](#programming-notes)
### Single Input Copy

To support updating a range with same values or number-format or applying same formula across a range, the following convention is used in the set API. In Excel, this behavior is similar to inputting values or formulas to a range in the CTRL+Enter mode. 

API will look for *single cell value* and and if the target range dimension doesn't match the input range dimension it will apply the update to the entire range in the CTRL+Enter model with the value or formula provided in the request.

#### Examples

Following request updates selected range with the a text of "Due Date". Note that Range has 20 cells whereas the provided input only has 1 cell value.

```js
var sheetName = 'Sheet1';
var rangeAddress = 'A1:A20';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
range.values = 'Due Date';
range.load(text);
ctx.executeAsync().then(function() {
	Console.log(range.text);
});

```

Following request updates selected range with date of 3/11/2015".  

```js
var sheetName = 'Sheet1';
var rangeAddress = 'A1:A20';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
range.numberFormat = 'm/d/yyyy';
range.values = '3/11/2015';
range.load(text);
ctx.executeAsync().then(function() {
	Console.log(range.text);
});

```
Following request updates selected range with a formula of that will be applied across in the CTRL+Enter mode.  

```js
var sheetName = 'Sheet1';
var rangeAddress = 'A1:A20';
var ctx = new Excel.ExcelClientContext();
var worksheet = ctx.workbook.worksheets.getItem(sheetName);
var range = worksheet.getRange(rangeAddress);
range.formula = '=DAYS(B15,42060)';
range.load(text);
ctx.executeAsync().then(function() {
	Console.log(range.text);
});
```
[Back](#programming-notes)
### Throttling 

Excel Service uses throttling to maintain optimal performance and reliability of the service. Throttling limits the number of user actions or concurrent calls (by script or code) to prevent overuse of resources.

Though this is less common, certain pattern of API usage such as high frequency requests or high volume requests that increases CPU or memory utilization of the servers beyond limit would likely get you throttled.

When a user exceeds usage limits, Excel service throttles any further requests from that user account for a short period. All user actions are throttled while the throttle is in effect.

API requests while the throttle is in effect will result in below error condition:

```js
ctx.executeAsync().then(
function () {
	Console.log("...");
	},
	function (error) {
	   some.log("ErrorCode =" + error.code); //"ActivityLimitReached"
	   some.log("ErrorMessage =" + error.message); //"Activity limit has been reached."
	});
```
[Back](#programming-notes)

## Error Messages

Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur. 

|error.code | error.message |
|----------:|--------------:|
|InvalidArgument |The argument is invalid or missing or has an incorrect format.|
|InvalidRequest  |Cannot process the request.|
|InvalidReference|This reference is not valid for the current operation.|
|InvalidBinding  |This object binding is no longer valid due to previous updates.|
|InvalidSelection|The current selection is invalid for this operation.|
|Unauthenticated |Required authentication information is either missing or invalid.|
|AccessDenied	|You cannot perform the requested operation.|
|ItemNotFound	|The requested resource doesn't exist.|
|ActivityLimitReached|Activity limit has been reached.|
|GeneralException|There was an internal error while processing the request.|
|NotImplemented  |The requested feature isn't implemented.|
|ServiceNotAvailable|The service is unavailable.|
|Conflict	|Request could not be processed because of conflict.|
|ItemAlreadyExists|The resource being created already exists.|
|UnsupportedOperation|The operation being attempted is not supported.|
|RequestAborted|The request was aborted during run time.|
|ApiNotAvailable|The requested API is not available.|
|InsertDeleteConflict|The insert or delete operation attempted resulted in conflict.|
|InvalidOperation|The operation attempted is invalid on the object.|

#### Examples

```js
ctx.executeAsync().then(
function () {
	Console.log("...");
	},
	function (error) {
	   some.log("ErrorCode =" + error.code); //"InvalidArgument"
	   some.log("ErrorMessage =" + error.message); //"The argument is invalid or missing or has an incorrect format."
	});
```

[top](#excel-javascript-apis)

