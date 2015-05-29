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
* [Chart](#chart): Represents a chart object in a workbook, which is a visual representation of underlying data.   
* [Chart Collection](resources/chartcollection.md): A collection of charts in a workbook or a worksheet.    
* [Named-Item](resources/nameditem.md): Represents a defined name for a range of cells or a value. Names can be primitive named objects (as seen in the type below), range object, etc.
* [Names Collection](resources/nameditemcollection.md): a collection of named items of a workbook.

Also see: 

* [Error Messages](#error-messages): Provide important programming details related to Excel APIs.
* [Programming Notes](#programming-notes): Provide important programming details related to Excel APIs.

## Chart

[Chart](resources/chart.md) represents a chart object on a worksheet. 

Following are the methods supported for this resource:

| Task                               | Description                                | 
|:------------------------------------|:-------------------------------------------|
| [Add-Chart](#add-chart)     | Inserts a chart directly onto the grid.  |
| [Get-Chart](#get-chart)   | Gets a chart by name. |
| [Delete-Chart](#delete-chart)     | Deletes a chart directly on the grid.  |
| [Update-Chart](#update-chart)   | Update a chart including renaming, positioning and resizing. |
| [Set-Chart-SourceData](#set-chart-sourcedata)   | Sets the sourceData and seriesBy of a Chart.|
| [Format-Chart](#format-chart)   | Format a chart.|
| [Get-Chart-Title](#get-chart-title)   | Get the title of a chart. |
| [Set-Chart-Title](#set-chart-title)   | Set the title of a chart, including `text`, `position` and `overlay`. |
| [Delete-Chart-Title](#delete-chart-title)   | Delete the title from a chart. |
| [Format-Chart-Title](#format-chart-title)   | Format the Chart Title. |
| [Set-Chart-Legend](#set-chart-legend)   | Hide/Show Chart Legent and set position. |
| [Set-Chart-DataLabels](#set-chart-datalabels)   | Set display content and position of DataLabels. |
| [Set-Chart-Axis](#set-chart-axis)   | Set the `maximum`, `minimum`, `majorunit` and `minorunit` of an axis. |
| [Set-Chart-AxisTitle](#set-chart-axistitle)   | Change the Axis Title text and visibility. |
| [Add-Chart-Gridlines](#add-chart-gridlines)   | Show Gridlines on an Axis |
| [Format-Chart-Series](#format-chart-series)   | Change the Fill Color of a series |

### Add-Chart

Inserts a chart directly onto the grid.

#### Syntax

```js
chartsCollection.add(chartType, sourceData, seriesBy);
```

#### Parameters

| Parameter         | Value    |Description|
|:-----------------|:--------|:----------|
| `type` | String | A String value that represents the type of a chart.  |
| `sourceData`  | String | Sets an address or name of the Range object as the data source.|
| `seriesBy` | String | Sets the way columns or rows are used as data series on the chart. Can be `auto`, `Rows` or `Columns`.|


#### Returns

[Chart](resources/chart.md) object. 

#### Examples

##### Add a chart of `chartType` "ColumnClustered" on worksheet "Charts" with `sourceData` from Range "A1:B4" and `seriresBy` is set to be "auto".

```js
var sheetName = "Charts";
var sourceData = sheetName + "!" + "A1:B4";
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem(sheetName).charts.add("ColumnClustered", sourceData, "auto");
ctx.executeAsync().then(function () {
		logComment("New Chart Added");
});
```
[Back](#chart)

### Get-Chart

Gets a chart by name.

#### Syntax
```js
chartsCollection.getItem(name);	
```

#### Parameters
None.

#### Returns

[Chart](resources/chart.md) object. 

#### Examples

##### Get the Chart named "Chart1"
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

ctx.load(chart);
ctx.executeAsync().then(function () {
		logComment("Chart1 Loaded");
});
```

[Back](#chart)

### Delete-Chart

Deletes a chart directly on the grid.

#### Syntax

```js
chartObject.deleteObject();
```

#### Parameters
None.

#### Returns

Nothing.

#### Examples

##### Delete the Chart named "Chart1"

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.deleteObject();
ctx.executeAsync().then(function () {
		logComment"Chart Deleted");
});
```
[Back](#chart)

### Update-Chart

Update a chart including renaming, positioning and resizing.

#### Syntax

```js
chartObject.name="New Name";
chartObject.top = 100;
chartObject.left = 100;
chartObject.height = 200;
chartObject.weight = 200;
```

#### Properties
| Property         | Value    |Description|
|:-----------------|:--------|:----------|
| `name`  | String|A String value that represents the name of a Chart object                              |
| `height`|  Double |Returns or sets a Double value that represents the height, in points, of the object |
| `width` |  Double |Returns or sets a Double value that represents the width, in points, of the object. | 
| `top` |  Double |Returns or sets a Double value that represents the distance, in points, from the top edge of the object to the top of row 1 (on a worksheet) or the top of the chart area (on a chart).|
| `left` |  Double |Returns or sets a Double value that represents the distance, in points, from the left edge of the object to the left edge of column A (on a worksheet) or the left edge of the chart area (on a chart). | 

#### Returns

[Chart](resources/chart.md) object. 

#### Examples

##### Rename the chart to new name, resize the chart to 200 points in both height and weight. Move Chart1 to 100 points to the top and left. 
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");

chart.name="New Name";	
chart.top = 100;
chart.left = 100;
chart.height =200;
chart.width =200;
ctx.executeAsync().then(function () {
		logComment("Chart Updated");
});
```
[Back](#chart)


### Set-Chart-SourceData

Sets the sourceData and seriesBy of a Chart.

#### Syntax

```js
chartObject.setData(sourceData, seriesBy);
```

#### Parameters
| Parameter         | Value    |Description|
|:-----------------|:--------|:----------|
| `sourceData`  | String|  Sets an address or name of the Range object as the data source.|
| `seriesBy`  | String |  Sets the way columns or rows are used as data series on the chart. Can be one of the following `Rows`, `Columns` or `Auto`.|

#### Returns

[Chart](resources/chart.md) object. 

#### Examples

##### Set the `sourceData` to be "A1:B4" and `seriesBy` to be "Columns"
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	
var sourceData = "A1:B4";

chart.setData(sourceData, "Columns");
ctx.executeAsync().then();
```
[Back](#chart)


### Format-Chart

Format a chart.

#### Syntax

```js
chartObject.fillFormat.SetSolidColor(color);
```

#### Parameters
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|color| String | HTML color code representing the color of the interior/background. |

#### Returns
Nothing.

#### Examples

##### Set "Chart1" background to be red.

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.fillFormat.SetSolidColor("#FF0000");
ctx.executeAsync().then(function () {
		logComment("Chart Color Changed ");
});
```
[Back](#chart)


### Get-Chart-Title

Get the title of a chart.

#### Syntax

```js
chartObject.title.text;
```

#### Parameters
None. 

#### Returns
[ChartTitle](resources/chartTitle.md) object. 

#### Examples

##### Get the `text` of Chart Title from Chart1
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

var title = chart.title;
ctx.load(title);
ctx.executeAsync().then(function () {
		logComment(title.text);
});
```
[Back](#chart)


### Set-Chart-Title

Set the title of a chart, including `text`, `position` and `overlay`.

#### Syntax

```js
chartObject.title.text= "My Chart"; 
chartObject.title.overlay=true;
```

#### Properties
| Property         | Type    |Description| 
|:-----------------|:--------|:----------|
| `text` | String |A String value that represents the title text of a chart. When a title text is set, the display property will be automaticlly set to top and the chart title will be displayed on top of the chart without overlapping. |  
| `overlay` | Boolean |True if the title overlays the chart. | 

#### Returns
[ChartTitle](resources/chartTitle.md) object. 

#### Examples

##### Set the `text` of Chart Title to "My Chart" and Make it show on top of the chart without overlaying.
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.title.text= "My Chart"; 
chart.title.overlay=true;

ctx.executeAsync().then(function () {
		logComment("Char Title Changed");
});
```
[Back](#chart)

### Delete-Chart-Title

Delete the title from a chart.

#### Syntax

```js
chartObject.title.visible = false; 
```

#### Parameters
None. 

#### Returns
None.

#### Examples

##### Hide the title of Chart1
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.title.visible = false; 
ctx.executeAsync().then(function () {
		logComment("Title Hidden");
});
```
[Back](#chart)

### Format-Chart-Title

Formats the title from a chart.

#### Syntax

```js
chartObject.title.font.bold = true; 
chartObject.title.font.color = "#FF0000";
```

#### Properties
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`name`|String|Font name (e.g., "Calibri")|
|`size`|Double|Size of the font (e.g., 11)|
|`color`|String|HTML color code representation of the text color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |
|`italic`|Boolean|Represents the bold status of italic. true if the font style is italic|
|`bold`|Boolean|Represents the bold status of font. true if the font is bold. |
|`underline`|Boolean|Type of underline applied to the font. Can be one of the following constants. Possible Values: `None`, `Single`, `Double`, `SingleAccounting`, `DoubleAccounting`|

#### Returns
None.

#### Examples

##### Make the title of Chart1 to be bold and red

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.title.font.bold = true; 
chart.title.font.color = "#FF0000";

ctx.executeAsync().then(function () {
		logComment("Title Format Updated");
});
```
[Back](#chart)

### Set-Chart-Legend

 Hide/Show Chart Legent and set position. 

#### Syntax

```js
chartObject.legend.visible = true;
chartObject.legend.position = "top"; 
```

#### Properties
| Property         | Type    |Description| 
|:-----------------|:--------|:----------|
| `visible` | Boolean |A boolean value the represents the visibility of a ChartLegend object. If visible is set to be ture, the legend will be visible on the chart. |  
| `position` | String |Returns or sets a Legend Position value that represents the position of the legend on the chart, including `Top`,`Bottom`,`Cornor`,`Left`,`Right`,'Custom','Invalid'| 
| `overlay` | Boolean |True if the legend with be overlapping with the chart. | 


#### Returns
None.

#### Examples

##### Show Legend of Chart1 and make it on top of the chart.
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.legend.visible = true;
chart.legend.position = "top"; 
ctx.executeAsync().then(function () {
		logComment("Legend Shown ");
});
```
[Back](#chart)

### Set-Chart-DataLabels

Set display content and position of DataLabels.

#### Syntax

```js
chartObject.datalabels.visible = true;
chartObject.datalabels.position = "top";
chartObject.datalabels.ShowSeriesName = true;
```

#### Properties
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`position`          |String|Returns or sets a XlDataLabelPosition value that represents the position of the data label.  |
|`separator`         |String|Sets or returns a Variant representing the separator used for the data labels on a chart. |
|`showBubbleSize`          |Boolean|True to show the bubble size for the data labels on a chart. False to hide.|
|`showCategoryName`          |Boolean|True to display the category name for the data labels on a chart. False to hide. |
|`showLegendKey`          |Boolean|True if the data label legend key is visible.  |
|`showPercentage`          |Boolean|True to display the percentage value for the data labels on a chart. False to hide.  |
|`showSeriesName`          |Boolean|Returns or sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.  |
|`ShowValue`          |Boolean|Returns or sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.|

#### Returns
None.

#### Examples

##### Make Series Name shown in Datalabels and set the `position` of datalabels to be "top";
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.datalabels.visible = true;
chart.datalabels.position = "top";
chart.datalabels.ShowSeriesName = true;

ctx.executeAsync().then(function () {
		logComment("Datalabels Shown");
});
```
[Back](#chart)


### Set-Chart-Axis

 Set the  `maximum` ,  `minimum` ,  `majorunit` , `minorunit` of an axis. 

#### Syntax

```js
chartObject.axes.valueaxis.maximum = 5;
chartObject.axes.valueaxis.minimum = 0;
chartObject.axes.valueaxis.majorunit = 1;
chartObject.axes.valueaxis.minorunit = 0.2;
```

#### Properties
| Property         | Value    |Description|
|:-----------------|:--------|:----------|
| `minimum` | Object |Returns or sets the minimum value on the value axis. Auto if left empty.  | 
| `maximum` | Object |Returns or sets the maximum value on the value axis. Auto if left empty. | 
| `majorunit` | Object |Returns or sets the interval between two major tick marks. Auto if left empty.  | 
| `minorunit` | Object |eturns or sets the interval between two minor tick marks.  Auto if left empty. | 

#### Returns
None.

#### Examples

#####  Set the  `maximum`,  `minimum` ,  `majorunit` , `minorunit` of valueaxis. 
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.axes.valueaxis.maximum = 5;
chart.axes.valueaxis.minimum = 0;
chart.axes.valueaxis.majorunit = 1;
chart.axes.valueaxis.minorunit = 0.2;

ctx.executeAsync().then(function () {
		logComment("Axis Settings Changed");
});
```
[Back](#chart)


### Set-Chart-AxisTitle

 Change the Axis Title text and visibility. 

#### Syntax

```js

chartObject.axes.valueaxis.title.text = "Catagory";

```

#### Properties

| Property         | Type    |Description| 
|:-----------------|:--------|:----------|
| `text` | String |A String value that represents the title of a Axis. | 
| `visible` | Boolean |A boolean that specifies the visibility of an Axis Title. True if the axis or chart has a visible title.  |


#### Returns
None.

#### Examples

##### Add Catagory as the title for the catagory Axis
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.axes.valueaxis.title.text = "Catagory";

ctx.executeAsync().then(function () {
		logComment("Axis Title Added ");
});
```
[Back](#chart)

### Add-Chart-Gridlines

Show Gridlines on an Axis. 

#### Syntax

```js
chartObject.axes.valueaxis.majorgridlines.visible = true;
```

#### Parameters
None.

#### Returns
None.

#### Examples

##### Show Major Gridlines on ValueAxis of Chart1

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.axes.valueaxis.majorgridlines.visible = true;

ctx.executeAsync().then(function () {
		logComment("Axis Title Added ");
});
```
[Back](#chart)

### Format-Chart-Series

Change the Fill Color of a series.

#### Syntax

```js
chartObject.series.GetItemAt(1).fillFormat.SetSolidColor("#FF0000");
```

#### Parameters
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|color| String | HTML color code representing the color of the interior/background. |

#### Returns
None.

#### Examples

##### Change the fill color of Series1 to be red
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Charts").charts.getItem("Chart1");	

chart.series.GetItemAt(1).fillFormat.SetSolidColor("#FF0000");

ctx.executeAsync().then(function () {
		logComment("Series Fill Color Changed ");
});
```
[Back](#chart)




## Error Messages

Errors are returned using an error object that consists of a code and a message. The following table provides a list of possible error conditions that can occur. 

|error.code|error.message|
|---------:|---------:|
|InvalidArgument |The argument is invalid or missing or has an incorrect format.|
|InvalidRequest  |Cannot process the request.|
|InvalidReference|This reference is not valid for the current operation.|
|InvalidBinding  |This object binding is no longer valid due to previous updates.|
|InvalidSelection|The current selection is invalid for this operation.|
|Unauthenticated |Required authentication information is either missing or invalid.|
|AccessDenied    |You cannot perform the requested operation.|
|ItemNotFound    |The requested resource doesn't exist.|
|InvalidMethod   | The method in the request is not allowed on the resource. |
|EditConflict    |Request could not be processed because of conflict.|
|ActivityLimitReached|Activity limit has been reached.|
|GeneralException|There was an internal error while processing the request.|
|NotImplemented  |The requested feature isn't implemented.|
|ServiceNotAvailable|The service is unavailable.|

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

## Programming Notes

Following sections provide important programming details related to Excel APIs.

* [Properties and Relations Selection](#properties-and-relations-selection)
* [Document Binding](#null-input)
* [Reference Binding](#null-input)
* [Null-Input](#null-input)
* [Null-Input](#null-input)
* [Null-Response](#null-response)
* [Blank Input and Output](#blank-input-and-output)
* [Unbounded-Range](#unbounded-range)
* [Large-Range](#large-range)
* [Single Input Copy](#single-input-copy)
* [Throttling](#throttling)

[top](#excel-javascript-apis)

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
context.load (<object-var>, select: []);
context.load (<object-var>, select: "comma separated list of properties");
```

#### Examples

```js
var sheetName = "Sheet1";
var rangeAddress = "A1:B2";
var ctx = new Excel.ExcelClientContext();
var myRange = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

//load statement below loads the address, values, numberFormat properties of the Range and then expands on the format, format/background, entireRow relations
 
ctx.load (myRange, select: ["address", "values", "numberFormat", format, format/background, entireRow ]);

ctx.executeAsync().then(function () {
		console.log (myRange.address); //ok
		console.log (myRange.cellCount); //not-ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.background.color); //ok
		console.log (myRange.format.font.color); //not-ok
		console.log (myRange.entireRow.address); //ok
		console.log (myRange.entireColumn.address); //not-ok
// . . . 

//load statement below loads all the properties of the Range and then expands on the format, format/background, entireRow relations. If the "*" is left out of the load, none of the Range’s direct properties will be included in the load statement.
 
ctx.load (myRange, select: ["*", "format", "format/background", "entireRow" ]);

ctx.executeAsync().then(function () {
		console.log (myRange.address); //ok
		console.log (myRange.cellCount); //ok
		console.log (myRange.format.wrapText); //ok
		console.log (myRange.format.background.color); //ok
		console.log (myRange.format.font.color); //not-ok
		console.log (myRange.entireRow.address); //ok
		console.log (myRange.entireColumn.address); //not-ok

```

[Back](#programming-notes)
### Document Binding

[Back](#programming-notes)
### Reference Binding

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
ctx.load(range);
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
ctx.load(range);
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
ctx.load(range);
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
ctx.load(range);
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

[top](#excel-javascript-apis)