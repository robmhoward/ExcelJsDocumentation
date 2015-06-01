# AxisTitle
Represents the title of a specified axis.

## [Properties](#get-chart-axis-title)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `text` | String |A String value that represents the title of a Axis. | 
| `visible` | Boolean |A boolean that specifies the visibility of an Axis Title. True if the axis or chart has a visible title.  | 

## Relationships
The ChartAxisTitle resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `format`          |[ChartAxisTitleFormat](chartAxisTitleFormat.md) Object | Represents the format of a chart axis title font formatting.

## Methods
None.


## API Specification 

### Get Chart Axis Title

Gets a ChartAxisTitle object.

#### Syntax
Use value axis as an example here.

```js
chartObject.axes.valueaxis.title;
```
#### Properties
| Property         | Type    |Description| 
|:-----------------|:--------|:----------|
| `text` | String |A String value that represents the title of a Axis. | 
| `visible` | Boolean |A boolean that specifies the visibility of an Axis Title. True if the axis or chart has a visible title.  |

#### Returns

[ChartAxisTitle](chartAxisTitle.md) object. 

#### Examples

##### Get the `text` of Chart Axis Title from the value axis of Chart1.
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var title = chart.axes.valueaxis.title;
ctx.load(title);
ctx.executeAsync().then(function () {
		logComment(title.text);
});
```

[Back](#properties)

### Set Chart Axis Title

Set chart axis title properties including text and visibility.

#### Syntax
Use value axis as an example here.
```js
chartObject.axes.valueaxis.title.text= "My Chart"; 
chartObject.axes.valueaxis.title.visible = true;
```

#### Properties
| Property         | Type    |Description| 
|:-----------------|:--------|:----------|
| `text` | String |A String value that represents the title of a Axis. | 
| `visible` | Boolean |A boolean that specifies the visibility of an Axis Title. True if the axis or chart has a visible title.  |

#### Returns

[ChartAxisTitle](chartAxisTitle.md) object. 


#### Examples

##### Add "Values" as the title for the value Axis
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.axes.valueaxis.title.text = "Values";

ctx.executeAsync().then(function () {
		logComment("Axis Title Added ");
});
```
[Back](#properties)