# ChartTitle
Represents a chart title object of a chart. A ChartTitle object specifies the text, visibility and formating of the chart title.

## [Properties](#get-chart-title)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `text` | String |A String value that represents the title text of a chart. When a title text is set, the display property will be automaticlly set to top and the chart title will be displayed on top of the chart without overlapping. | Chart.ChartTitle |
| `visible` | Boolean |A boolean value the represents the visibility of a chart title object. If visible is set to be ture, the chart title will be visible on the chart. |  |
| `overlay` | Boolean |True if the title overlays the chart. | Chart.ChartTitle.Position |

## Relationships
The ChartTitle resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `format`          |[ChartTitleFormat](chartTitlerangeformat.md) Object | Represents the format of a chart title, which includes fill(interior/background) and font formatting.
     
## Methods
None.

## API Specification 

### Get Chart Title

Gets a ChartTitle object.

#### Syntax
```js
chartObject.title;
```
#### Properties
| Property         | Type    |Description| 
|:-----------------|:--------|:----------|
| `text` | String |A String value that represents the title text of a chart. When a title text is set, the display property will be automaticlly set to top and the chart title will be displayed on top of the chart without overlapping. | 
| `visible` | Boolean |A boolean value the represents the visibility of a chart title object. If visible is set to be ture, the chart title will be visible on the chart. |  |
| `overlay` | Boolean |True if the title overlays the chart. | 

#### Returns

[ChartTitle](resources/chartTitle.md) object. 

#### Examples

##### Get the `text` of Chart Title from Chart1
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var title = chart.title;
ctx.load(title);
ctx.executeAsync().then(function () {
		logComment(title.text);
});
```

[Back](#properties)

### Set Chart Title

Set chart title properties including text and visibility.

#### Syntax

```js
chartObject.title.text= "My Chart"; 
chartObject.title.visible=true;
chartObject.title.overlay=true;
```

#### Properties
| Property         | Type    |Description| 
|:-----------------|:--------|:----------|
| `text` | String |A String value that represents the title text of a chart. When a title text is set, the display property will be automaticlly set to top and the chart title will be displayed on top of the chart without overlapping. | 
| `visible` | Boolean |A boolean value the represents the visibility of a chart title object. If visible is set to be ture, the chart title will be visible on the chart. |  |
| `overlay` | Boolean |True if the title overlays the chart. | 

#### Returns

[ChartTitle](resources/chartTitle.md) object. 


#### Examples

##### Set the `text` of Chart Title to "My Chart" and Make it show on top of the chart without overlaying.
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.title.text= "My Chart"; 
chart.title.visible=true;
chart.title.overlay=true;

ctx.executeAsync().then(function () {
		logComment("Char Title Changed");
});
```
[Back](#properties)