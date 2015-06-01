# Gridlines
Represents major or minor gridlines on a chart axis.

## [Properties](#get-chart-gridlines)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|visible| Boolean | True if the axis has gridlines. ||

## Relationships
The ChartGridlines resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `format`          |[ChartGridlinesFormat](chartGridlinesrangeformat.md) Object | Represents the format of chart gridlines.
          

## Methods
None.

## API Specification 
### Get Chart Gridlines

Gets a ChartGridlines object.

#### Syntax
Use major gridlines on value axis as an example here.

```js
chartObject.axes.valueaxis.majorGridlines;
```
#### Properties
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|visible| Boolean | True if the axis has gridlines. |

#### Returns

[ChartGridlines](resources/chartGridlines.md) object. 

#### Examples

##### Get the `visible` of Major Gridlines on value axis of Chart1
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

var majGridlines = chart.axes.valueaxis.majorGridlines;
ctx.load(majGridlines);
ctx.executeAsync().then(function () {
		logComment(majGridlines.visible);
});
```

[Back](#properties)

### Set Chart Gridlines

Show Gridlines on an Axis. 

#### Syntax
Use major gridlines on value axis as an example here.
```js
chartObject.axes.valueaxis.majorgridlines.visible = true;
```

#### Properties
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|visible| Boolean | True if the axis has gridlines. |

#### Returns
[ChartGridlines](resources/chartGridlines.md) object. 

#### Examples

##### Show Major Gridlines on ValueAxis of Chart1

```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.axes.valueaxis.majorgridlines.visible = true;

ctx.executeAsync().then(function () {
		logComment("Axis Gridlines Added ");
});
```
[Back](#properties)