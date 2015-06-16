# Chart Data Labels
Represents a colection of all the data labels on a chart point or trendline.

## [Properties](#set-chart-datalabels)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`position`          |String|DataLabelPosition value that represents the position of the data label. Valid position for DataLabels are: "Invalid", "None", "Center", "InsideEnd", "InsideBase", "OutsideEnd","Left", "Right", "Top","Bottom", "BestFit", "Callout". |DataLabel.Position|
|`separator`         |String|String representing the separator used for the data labels on a chart. |DataLabel.separator|
|`showBubbleSize`          |Boolean|Set to true to show the bubble size for the data labels on a chart. Set to false to hide.|DataLabel.showBubbleSize|
|`showCategoryName`          |Boolean|Set to true to display the category name for the data labels on a chart. Set to false to hide. |DataLabel.showCategoryName|
|`showLegendKey`          |Boolean|True if the data label legend key is visible.  |DataLabel.showLegendKey|
|`showPercentage`          |Boolean|Set to true to display the percentage value for the data labels on a chart. Seto to false to hide.  |DataLabel.showPercentage|
|`showSeriesName`          |Boolean|Set to true to display the series name for the data labels on a chart. Set to false to hide. |DataLabel.showSeriesName|
|`ShowValue`          |Boolean|Set to true to display the value for the data labels on a chart. Set to false hide.|DataLabel.ShowValue|


## Relationships

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `format`          |[Chart Data Label Format](chartDataLabelFormat.md) object | Represents the format of chart datalabels, which includes fill(interior/background) and font formatting.

## Methods
None.

## API Specification 


### Set Chart DataLabels

Set the properties of the chart datalables.

#### Syntax

```js
chartObject.datalabels.visible = true;
chartObject.datalabels.position = "top";
chartObject.datalabels.ShowSeriesName = true;
```

#### Properties
| Property         | Type    |Description|
|:-----------------|:--------|:----------|
|`position`          |String|DataLabelPosition value that represents the position of the data label. Valid position for DataLabels are: "Invalid", "None", "Center", "InsideEnd", "InsideBase", "OutsideEnd","Left", "Right", "Top","Bottom", "BestFit", "Callout". | 
|`separator`         |String|String representing the separator used for the data labels on a chart. | 
|`showBubbleSize`          |Boolean|Set to true to show the bubble size for the data labels on a chart. Set to false to hide.| 
|`showCategoryName`          |Boolean|Set to true to display the category name for the data labels on a chart. Set to false to hide. | 
|`showLegendKey`          |Boolean|True if the data label legend key is visible.  |
|`showPercentage`          |Boolean|Set to true to display the percentage value for the data labels on a chart. Seto to false to hide.  |
|`showSeriesName`          |Boolean|Set to true to display the series name for the data labels on a chart. Set to false to hide. |
|`ShowValue`          |Boolean|Set to true to display the value for the data labels on a chart. Set to false hide.|
#### Returns
None.


#### Examples
##### Make Series Name shown in Datalabels and set the `position` of datalabels to be "top";
```js
var ctx = new Excel.ExcelClientContext();
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.datalabels.visible = true;
chart.datalabels.position = "top";
chart.datalabels.ShowSeriesName = true;

ctx.executeAsync().then(function () {
		logComment("Datalabels Shown");
});
```
[Back](#properties)