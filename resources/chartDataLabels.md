# DataLabels
Represents the data label on a chart point or trendline.

## [Properties](#set-chart-datalabels)

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`position`          |String|Returns or sets a DataLabelPosition value that represents the position of the data label. Valid position for DataLabels are: "Invalid", "None", "Center", "InsideEnd", "InsideBase", "OutsideEnd","Left", "Right", "Top","Bottom", "BestFit", "Callout". |DataLabel.Position|
|`separator`         |String|Returns or setsor returns a String representing the separator used for the data labels on a chart. |DataLabel.separator|
|`showBubbleSize`          |Boolean|True to show the bubble size for the data labels on a chart. False to hide.|DataLabel.showBubbleSize|
|`showCategoryName`          |Boolean|True to display the category name for the data labels on a chart. False to hide. |DataLabel.showCategoryName|
|`showLegendKey`          |Boolean|True if the data label legend key is visible.  |DataLabel.showLegendKey|
|`showPercentage`          |Boolean|True to display the percentage value for the data labels on a chart. False to hide.  |DataLabel.showPercentage|
|`showSeriesName`          |Boolean|True to display the series name for the data labels on a chart. False to hide. |DataLabel.showSeriesName|
|`ShowValue`          |Boolean|True to display the value for the data labels on a chart. False to hide.|DataLabel.ShowValue|


## Relationships
The ChartDataLabels resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `format`          |[ChartDataLabelsFormat](chartDataLabelFormat.md) Object | Represents the format of chart datalabels, which includes fill(interior/background) and font formatting.

## Methods
None.

## API Specification 


### Set Chart DataLabels

Set chart datalables properties.

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
var chart = ctx.workbook.worksheets.getItem("Sheet1").charts.getItem("Chart1");	

chart.datalabels.visible = true;
chart.datalabels.position = "top";
chart.datalabels.ShowSeriesName = true;

ctx.executeAsync().then(function () {
		logComment("Datalabels Shown");
});
```
[Back](#properties)