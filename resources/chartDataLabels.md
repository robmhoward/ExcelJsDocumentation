# DataLabels
Represents the data label on a chart point or trendline.


## JSON representation

JSON representation of a Range resource.
<!-- { "blockType": "resource", "@odata.type": "ChartDataLabels", 
	"optionalProperties": [ "fillFormat", "lineFormat", "font" ]
	 } 
-->
```json
{
  "position" : "InsideEnd",
  "separator" : ",",
  "showBubbleSize" : false,
  "showCategoryName" : false,
  "showLegendKey" : false,
  "showPercentage" :false ,
  "showSeriesName" : true,
  "ShowValue" : true,

  "fillFormat" :    {"@odata.type": "ChartFillFormat"},
  "lineformat" :    {"@odata.type": "ChartLineFormat"},
  "font" :    {"@odata.type": "ChartFont"}

}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`position`          |String|Returns or sets a XlDataLabelPosition value that represents the position of the data label.  |DataLabel.Position|
|`separator`         |String|Sets or returns a Variant representing the separator used for the data labels on a chart. |DataLabel.separator|
|`showBubbleSize`          |Boolean|True to show the bubble size for the data labels on a chart. False to hide.|DataLabel.showBubbleSize|
|`showCategoryName`          |Boolean|True to display the category name for the data labels on a chart. False to hide. |DataLabel.showCategoryName|
|`showLegendKey`          |Boolean|True if the data label legend key is visible.  |DataLabel.showLegendKey|
|`showPercentage`          |Boolean|True to display the percentage value for the data labels on a chart. False to hide.  |DataLabel.showPercentage|
|`showSeriesName`          |Boolean|Returns or sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.  |DataLabel.showSeriesName|
|`ShowValue`          |Boolean|Returns or sets a Boolean corresponding to a specified chart's data label values display behavior. True displays the values. False to hide.|DataLabel.ShowValue|


Valid position for DataLabels are: "Invalid", "None", "Center", "InsideEnd", "InsideBase", "OutsideEnd","Left", "Right", "Top","Bottom", "BestFit", "Callout".





## Relationships
The ChartDataLabels resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `fillFormat`          |[ChartFillFormat](chartFillFormat.md) Object | Represents the fill format of an object, which includes background formating information. 
| `lineFormat`          |[ChartLineFormat](chartLineFormat.md) Object | Represents line and arrowhead formatting.
| `font`          |[ChartFont](chartFont.md) Object | Represents the font attributes (font name, font size, color, and so on) for an object. 



## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.