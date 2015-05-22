# ChartTitle
Represents a chart title object of a chart. A ChartTitle object specifies the text, visibility, position and formating of the chart title.


## JSON representation

JSON representation of a ChartTile resource.
<!-- { "blockType": "resource", "@odata.type": "ChartTitle", 
	"optionalProperties": [ "fillFormat", "lineFormat", "font" ]
	 } 
-->
```json
{
  "text" : "Revenue By Quarter",
  "visible": true,
  "position" : "Top",
  "overlay" : false,

  "fillFormat" :    {"@odata.type": "ChartFillFormat"},
  "lineformat" :    {"@odata.type": "ChartLineFormat"},
  "font" :    {"@odata.type": "ChartFont"}

}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `text` | String |A String value that represents the title text of a chart. When a title text is set, the display property will be automaticlly set to top and the chart title will be displayed on top of the chart without overlapping. | Chart.ChartTitle |
| `visible` | Boolean |A boolean value the represents the visibility of a chart title object. If visible is set to be ture, the chart title will be visible on the chart. |  |
| `position | String | A constant that specifies the postition of chart title, including `Top`, `None` and `Invalid`. | Chart.ChartTitle.Position |
| `overlay` | Boolean |True if the title overlays the chart. | Chart.ChartTitle.Position |



## Relationships
The ChartTitle resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `fillFormat`          |[ChartFillFormat](chartFillFormat.md) Object | Represents the fill format of an object, which includes background formating information. 
| `lineFormat`          |[ChartLineFormat](chartLineFormat.md) Object | Represents line and arrowhead formatting.
| `font`          |[ChartFont](chartFont.md) Object | Represents the font attributes (font name, font size, color, and so on) for an object. 


     

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.