# Legend
Represents the legend in a chart. Each chart can have only one legend.


## JSON representation

JSON representation of a Range resource.
<!-- { "blockType": "resource", "@odata.type": "ChartLegend", 
	"optionalProperties":  [ "fillFormat", "lineFormat", "font" ]
	 } 
-->
```json
{
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
| `visible` | Boolean |A boolean value the represents the visibility of a ChartLegend object. If visible is set to be ture, the legend will be visible on the chart. |  |
| `position` | String |Returns or sets a Legend Position value that represents the position of the legend on the chart, including `Top`,`Bottom`,`Cornor`,`Left`,`Right`,'Custom','Invalid'| Legend.position |
| `overlay` | Boolean |True if the legend with be overlapping with the chart. | Legend.IncludeInLayout |


## Relationships
The Chart Legend resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `fillFormat`          |[ChartFillFormat](chartFillFormat.md) Object | Represents the fill format of an object, which includes interior/background formating information. 
| `lineFormat`          |[ChartLineFormat](chartLineFormat.md) Object | Represents line and arrowhead formatting.
| `font`          |[ChartFont](chartFont.md) Object | Represents the font attributes (font name, font size, color, and so on) for an object. 


     

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.