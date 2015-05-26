# AxisTitle
Represents the title of a specified axis.


## JSON representation

JSON representation of a Range resource.
<!-- { "blockType": "resource", "@odata.type": "ChartAxisTitle", 
		"optionalProperties": [ "fillFormat", "lineFormat", "font" ]
	 } 
-->
```json
{
  "text" : "Date",
  "visible" : true,

  "fillFormat" :    {"@odata.type": "ChartFillFormat"},
  "lineformat" :    {"@odata.type": "ChartLineFormat"},
  "font" :    {"@odata.type": "ChartFont"}

}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `text` | String |A String value that represents the title of a Axis. | Chart.ChartTitle |
| `visible` | Boolean |A boolean that specifies the visibility of an Axis Title. True if the axis or chart has a visible title.  | Axis.HasTitle |




## Relationships
The ChartAxisTitle resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `fillFormat`          |[ChartFillFormat](chartFillFormat.md) Object | Represents the fill format of an object, which includes interior/background formating information. 
| `lineFormat`          |[ChartLineFormat](chartLineFormat.md) Object | Represents line and arrowhead formatting.
| `font`          |[ChartFont](chartFont.md) Object | Represents the font attributes (font name, font size, color, and so on) for an object. 

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.