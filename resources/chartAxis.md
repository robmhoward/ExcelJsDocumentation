# Axix
Represents a single axis in a chart.


## JSON representation

JSON representation of a Range resource.
<!-- { "blockType": "resource", "@odata.type": "ChartAxis", 
	"optionalProperties": ["title", "majorGridlines", "minorGridlines", "font"]
	 } 
-->
```json
{
  
  "minimum" : 0,
  "maximum" : 100,
  "majorUnit": 5,
  "majorUnit": 1,
  "visible": true,

  "title" :    {"@odata.type": "ChartAxisTitle"} ,
  "majorGridlines" : {"@odata.type": "ChartGridlines"} ,
  "minorGridlines"  : { "@odata.type" : "ChartGridlines" },
  "font" :    {"@odata.type": "ChartFont"}

}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `minimum` | Object |Returns or sets the minimum value on the value axis. Auto if left empty.  | Axis.MinimumScale|
| `maximum` | Object |Returns or sets the maximum value on the value axis. Auto if left empty. | Axis.MaximumScale|
| `majorunit` | Object |Returns or sets the interval between two major tick marks. Auto if left empty.  | Axis.majorunit|
| `minorunit` | Object |eturns or sets the interval between two minor tick marks.  Auto if left empty. | Axis.minorunit|
| `visible` | Boolean |True if the Axis is displayed. Read/write Boolean. | msoElementPrimaryCategoryAxisShow |


## Relationships
The Chart resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `title`          |[ChartAxisTitle](chartAxisTitle.md) Object | represents the title of the specified axis. | Axis.AxisTitle
| `majorGridlines` | [ChartGridlines](chartGridlines.md) Object   | Returns a Gridlines object that represents the major gridlines for the specified axis.   | Axis.MajorGridlines|
| `minorGridlines` | [ChartGridlines](chartGridlines.md) Object   |AReturns a Gridlines object that represents the minor gridlines for the specified axis.  | Axis.MinorGridlines|
| `font`          |[ChartGridlines](chartFont.md) Object | Represents the font attributes (font name, font size, color, and so on) for an object. 

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.