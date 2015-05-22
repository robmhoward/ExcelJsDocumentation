# Series
Represents a series in a chart.


## JSON representation

JSON representation of a Range resource.
<!-- { "blockType": "resource", "@odata.type": "ChartSeries", 
	"optionalProperties": ["points", "fillFormat"]
	 } 
-->
```json
{
  "name" : "Series1",

  "points" :    {"@odata.type": "ChartPoints"},
  "fillFormat" :    {"@odata.type": "ChartFillFormat"}
}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`name`          |String|A String value that represents a Series object |Series.Name|

## Relationships
The ChartSeries resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `points`          |[ChartPoints](chartPoints.md) Object | Represents the Points in a series in a chart.
| `fillFormat`          |[ChartFillFormat](chartFillFormat.md) Object | Represents the fill format of an object, which includes background formating information. 


## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.