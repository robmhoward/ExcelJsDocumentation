# Axix
Represents a single axis in a chart.

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `minimum` | Object |Returns or sets the minimum value on the value axis. Auto if left empty.  | Axis.MinimumScale|
| `maximum` | Object |Returns or sets the maximum value on the value axis. Auto if left empty. | Axis.MaximumScale|
| `majorunit` | Object |Returns or sets the interval between two major tick marks. Auto if left empty.  | Axis.majorunit|
| `minorunit` | Object | Returns or sets the interval between two minor tick marks. Auto if left empty. | Axis.minorunit|


## Relationships
The Chart resource has the following relationships defined:

| Relationships    | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `title`          |[ChartAxisTitle](chartAxisTitle.md) Object | Represents the title of a specified axis. | Axis.AxisTitle
| `majorGridlines` | [ChartGridlines](chartGridlines.md) Object   | Returns a Gridlines object that represents the major gridlines for the specified axis.   | Axis.MajorGridlines|
| `minorGridlines` | [ChartGridlines](chartGridlines.md) Object   | Returns a Gridlines object that represents the minor gridlines for the specified axis.  | Axis.MinorGridlines|
| `font`          |[ChartGridlines](chartFont.md) Object | Represents the font attributes (font name, font size, color, and so on) for an object. 

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.