# Format

Format object represents format settings of a Range. This includes Font, Background, Borders, Alignment, Style, etc. 

## Properties
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`font`            |[Font](font.md) Object                |Returns the Font object defined on the overall Range selected|Range.Font|
|`background`      |[Background](background.md) Object    |Returns the Background object defined on the overall Range selected|Range.Interior|
|`horizontalAlignment`    | String  |Represents the horizontal alignment for the specified object. The value of this property can be to one of the following constants: `General`, `Fill`, `CenterAcrossSelection`, `Center`, `Distributed`, `Justify`, `Left`, `Right`. `null` indicates that the entire range doesn't have uniform horizontal alignment.|Range.HorizontalAlignment|
|`verticalAlignment`    | String  |Represents the vertical alignment for the specified object. The value of this property can be to one of the following constants: `Bottom`, `Center`, `Distributed`, `Justify`, `Top`. `null` indicates that the entire range doesn't have uniform vertical alignment.|Range.VerticalAlignment|
|`wrapText`    | Boolean  |Indicates if Excel wraps the text in the object. `null` indicates that the entire range doesn't have uniform wrap setting|Range.WrapText|


## Relationships
## Properties
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`borders`         |[Border](border.md) collection|Collection of border objects that apply to the overall Range selected|Range.Borders|

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.