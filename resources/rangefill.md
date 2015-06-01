# Range Fill

Represents the interior of an object, which includes fill formating information. 

## Properties
| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
|`color`|String|HTML color code representation of the fill color. HTML color codes are strings that represents hexadecimal triplets of red, green, and blue values (#RRGGBB). e.g., `#FF0000` represents Red. ('255' red, '0' green, and '0' blue) |Conversion from Range.Interior.Color value to html color string|

## Programming notes about `color` property: 

A `color` hex code is a way of specifying color using hexadecimal values. The code itself is a hex triplet, which represents three separate values that specify the levels of the component colors. The code starts with a pound sign (#) and is followed by six hex values or three hex value pairs (for example, #AFD645). 

Of the 6 Hex values, first two characters represent the values 0 through 255 for red in hex; the middle two for green and the last two for blue (#RRGGBB). For example, FF is equal to 255. Therefore, the purest white obtainable is the highest intensity of red, green and blue, which is #FFFFFF (red=255, green=255 and blue=255). Black is the lack of all RGB (#0000000).

When `color` value is updated, the input value needs to follow the appropriate formatting as mentioned above. The Alpha characters of the hex color code can be lower or upper case. 


## Relationships
None

## Methods
None

