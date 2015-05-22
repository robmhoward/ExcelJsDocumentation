# Named Item

Represents a defined name for a range of cells or value. Names can be primitive named objects (as seen in the type below), range object, reference to a range.
This object can be used to obtain Range object associated with names.

## JSON representation  

JSON representation of a Named Item resource.

<!-- { "blockType": "resource", "@odata.type": "NamedItem", 
		"optionalProperties": ["range"],	 
	 } 
-->
```json
{
  "name" : "String",
  "value" : "String",
  "visible": true,
  "type" : "String",

  "range" : {"@odata.type": "Range"}
}

## Properties

| Property         | Type    |Description|Maps to (in VBA) |
|:-----------------|:--------|:----------|:-----|
| `name`  | String|String value representing the name of the object.| Name.Name|
| `value`| String |Represents the formula that the name is defined to refer to. e.g., `=Sheet14!$B$2:$H$12`, `=4.75`, etc. | Name.Value|
| `visibile` | Boolean |Boolean value that determines whether the object is visible. | Name.Visible |
| `type` | String|Indicates what type of reference is associated with the name. Possible options are: `Range`, `String`, `Integer`, `Double`, `Boolean`. | Derived property |
| `range` | Range Object|Range object that is associated with the name. `null` if the name is not of the type `Range`.| Name.RefersTo (Range object derived based on the Range reference) |

## Relationships
None
     
## Methods

The complete list of methods for this resource is available in the [API](../README.md) topic.
