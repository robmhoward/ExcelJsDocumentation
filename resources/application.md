# Application

Represents the Excel application which is managing the workbook. 

## JSON representation 

JSON representation of a Workbook resource

<!-- { "blockType": "resource", "@odata.type": "Application"]
	 } 
-->
```json
{
  "calculationMode":  "String"
}
```

## Properties

| Property         | Type    |Description|Notes |
|:-----------------|:--------|:----------|:-----|
| `calculationMode`        | String      | Specifies the Calculation mode of the workbook. Possible values are `Automatic`: Excel controls recalculation, `Manual`: Calculation is done when the user requests it, or `Semiautomatic`: Excel controls recalculation but ignores changes in tables.         |Workbook.Application.Calculation|


## Relationships
None

## Methods

The complete list of methods for this resource is available in
the [API](../README.md) topic.