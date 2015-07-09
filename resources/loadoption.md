# LoadOption

An object that can be passed to the load method to specify the options such as select, expand parameters. 

## Properties
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|select|object|Provide comma delimited list or an array of parameter/relationship names to be loaded upon executeAsync call. e.g. "property1, relationship1", [ "property1", "relationship1"]. Optional.||
|expand|object|Provide comma delimited list or an array of relationship names to be loaded upon executeAsync call. e.g. "relationship1, relationship2", [ "relationship1", "relationship2"]. Optional.||
|top|int| Specify the number of items in the queried collection to be included in the result. Optional.||
|skip|int|Specify the number of items in the collection that are to be skipped and not included in the result. If `top` is specified, the selection of result will start after skipping the specified number if items. Optional.||

