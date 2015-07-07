# Application

Represents the Excel application that manages the workbook.

## [Properties](#getter-examples)
| Property       | Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|calculationMode|string|Returns the calculation mode used in the workbook. Read-only. Possible values are: `Automatic` Excel controls recalculation.,`AutomaticExceptTables` Excel controls recalculation but ignores changes in tables.,`Manual` Calculation is done when the user requests it.||

## Relationships
None


## Methods

| Method           | Return Type    |Description|Notes |
|:---------------|:--------|:----------|:-----|
|[calculate(calculationType: string)](#calculatecalculationtype-string)|void|Recalculate all currently opened workbooks in Excel.||

## API Specification

### calculate(calculationType: string)
Recalculate all currently opened workbooks in Excel.

#### Syntax
```js
applicationObject.calculate(calculationType);
```

#### Parameters
| Parameter       | Type    |Description|
|:---------------|:--------|:----------|
|calculationType|string|Specifies the calculation type to use. Possible values are: `Recalculate` Default-option. Performs normal calculation by calculating all the formulas in the workbook.,`Full` Forces a full calculation of the data.,`FullRebuild`  Forces a full calculation of the data and rebuilds the dependencies.|

#### Returns
void

#### Examples
```js
var ctx = new Excel.ExcelClientContext();
ctx.workbook.application.calculate('Full');
ctx.executeAsync().then();
```
[Back](#methods)

#### Getter Examples
```js
var ctx = new Excel.ExcelClientContext();
var application = ctx.workbook.application;
ctx.load(application);
ctx.executeAsync().then(function() {
	Console.log(application.calculationMode);
});
```

[Back](#properties)
