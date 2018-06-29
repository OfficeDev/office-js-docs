# HeaderFooterGroup Object (JavaScript API for Excel)

Represents the options in page layout margins.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|useSheetMargins|bool|Gets or sets a flag indicating if headersfooters are aligned with the page margins set in the page layout options for the worksheet.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|useSheetScale|bool|Gets or sets a flag indicating if headersfooters should be scaled by the page percentage scale set in the page layout options for the worksheet.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|defaultForAllPages|[HeaderFooter](headerfooter.md)|The general headerfooter, used for all pages unless evenodd or first page is specified. Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|evenPages|[HeaderFooter](headerfooter.md)|The headerfooter to use for even pages, odd headerfooter needs to be specified for odd pages. Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|firstPage|[HeaderFooter](headerfooter.md)|The first page headerfooter, for all other pages general or evenodd is used. Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|oddPages|[HeaderFooter](headerfooter.md)|The headerfooter to use for odd pages, even headerfooter needs to be specified for even pages. Read-only.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|
|state|string|Gets or sets the state of which headersfooters are set.|[ApiSet.InProgressFeatures.PageLayout](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

