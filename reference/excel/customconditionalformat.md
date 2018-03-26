# CustomConditionalFormat Object (JavaScript API for Excel)

Represents a custom conditional format type.

## Properties

None

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|format|[ConditionalRangeFormat](conditionalrangeformat.md)|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|
|rule|[ConditionalFormatRule](conditionalformatrule.md)|Represents the Rule object on this conditional format. Read-only.|[1.6](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None


## Examples
```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("A1:A5");
    range.values = [[1], [20], [""], [5], ["test"]];
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    var cfCustom = cf.customOrNullObject;
    cfCustom.rule.formula = "=ISBLANK(A1)";
    cfCustom.format.fill.color = "#00FF00";
    return ctx.sync().then(function () {
        console.log("Added new custom conditional format highlighting all blank cells.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```