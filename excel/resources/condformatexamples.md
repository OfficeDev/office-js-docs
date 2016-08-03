# Conditional format examples

Below are list of conditional format examples: 

### getItemAt(index: int)
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormats = range.conditionalFormats;
    var conditionalFormat = conditionalFormats.getItemAt(3);
    return ctx.sync().then(function () {
        console.log("Conditional Format at Item 3 Loaded");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### clearAll()
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormats = range.conditionalFormats;
    var conditionalFormat = conditionalFormats.clearAll();
    return ctx.sync().then(function () {
        console.log("Cleared all conditional formats from this range.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### add(type: ConditionalFormatType)
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    conditionalFormat.iconOrNull.style = "YellowThreeArrows";
    return ctx.sync().then(function () {
        console.log("Added new yellow three arrow icon set.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### getCount()
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    conditionalFormat.iconOrNull.style = Excel.IconSet.fourTrafficLights;
    var cfCount = range.conditionalFormats.getCount(); 

    return ctx.sync().then(function () {
        console.log("Count: " + cfCount.value);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
### Getters and Setters
#### Get Priority
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.iconOrNull.style = Excel.IconSet.threeArrows;

    cf.load('priority');
    return ctx.sync().then(function () {
        console.log("Local Priority" + cf.priority);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

#### StopIfTrue, get
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.iconOrNull.style = Excel.IconSet.threeArrows;

    cf.load('stopIfTrue');
    return ctx.sync().then(function () {
        console.log("Stop If True?" + cf.stopIfTrue);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

#### Set Reverse
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.reverse = true;
    return ctx.sync().then(function () {
        console.log("Reverse: true?" + cf.reverse);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
#### Get Type
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.load('type');
    return ctx.sync().then(function () {
        console.log("Error: " + Error);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### Add Custom IconSet
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet)
    conditionalFormat.iconOrNull.criteria = [
        { type: "percent", formula: 10, operator: "lessThan", customIcon: Excel.icons.fiveArrows.yellowDownInclineArrow },
        { type: "percent", formula: 30, operator: "lessThan", customIcon: Excel.icons.fourRating.fourBars },
        { type: "percent", formula: 50, operator: "lessThan", customIcon: null },
        { type: "percent", formula: 70, operator: "lessThan", customIcon: Excel.icons.fiveQuarters.blackCircle }
    ];
    
    return ctx.sync().then(function () {
        console.log("Added new custom icon conditional formatting.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Add Preset IconSet
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet).iconOrNull.style = Excel.IconSet.threeArrows;
    return ctx.sync().then(function () {
        console.log("Added new yellow three arrow icon set.");
    })
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```
#### Range Success:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    var rangeVal = cf.getRangeOrNull();
  
    return ctx.sync().then(function () {
        console.log("Range: " + rangeVal);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

#### Range = NULL
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);

    /// remove conditional format from just A2
    var subsetRangeAddress = "A2";
    var subsetRange = ctx.workbook.worksheets.getItem(sheetName).getRange(subsetRangeAddress);
    range.conditionalFormats.getItemAt(0).deleteFromCurrentRange();

    var rangeOnConditionalFormat = cf.getRangeOrNull();
    return ctx.sync().then(function () {
        console.log("Range is Discontiguous:" + rangeOnConditionalFormat);
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```



### Add Preset Color Scale
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    cf.colorScaleOrNull.minimum.color = "red";
    cf.colorScaleOrNull.midpoint.color = "yellow";
    cf.colorScaleOrNull.maximum.color = "green"; 

    return ctx.sync().then(function () {
        console.log("Added red-yellow-green color scale.");
    })
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### Add Custom Color Scale
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.colorScale);
    cf.colorScaleOrNull.minimum = {
        type: Excel.ConditionalFormatRuleType.percent,
        formula: 10,
        color: "red"
    };
    cf.colorScaleOrNull.maximum = {
        type: Excel.ConditionalFormatRuleType.percent,
        formula: 90,
        color: "blue"
    };
    return ctx.sync().then(function () {
        console.log("Added new red-blue 10%-90% color scale format");
    })
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

## ConditionalFormatting Custom Types

### Average:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);

    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.average;
    conditionalFormat.customOrNull.rule.average.selection = Excel.ConditionalFormatAverageSelection.below;
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.bad;
    return ctx.sync().then(function () {
        console.log("Added new below average rule type, formatted bad.");
	});
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Between: 
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.between;
    conditionalFormat.customOrNull.rule.between = {
        inclusive: true,
        lowerBound: "A2",
        upperBound: "B3"
    };
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.good;
    return ctx.sync().then(function () {
        console.log("Added new between rule based on cells A2 and B3, formatted good");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### Formula: 
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.formula;
    conditionalFormat.customOrNull.rule.formula = "=COUNTIF($D$2:$D11,D2)>1";
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.neutral;
    return ctx.sync().then(function () {
        console.log("Added new custom formula rule with neutral formatting");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Top/Bottom Count:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.count;
    conditionalFormat.customOrNull.rule.count = {
        count: 10,
        direction: Excel.ConditionalFormatDirection.top
    };
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.heading1;
    return ctx.sync().then(function () {
        console.log("Added new top 10 items formatted like heading1");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Top/Bottom Percent:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.percent;
    conditionalFormat.customOrNull.rule.percent = {
        percent: 10,
        direction: Excel.ConditionalFormatDirection.bottom
    };
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.title;
    return ctx.sync().then(function () {
        console.log("Added new conditional format on bottom 10% formatted like a title");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### TextContains:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.textContains;
    conditionalFormat.customOrNull.rule.textContains = {
        type: Excel.StringMatchType.beginningWith,
        text: "G"
    };

    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.good;
    return ctx.sync().then(function () {
        console.log("Added new conditional format for strings beginning with 'G' - formatted 'good'");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### DatesOccurring:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.textContains;
    conditionalFormat.customOrNull.rule.textContains = {
        type: Excel.StringMatchType.beginningWith,
        text: "G"
    };

    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.good;
    return ctx.sync().then(function () {
        console.log("Added new conditional format for strings beginning with 'G' - formatted 'good'");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Blanks:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.blank;
    conditionalFormat.customOrNull.rule.blank = true;
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.bad;
    return ctx.sync().then(function () {
        console.log("Added new conditional format to format blank cells");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Errors:
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.error;
    conditionalFormat.customOrNull.rule.error = true;
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.bad;
    return ctx.sync().then(function () {
        console.log("Added new conditional format to format cells with errors");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### Unique: Formatting Duplicates
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.custom);
    conditionalFormat.customOrNull.rule.type = Excel.ConditionalFormatCustomRuleType.unique;
    conditionalFormat.customOrNull.rule.unique = false;
    conditionalFormat.customOrNull.format.style = Excel.PresetStyle.bad;
    return ctx.sync().then(function () {
        console.log("Added new conditional format to format duplicates");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```


### Add Preset Databar
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    conditionalFormat.dataBarOrNull.positiveFormat = {
        color: "green",
        gradient: false,
    }; // does this make sense/easy to use???
    // perhaps: conditionalFormat.addDataBar("green", false);

    return ctx.sync().then(function () {
        console.log("Added green data bar, without gradient");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### Add Custom Databar
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var conditionalFormat = range.conditionalFormats.add(Excel.ConditionalFormatType.dataBar);
    conditionalFormat.dataBarOrNull.showDataBarOnly = false;
    conditionalFormat.dataBarOrNull.Direction = "LeftToRight";
    conditionalFormat.dataBarOrNull.lowerBoundRule = { type: "percent", formula: "10" };
    conditionalFormat.dataBarOrNull.upperBoundRule = { type: "percent", formula: "90" };
    conditionalFormat.dataBarOrNull.positiveFormat = { color: "blue", gradient: false };

    return ctx.sync().then(function () {
        console.log("Added new blue 10%-90% databar format");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
    });
```

### Delete()
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);
    cf.iconOrNull.style = Excel.IconSet.threeArrows;
    cf.delete();

    return ctx.sync().then(function (ctx) {
        console.log("Removed conditional format from worksheet.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### DeleteFromCurrentRange()
```js
Excel.run(function (ctx) {
    var sheetName = "Sheet1";
    var rangeAddress = "A1:C3";
    var range = ctx.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
    var cf = range.conditionalFormats.add(Excel.ConditionalFormatType.iconSet);

    /// remove conditional format from just A2
    var subsetRangeAddress = "A2";
    var subsetRange = ctx.workbook.worksheets.getItem(sheetName).getRange(subsetRangeAddress);
    var subsetCF = range.conditionalFormats.getItemAt(0);
    subsetCF.deleteFromCurrentRange();
    return ctx.sync().then(function () {
        console.log("Removed conditional format just from A2.");
    });
}).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

### GetRangeOrNull()
