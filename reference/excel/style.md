# Style Object (JavaScript API for Excel)

An object encapsulating a style's format and other properties.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|addIndent|bool|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|autoIndent|bool|Indicates if text is automatically indented when the text alignment in a cell is set to equal distribution.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|builtIn|bool|Indicates if the style is a built-in style. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|formulaHidden|bool|Indicates if the formula will be hidden when the worksheet is protected.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|horizontalAlignment|string|Represents the horizontal alignment for the style. Possible values are: General, Left, Center, Right, Fill, Justify, CenterAcrossSelection, Distributed.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|includeAlignment|bool|Indicates if the style includes the AutoIndent, HorizontalAlignment, VerticalAlignment, WrapText, IndentLevel, and TextOrientation properties.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|includeBorder|bool|Indicates if the style includes the Color, ColorIndex, LineStyle, and Weight border properties.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|includeFont|bool|Indicates if the style includes the Background, Bold, Color, ColorIndex, FontStyle, Italic, Name, Size, Strikethrough, Subscript, Superscript, and Underline font properties.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|includeNumber|bool|Indicates if the style includes the NumberFormat property.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|includePatterns|bool|Indicates if the style includes the Color, ColorIndex, InvertIfNegative, Pattern, PatternColor, and PatternColorIndex interior properties.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|includeProtection|bool|Indicates if the style includes the FormulaHidden and Locked protection properties.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|indentLevel|int|An integer from 0 to 250 that indicates the indent level for the style.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|locked|bool|Indicates if the object is locked when the worksheet is protected.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|The name of the style. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormat|string|The format code of the number format for the style.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|numberFormatLocal|string|The localized format code of the number format for the style.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|orientation|int|The text orientation for the style.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|readingOrder|string|The reading order for the style. Possible values are: Context, LeftToRight, RightToLeft.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|shrinkToFit|bool|Indicates if text automatically shrinks to fit in the available column width.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|textOrientation|int|The text orientation for the style.|[1.8](../requirement-sets/excel-api-requirement-sets.md)|
|verticalAlignment|string|Represents the vertical alignment for the style. Possible values are: Top, Center, Bottom, Justify, Distributed.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|wrapText|bool|Indicates if Microsoft Excel wraps the text in the object.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|borders|[RangeBorderCollection](rangebordercollection.md)|A Border collection of four Border objects that represent the style of the four borders. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|fill|[RangeFill](rangefill.md)|The Fill of the style. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|
|font|[RangeFont](rangefont.md)|A Font object that represents the font of the style. Read-only.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes this style.|[1.7](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes this style.

#### Syntax
```js
styleObject.delete();
```

#### Parameters
None

#### Returns
void
