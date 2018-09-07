# ShapeFont Object (JavaScript API for Excel)

This object represents the font attributes (font name, font size, color, etc.) for a TextRange in the Shape.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|bold|bool|Represents the bold status of font. Returns null the TextRange includes both bold and non-bold text fragments.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|color|string|HTML color code representation of the text color. E.g. #FF0000 represents Red. Returns null if the TextRange includes text fragments with different colors.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|italic|bool|Represents the italic status of font. Return null if the TextRange includes both italic and non-italic text fragments.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|name|string|Represents font name (e.g. "Calibri"). If the text is Complex Script or East Asian language, represents corresponding font name; otherwise represents Latin font name.|[beta](../requirement-sets/excel-api-requirement-sets.md)|
|size|double|Represents font size in points (e.g. 11). Return null if the TextRange includes text fragments with different font sizes.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|underline|string|Type of underline applied to the font. Return null if the TextRange includes text fragments with different underline styles.|[beta](../requirement-sets/excel-api-requirement-sets.md)|

## Methods
None

