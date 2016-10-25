# List Object (JavaScript API for Word)

_Word 2016, Word for iPad, Word for Mac, Word Online_

Contains a collection of [paragraph](paragraph.md) objects.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|int|Gets the list's id. Read-only.|[1.3](../reqset/word-requirement.md)|
|levelExistences|bool|Checks whether each of the 9 levels exists in the list. A true value indicates the level exists, which means there is at least one list item at that level. Read-only.|[1.3](../reqset/word-requirement.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|levelTypes|[ListLevelType](listleveltype.md)|Gets all 9 level types in the list. Each type can be 'Bullet', 'Number' or 'Picture'. Read-only.|[1.3](../reqset/word-requirement.md)|
|paragraphs|[ParagraphCollection](paragraphcollection.md)|Gets paragraphs in the list. Read-only.|[1.3](../reqset/word-requirement.md)|

## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[getLevelFont(level: number)](#getlevelfontlevel-number)|[Font](font.md)|Gets the font of the bullet, number or picture at the specified level in the list.|[1.4](../reqset/word-requirement.md)|
|[getLevelParagraphs(level: number)](#getlevelparagraphslevel-number)|[ParagraphCollection](paragraphcollection.md)|Gets the paragraphs that occur at the specified level in the list.|[1.3](../reqset/word-requirement.md)|
|[getLevelPicture(level: number)](#getlevelpicturelevel-number)|string|Gets the base64 encoded string representation of the picture at the specified level in the list.|[1.4](../reqset/word-requirement.md)|
|[getLevelString(level: number)](#getlevelstringlevel-number)|string|Gets the bullet, number or picture at the specified level as a string.|[1.3](../reqset/word-requirement.md)|
|[insertParagraph(paragraphText: string, insertLocation: InsertLocation)](#insertparagraphparagraphtext-string-insertlocation-insertlocation)|[Paragraph](paragraph.md)|Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.|[1.3](../reqset/word-requirement.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../reqset/word-requirement.md)|
|[resetLevelFont(level: number, resetFontName: bool)](#resetlevelfontlevel-number-resetfontname-bool)|void|Resets the font of the bullet, number or picture at the specified level in the list.|[1.4](../reqset/word-requirement.md)|
|[setLevelAlignment(level: number, alignment: Alignment)](#setlevelalignmentlevel-number-alignment-alignment)|void|Sets the alignment of the bullet, number or picture at the specified level in the list.|[1.3](../reqset/word-requirement.md)|
|[setLevelBullet(level: number, listBullet: ListBullet, charCode: number, fontName: string)](#setlevelbulletlevel-number-listbullet-listbullet-charcode-number-fontname-string)|void|Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.|[1.3](../reqset/word-requirement.md)|
|[setLevelIndents(level: number, textIndent: float, textIndent: float)](#setlevelindentslevel-number-textindent-float-textindent-float)|void|Sets the two indents of the specified level in the list.|[1.3](../reqset/word-requirement.md)|
|[setLevelNumbering(level: number, listNumbering: ListNumbering, formatString: object[])](#setlevelnumberinglevel-number-listnumbering-listnumbering-formatstring-object)|void|Sets the numbering format at the specified level in the list.|[1.3](../reqset/word-requirement.md)|
|[setLevelPicture(level: number, base64EncodedImage: string)](#setlevelpicturelevel-number-base64encodedimage-string)|void|Sets the picture at the specified level in the list.|[1.4](../reqset/word-requirement.md)|
|[setLevelStartingNumber(level: number, startingNumber: number)](#setlevelstartingnumberlevel-number-startingnumber-number)|void|Sets the starting number at the specified level in the list. Default value is 1.|[1.3](../reqset/word-requirement.md)|

## Method Details


### getLevelFont(level: number)
Gets the font of the bullet, number or picture at the specified level in the list.

#### Syntax
```js
listObject.getLevelFont(level);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|

#### Returns
[Font](font.md)

### getLevelParagraphs(level: number)
Gets the paragraphs that occur at the specified level in the list.

#### Syntax
```js
listObject.getLevelParagraphs(level);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|

#### Returns
[ParagraphCollection](paragraphcollection.md)

### getLevelPicture(level: number)
Gets the base64 encoded string representation of the picture at the specified level in the list.

#### Syntax
```js
listObject.getLevelPicture(level);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|

#### Returns
string

### getLevelString(level: number)
Gets the bullet, number or picture at the specified level as a string.

#### Syntax
```js
listObject.getLevelString(level);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|

#### Returns
string

### insertParagraph(paragraphText: string, insertLocation: InsertLocation)
Inserts a paragraph at the specified location. The insertLocation value can be 'Start', 'End', 'Before' or 'After'.

#### Syntax
```js
listObject.insertParagraph(paragraphText, insertLocation);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|paragraphText|string|Required. The paragraph text to be inserted.|
|insertLocation|InsertLocation|Required. The value can be 'Start', 'End', 'Before' or 'After'.|

#### Returns
[Paragraph](paragraph.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### resetLevelFont(level: number, resetFontName: bool)
Resets the font of the bullet, number or picture at the specified level in the list.

#### Syntax
```js
listObject.resetLevelFont(level, resetFontName);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|
|resetFontName|bool|Optional. Optional. Indicates whether to reset the font name. Default is false that indicates the font name is kept unchanged.|

#### Returns
void

### setLevelAlignment(level: number, alignment: Alignment)
Sets the alignment of the bullet, number or picture at the specified level in the list.

#### Syntax
```js
listObject.setLevelAlignment(level, alignment);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|
|alignment|Alignment|Required. The level alignment that can be 'left', 'centered' or 'right'.|

#### Returns
void

### setLevelBullet(level: number, listBullet: ListBullet, charCode: number, fontName: string)
Sets the bullet format at the specified level in the list. If the bullet is 'Custom', the charCode is required.

#### Syntax
```js
listObject.setLevelBullet(level, listBullet, charCode, fontName);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|
|listBullet|ListBullet|Required. The bullet.|
|charCode|number|Optional. Optional. The bullet character's code value. Used only if the bullet is 'Custom'.|
|fontName|string|Optional. Optional. The bullet's font name. Used only if the bullet is 'Custom'.|

#### Returns
void

### setLevelIndents(level: number, textIndent: float, textIndent: float)
Sets the two indents of the specified level in the list.

#### Syntax
```js
listObject.setLevelIndents(level, textIndent, textIndent);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|
|textIndent|float|Required. The text indent in points. It is the same as paragraph left indent.|
|textIndent|float|Required. The relative indent, in points, of the bullet, number or picture. It is the same as paragraph first line indent.|

#### Returns
void

### setLevelNumbering(level: number, listNumbering: ListNumbering, formatString: object[])
Sets the numbering format at the specified level in the list.

#### Syntax
```js
listObject.setLevelNumbering(level, listNumbering, formatString);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|
|listNumbering|ListNumbering|Required. The ordinal format.|
|formatString|object[]|Optional. Optional. The numbering string format defined as an array of strings and/or integers. Each integer is a level of number type that is higher than or equal to this level. For example, an array of ["(", level - 1, ".", level, ")"] can define the format of "(2.c)", where 2 is the parent's item number and c is this level's item number.|

#### Returns
void

### setLevelPicture(level: number, base64EncodedImage: string)
Sets the picture at the specified level in the list.

#### Syntax
```js
listObject.setLevelPicture(level, base64EncodedImage);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|
|base64EncodedImage|string|Optional. Optional. The base64 encoded image to be set. If not given, the default picture is set.|

#### Returns
void

### setLevelStartingNumber(level: number, startingNumber: number)
Sets the starting number at the specified level in the list. Default value is 1.

#### Syntax
```js
listObject.setLevelStartingNumber(level, startingNumber);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|level|number|Required. The level in the list.|
|startingNumber|number|Required. The number to start with.|

#### Returns
void
