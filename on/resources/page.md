# Page Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents the Page in OneNote.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID. Read-only.|
|pageLevel|int|Gets the page level which indicates the number of indentation in the page tab. Read-only.|
|title|string|Gets the title. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|parent|[Section](section.md)|Gets the parent Section. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[addImageFromBase64(xPosition: number, yPosition: number, base64EncodedImage: String)](#addimagefrombase64xposition-number-yposition-number-base64encodedimage-string)|[Image](image.md)|Adds an image in a new PageContent under this Page.|
|[addOutline(xPosition: number, yPosition: number, html: String)](#addoutlinexposition-number-yposition-number-html-string)|[Outline](outline.md)|Adds an outline in a new PageContent under this Page.|
|[changeTitle(title: string)](#changetitletitle-string)|void|Updates the Title.|
|[getContents()](#getcontents)|[PageContentCollection](pagecontentcollection.md)|Gets the collection of child PageContents.|
|[insertPageAsSibling(location: string, title: string)](#insertpageassiblinglocation-string-title-string)|[Page](page.md)|insert a new page as a sibling of this page|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[setPageLevel(level: number)](#setpagelevellevel-number)|void|Sets the page level which indicates the number of indentation in the page tab.|

## Method Details


### addImageFromBase64(xPosition: number, yPosition: number, base64EncodedImage: String)
Adds an image in a new PageContent under this Page.

#### Syntax
```js
pageObject.addImageFromBase64(xPosition, yPosition, base64EncodedImage);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|xPosition|number|Position x of the Outline.|
|yPosition|number|Position y of the Outline.|
|base64EncodedImage|String|Base64 encoded image.|

#### Returns
[Image](image.md)

### addOutline(xPosition: number, yPosition: number, html: String)
Adds an outline in a new PageContent under this Page.

#### Syntax
```js
pageObject.addOutline(xPosition, yPosition, html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|xPosition|number|Position x of the Outline.|
|yPosition|number|Position y of the Outline.|
|html|String|Html string describing the Outline visual presentation.|

#### Returns
[Outline](outline.md)

### changeTitle(title: string)
Updates the Title.

#### Syntax
```js
pageObject.changeTitle(title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|title|string|the new title.|

#### Returns
void

### getContents()
Gets the collection of child PageContents.

#### Syntax
```js
pageObject.getContents();
```

#### Parameters
None

#### Returns
[PageContentCollection](pagecontentcollection.md)

### insertPageAsSibling(location: string, title: string)
insert a new page as a sibling of this page

#### Syntax
```js
pageObject.insertPageAsSibling(location, title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|location|string|location of the added page.  Possible values are: Before, After|
|title|string|title for the added page.|

#### Returns
[Page](page.md)

### load(param: object)
Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.

#### Syntax
```js
object.load(param);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|param|object|Optional. Accepts parameter and relationship names as delimited string or an array. Or, provide [loadOption](loadoption.md) object.|

#### Returns
void

### setPageLevel(level: number)
Sets the page level which indicates the number of indentation in the page tab.

#### Syntax
```js
pageObject.setPageLevel(level);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|level|number|Indentation level.|

#### Returns
void
