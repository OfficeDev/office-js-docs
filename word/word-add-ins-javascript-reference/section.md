# Section Object (JavaScript API for Word)

_Applies to: Word 2016, Word for iPad, Word for Mac 2016, Office 2016_

Represents a section in a Word document.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|body|[Body](body.md)|Gets the body of the section. This does not include the headerfooter and other section metadata. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getFooter(type: HeaderFooterType)](#getfootertype-headerfootertype)|[Body](body.md)|Gets one of the section's footers.|
|[getHeader(type: HeaderFooterType)](#getheadertype-headerfootertype)|[Body](body.md)|Gets one of the section's headers.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getFooter(type: HeaderFooterType)
Gets one of the section's footers.

#### Syntax
```js
sectionObject.getFooter(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|HeaderFooterType|Required. The type of footer to return. This value can be: 'primary', 'firstPage' or 'evenPages'.|

#### Returns
[Body](body.md)

### getHeader(type: HeaderFooterType)
Gets one of the section's headers.

#### Syntax
```js
sectionObject.getHeader(type);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|type|HeaderFooterType|Required. The type of header to return. This value can be: 'primary', 'firstPage' or 'evenPages'.|

#### Returns
[Body](body.md)

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
