# Outline Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents a collection of PageContents.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID. Read-only.|



## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|paragraphs|[Paragraphs](paragraphs.md)|Gets the child Paragraphs. Read-only.|
|parent|[PageContent](pagecontent.md)|Gets the parent PageContent. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[append(location: string, html: string)](#appendlocation-string-html-string)|void|Appends specified html to this Outline.|
|[getHtml()](#gethtml)|string|Gets the html representation of the Outline.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|
|[replace(html: string)](#replacehtml-string)|void|Replace this Outline with specified html.|

## Method Details


### append(location: string, html: string)
Appends specified html to this Outline.

#### Syntax
```js
outlineObject.append(location, html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|location|string|todo.  Possible values are: Begin, End|
|html|string|todo.|

#### Returns
void

### getHtml()
Gets the html representation of the Outline.

#### Syntax
```js
outlineObject.getHtml();
```

#### Parameters
None

#### Returns
string

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

### replace(html: string)
Replace this Outline with specified html.

#### Syntax
```js
outlineObject.replace(html);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|html|string|todo.|

#### Returns
void
