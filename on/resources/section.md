# Section Object (JavaScript API for Excel)

_Applies to: Excel 2016, Office 2016_

Represents the Section in OneNote.

## Properties

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|id|string|Gets the ID. Read-only.|
|name|string|Gets the name. Read-only.|

_See property access [examples.](#property-access-examples)_

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|parentNotebook|[Notebook](notebook.md)|Gets the Notebook of this Section. Read-only.|
|parentSectionGroup|[SectionGroup](sectiongroup.md)|Gets the parent SectionGroup if it exists. Otherwise returns null which indicates this is under Notebook. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[addPage(location: string, title: string)](#addpagelocation-string-title-string)|[Page](page.md)|add a new page at begining or end.|
|[getPages()](#getpages)|[PageCollection](pagecollection.md)|Gets the child Pages.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### addPage(location: string, title: string)
add a new page at begining or end.

#### Syntax
```js
sectionObject.addPage(location, title);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|location|string|location of the added page.  Possible values are: Begin, End|
|title|string|title for the added page.|

#### Returns
[Page](page.md)

### getPages()
Gets the child Pages.

#### Syntax
```js
sectionObject.getPages();
```

#### Parameters
None

#### Returns
[PageCollection](pagecollection.md)

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
