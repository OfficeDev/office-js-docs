# OneNoteOM Object (JavaScript API for OneNote)

_Applies to: OneNote Online_

Represents the top level object which contains any globaly addressable objects such as Notebooks, Current Notebook, Current Section, etc.

## Properties

None

## Relationships
| Relationship | Type	|Description|
|:---------------|:--------|:----------|
|notebooks|[Notebooks](notebooks.md)|Gets the collection of notebooks objects that are open in the OneNote Application instance. Read-only.|

## Methods

| Method		   | Return Type	|Description|
|:---------------|:--------|:----------|
|[getCurrentNotebook()](#getcurrentnotebook)|[Notebook](notebook.md)|Gets the Active Notebook.|
|[getCurrentPage()](#getcurrentpage)|[Page](page.md)|Gets the Active Page.|
|[getCurrentSection()](#getcurrentsection)|[Section](section.md)|Gets the Active Section.|
|[getNotebookById(id: string)](#getnotebookbyidid-string)|[Notebook](notebook.md)|Gets Notebook with matching ID.|
|[getPageById(id: string)](#getpagebyidid-string)|[Page](page.md)|Gets Page with matching ID.|
|[getSectionById(id: string)](#getsectionbyidid-string)|[Section](section.md)|Gets Section with matching ID.|
|[getSectionGroupById(id: string)](#getsectiongroupbyidid-string)|[SectionGroup](sectiongroup.md)|Gets SectionGroup with matching ID.|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|

## Method Details


### getCurrentNotebook()
Gets the Active Notebook.

#### Syntax
```js
oneNoteOMObject.getCurrentNotebook();
```

#### Parameters
None

#### Returns
[Notebook](notebook.md)

### getCurrentPage()
Gets the Active Page.

#### Syntax
```js
oneNoteOMObject.getCurrentPage();
```

#### Parameters
None

#### Returns
[Page](page.md)

### getCurrentSection()
Gets the Active Section.

#### Syntax
```js
oneNoteOMObject.getCurrentSection();
```

#### Parameters
None

#### Returns
[Section](section.md)

### getNotebookById(id: string)
Gets Notebook with matching ID.

#### Syntax
```js
oneNoteOMObject.getNotebookById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|ID of the Notebook.|

#### Returns
[Notebook](notebook.md)

### getPageById(id: string)
Gets Page with matching ID.

#### Syntax
```js
oneNoteOMObject.getPageById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|ID of the Page.|

#### Returns
[Page](page.md)

### getSectionById(id: string)
Gets Section with matching ID.

#### Syntax
```js
oneNoteOMObject.getSectionById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|ID of the Section.|

#### Returns
[Section](section.md)

### getSectionGroupById(id: string)
Gets SectionGroup with matching ID.

#### Syntax
```js
oneNoteOMObject.getSectionGroupById(id);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|id|string|ID of the SectionGroup.|

#### Returns
[SectionGroup](sectiongroup.md)

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
