# CustomXmlPart Object (JavaScript API for Excel)

Represents a custom XML part object in a workbook.

## Properties

| Property	   | Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|id|string|The custom XML part's ID. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|namespaceUri|string|The custom XML part's namespace URI. Read-only.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|

_See property access [examples.](#property-access-examples)_

## Relationships
None


## Methods

| Method		   | Return Type	|Description| Req. Set|
|:---------------|:--------|:----------|:----|
|[delete()](#delete)|void|Deletes the custom XML part.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[deleteAttribute(xpath: string, namespaceMappings: object, name: string)](#deleteattributexpath-string-namespacemappings-object-name-string)|void|Deletes an attribute with the given name from the element identified by xpath.|[ApiSetAttribute.Spec](../requirement-sets/excel-api-requirement-sets.md)|
|[deleteElement(xpath: string, namespaceMappings: object)](#deleteelementxpath-string-namespacemappings-object)|void|Deletes the element identified by xpath.|[ApiSetAttribute.Spec](../requirement-sets/excel-api-requirement-sets.md)|
|[getXml()](#getxml)|string|Gets the custom XML part's full XML content.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[insertAttribute(xpath: string, namespaceMappings: object, name: string, value: string)](#insertattributexpath-string-namespacemappings-object-name-string-value-string)|void|Inserts an attribute with the given name and value to the element identified by xpath.|[ApiSetAttribute.Spec](../requirement-sets/excel-api-requirement-sets.md)|
|[insertElement(xpath: string, xml: string, namespaceMappings: object, index: number)](#insertelementxpath-string-xml-string-namespacemappings-object-index-number)|void|Inserts the given XML under the parent element identified by xpath at child position index.|[ApiSetAttribute.Spec](../requirement-sets/excel-api-requirement-sets.md)|
|[load(param: object)](#loadparam-object)|void|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[1.1](../requirement-sets/excel-api-requirement-sets.md)|
|[query(xpath: string, namespaceMappings: object)](#queryxpath-string-namespacemappings-object)|[string[]](string[].md)|Queries the XML content.|[ApiSetAttribute.Spec](../requirement-sets/excel-api-requirement-sets.md)|
|[setXml(xml: string)](#setxmlxml-string)|void|Sets the custom XML part's full XML content.|[1.4](../requirement-sets/excel-api-requirement-sets.md)|
|[updateAttribute(xpath: string, namespaceMappings: object, name: string, value: string)](#updateattributexpath-string-namespacemappings-object-name-string-value-string)|void|Updates the value of an attribute with the given name of the element identified by xpath.|[ApiSetAttribute.Spec](../requirement-sets/excel-api-requirement-sets.md)|
|[updateElement(xpath: string, xml: string, namespaceMappings: object)](#updateelementxpath-string-xml-string-namespacemappings-object)|void|Updates the XML of the element identified by xpath.|[ApiSetAttribute.Spec](../requirement-sets/excel-api-requirement-sets.md)|

## Method Details


### delete()
Deletes the custom XML part.

#### Syntax
```js
customXmlPartObject.delete();
```

#### Parameters
None

#### Returns
void

### deleteAttribute(xpath: string, namespaceMappings: object, name: string)
Deletes an attribute with the given name from the element identified by xpath.

#### Syntax
```js
customXmlPartObject.deleteAttribute(xpath, namespaceMappings, name);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|xpath|string|Absolute path to the element in XPath notation.|
|namespaceMappings|object|An object whose properties represent namespace aliases and the values are the actual namespace URIs.|
|name|string|Name of the attribute.|

#### Returns
void

### deleteElement(xpath: string, namespaceMappings: object)
Deletes the element identified by xpath.

#### Syntax
```js
customXmlPartObject.deleteElement(xpath, namespaceMappings);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|xpath|string|Absolute path to the element in XPath notation.|
|namespaceMappings|object|An object whose properties represent namespace aliases and the values are the actual namespace URIs.|

#### Returns
void

### getXml()
Gets the custom XML part's full XML content.

#### Syntax
```js
customXmlPartObject.getXml();
```

#### Parameters
None

#### Returns
string

### insertAttribute(xpath: string, namespaceMappings: object, name: string, value: string)
Inserts an attribute with the given name and value to the element identified by xpath.

#### Syntax
```js
customXmlPartObject.insertAttribute(xpath, namespaceMappings, name, value);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|xpath|string|Absolute path to the element in XPath notation.|
|namespaceMappings|object|An object whose properties represent namespace aliases and the values are the actual namespace URIs.|
|name|string|Name of the attribute.|
|value|string|Value of the attribute.|

#### Returns
void

### insertElement(xpath: string, xml: string, namespaceMappings: object, index: number)
Inserts the given XML under the parent element identified by xpath at child position index.

#### Syntax
```js
customXmlPartObject.insertElement(xpath, xml, namespaceMappings, index);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|xpath|string|Absolute path to the parent element in XPath notation.|
|xml|string|XML content to be inserted.|
|namespaceMappings|object|An object whose properties represent namespace aliases and the values are the actual namespace URIs.|
|index|number|Zero-based position at which the new XML to be inserted. If omitted, the XML will be appended as the last child of this parent.|

#### Returns
void

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

### query(xpath: string, namespaceMappings: object)
Queries the XML content.

#### Syntax
```js
customXmlPartObject.query(xpath, namespaceMappings);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|xpath|string|An XPath query.|
|namespaceMappings|object|An object whose properties represent namespace aliases and the values are the actual namespace URIs.|

#### Returns
[string[]](string[].md)

### setXml(xml: string)
Sets the custom XML part's full XML content.

#### Syntax
```js
customXmlPartObject.setXml(xml);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|xml|string|XML content to set.|

#### Returns
void

### updateAttribute(xpath: string, namespaceMappings: object, name: string, value: string)
Updates the value of an attribute with the given name of the element identified by xpath.

#### Syntax
```js
customXmlPartObject.updateAttribute(xpath, namespaceMappings, name, value);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|xpath|string|Absolute path to the element in XPath notation.|
|namespaceMappings|object|An object whose properties represent namespace aliases and the values are the actual namespace URIs.|
|name|string|Name of the attribute.|
|value|string|New value of the attribute.|

#### Returns
void

### updateElement(xpath: string, xml: string, namespaceMappings: object)
Updates the XML of the element identified by xpath.

#### Syntax
```js
customXmlPartObject.updateElement(xpath, xml, namespaceMappings);
```

#### Parameters
| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|:---|
|xpath|string|Absolute path to the element in XPath notation.|
|xml|string|New XML content to be stored.|
|namespaceMappings|object|An object whose properties represent namespace aliases and the values are the actual namespace URIs.|

#### Returns
void
