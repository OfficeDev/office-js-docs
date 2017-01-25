# Office JavaScript APIs
Welcome to the Office JavaScript API documentation. Here you can find everything you need to know about Office add-ins.  For the best experience, we recommend you view this content on [dev.office.com](https://dev.office.com/docs/add-ins/overview/office-add-ins).

## Review upcoming feature: Document Properties

This branch is dedicated to providing an overview of the upcoming **document properties** feature of Office.Js. 

### Document properties

_Note: The listed features are still under the design and review phase and are not yet available as part of Office.js. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the regular docs._

* This feature is intended to give the ability for Office JavaScript add-ins and applications to access (get, set, delete) document level properties. 
* This will allow add-ins/apps to integrate with any workflow that depends on the document properties.
* Planned applications where this API will be available: Excel, Word, PowerPoint. 
* Stored as part of the document. 
* Avialable to all add-ins that run within the document. 
* Types of document properties planned for support are:

	** **Built-in/Standard properties**: By default, Microsoft Office documents are associated with a set of standard properties, such as author, title, and subject. You can specify your own text values for these properties to make it easier to organize and identify your documents. For example, in Word, you can use the Keywords property to add the keyword customers to your sales files. You can search for all sales files with that keyword. **Automatically updated properties** - These properties include both file system properties (for example, file size or the dates when a file was created or last changed) and statistics that are maintained for you by Office programs (for example, the number of words or characters in a document). You cannot specify or change the automatically updated properties. You can use the automatically updated properties to identify or find documents. For example, you can search for all files created after August 3, 2005, or for all files that were last changed yesterday. 

	** **Custom properties**: You can define additional custom properties for your Office documents. You can assign a text, time, or numeric value to custom properties, and you can also assign them the values yes or no. Custom properties may include - library or organization customized properties as well.


### Target features

Key features that we are targetting as part of the API set: 

* Addins should be able to get/read document property and also read through collection of properties (all kinds).
* Able to set writable built-in properties.
* Support all the property data-types natively supported in the application.    
* Set multiple properties in one call (lower priority)  
* Able to delete a custom property 
* Support above requirements for both add-in and REST (in Excel where REST API is supported).

### Design 

_Note: We are narrowing down the design to this particular option, which was listed in the original publication as the 'alternate design'._ 
 
Provide a dynamic collection object that can host both built-in and custom property types. 
Library/corporate properties are not included in this model separately. 

#### Add properties collection object to document object. 

| Relationship	   | Type	|Description
|:---------------|:--------|:----------|
|properties|property collection|Properties object.|

#### Property collection object 

A collection of properties.

##### Members

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|author|`string`|Author of the document.|
|created|`Date`|Date created. Read-only.|
|..other built-in properties..|..|..|
|.other built-in properties..|..|..|
|..other built-in properties..|..|..|
|custom|CustomProperty collection|A collection of custom property objects|

#### `customProperty` Object

Represents an individual custom property. 

##### Members

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|key|`string`|Key value of the property. `key` value is case-sensitive. |
|value|_varies_|Value of the property. Note: some of the values may be read-only. |
|type|`string`|Type of the value. Could be one of the following: `string` or `number` or `date` or `boolean`. Read-only.|

##### Get or set built-in property

Get and set operation on the built-in property would work just like any other object property's get/set. 

###### Built-in property get

Just load the properties that you are interested in reading on the properties collection. 

```js
Excel.run(function (ctx) { 
    var properties = ctx.workbook.properties;
    properties.load('author');
    return ctx.sync().then(function() {
        console.log(properties.author);
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

###### Built-in property set

Just load the properties that you are interested in reading on the properties collection. 

```js
Excel.run(function (ctx) { 
    var properties = ctx.workbook.properties;
    properties.author = "John Doe";
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```
###### Read collection of custom properties

Read collection of custom properties.

```js
Excel.run(function (ctx) { 
    var c_properties = ctx.workbook.properties.custom;
    c_properties.load('items');
    return ctx.sync().then(function() {
        console.log("Properties Count: " + properties.count);
        for (var i = 0; i < c_properties.items.length; i++)
        {
            console.log(properties.items[i].key + " : " + properties.items[i].value);
        }
    });
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

##### getItem(key: string)

Get a custom property. 

###### Syntax
```js
customPropertiesCollection.getItem(key, value);
```

###### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|`string`|Key of the property to be set.|

###### Returns

Property object. 


###### Example

Use the `getItem` method available on the `custom` collection to get a custom property.   

```js
Excel.run(function (ctx) { 
	var property = ctx.workbook.properties.custom.getItem("review-date");
    return ctx.sync() .then(function() {
        console.log("value: " + property.value);
    }); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

##### set(key: string, value: _variable_)

Set/create a custom property. 

###### Syntax
```js
customPropertiesCollection.set(key);

```

###### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|`string`|Key of the property to be set.|
|value|_variable_|Value of the property to be set (created or updated). Type can vary: `string` or `number` or `date` or `boolean`|

Note: The datatype is implied based on the JavaScript type of the value and properly set on the `property` object.

###### Returns

Property object just set or created. 

###### Example

Use the `set` method available on the `custom` collection.   

```js
Excel.run(function (ctx) { 
    var c_properties = ctx.workbook.properties.custom;
    var today = new Date();
	c_properties.set("review-date", today) 
    return ctx.sync(); 
}).catch(function(error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
});
```

#### delete()

Deletes the custom property object.

###### Syntax
```js
propertyObject.delete();
```

###### Parameters
None

###### Returns
void

###### Examples

```js
Excel.run(function (ctx) { 
	var c_property = ctx.workbook.properties.custom("review-date");	
	c_property.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

**[Tell us what you think](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-Doc-Properties)**


## Give us your feedback

Your feedback is important to us. 
* To let us know about any questions or issues you find in the docs, [submit an issue](https://github.com/OfficeDev/office-js-docs/issues) in this repository. Make sure you state the version+build number of the client you are using, and provide repro steps, console output, and error messages, as appropriate. 
* We also encourage you to fork, make the fix, and do a pull request of your proposed changes. For details, see [Contribute to this documentation](Contributing.md). 
* To let us know about your programming experience, what you would like to see in future versions, code samples, and so on, enter your suggestions and ideas at [Office Developer Platform UserVoice](https://officespdev.uservoice.com/).

## Copyright

Copyright (c) 2016 Microsoft Corporation. All rights reserved.
