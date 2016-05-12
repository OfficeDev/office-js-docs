# Office JavaScript APIs
Welcome to the Office JavaScript API documentation. Here you can find everything you need to about Office add-ins.  To view these documents, we recommend you view them on [dev.office.com](https://dev.office.com/docs/add-ins/overview/office-add-ins).

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
** *Built-in/Standard properties*: By default, Microsoft Office documents are associated with a set of standard properties, such as author, title, and subject. You can specify your own text values for these properties to make it easier to organize and identify your documents. For example, in Word, you can use the Keywords property to add the keyword customers to your sales files. You can search for all sales files with that keyword. *Automatically updated properties* - These properties include both file system properties (for example, file size or the dates when a file was created or last changed) and statistics that are maintained for you by Office programs (for example, the number of words or characters in a document). You cannot specify or change the automatically updated properties. You can use the automatically updated properties to identify or find documents. For example, you can search for all files created after August 3, 2005, or for all files that were last changed yesterday. 
** *Custom properties*: You can define additional custom properties for your Office documents. You can assign a text, time, or numeric value to custom properties, and you can also assign them the values yes or no. Custom properties may include - library or organization customized properties as well.


### Target features

Key features that we are targetting as part of the API set: 

* Addins should be able to get/read document property and also read through collection of properties (all kinds).
* Able to set writable built-in properties.
* Support all the property data-types natively supported in the application.    
* Set multiple properties in one call (lower priority)  
* Able to delete a custom property 
* Support above requirements for both add-in and REST (in Excel where REST API is supported).

### Design 
 
* Properties can be thought of as relationships (property bag) on the document (workbook, document, presentation) object. 
* Given that built-in and custom properties could have the same names, we'll have to provide two separate propertiy bags.  
* Data type is an important aspect of property and hence we need to think of each property as an object with key, value and type and other information.   
* Provide get, set (to both add and update properties) and delete methods. 


#### `property` Object

Represnts an individual property. 

##### Members

| Property	   | Type	|Description
|:---------------|:--------|:----------|
|key|string|Key value of the property. `key` value is case-sensitive. |
|value|_varies_|Value of the property. Note: some of the values may be read-only. |
|datatype|`string`|Datatype of the value. Could be one of the following: `string` or `number` or `date` or `boolean`|

#### Add new relations to document object. 

| Relationship	   | Type	|Description
|:---------------|:--------|:----------|
|builtinProperties|property collection|Collection of built-in/auto `property` objects.|
|customProperties|property collection|Collection of custom `property` objects.|
|libraryProperties|property collection|Collection of library `property` objects. This could be customized as per organizational needs.|
 

#### Methods on the document object 


##### getItem(name: string)

Get an item from the collection

###### Syntax
```js
propertiesCollection.getItem(key);
```

###### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|Key of the property to be retrieved.|

###### Returns

Property object. 

###### Examples

This example shows retrieval of a property for Excel object. This is very similar to how Word example will look.

```js
Excel.run(function (ctx) { 
	var property = ctx.workbook.builtInProperties.getItem("author");
	return ctx.sync().then(function() {
			console.log(property.value);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

Read a custom property.

```js
Excel.run(function (ctx) { 
	var property = ctx.workbook.customProperties.getItem("Objective-Id");
	return ctx.sync().then(function() {
			console.log(property.value);
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
propertiesCollection.set(key);

```

###### Parameters

| Parameter	   | Type	|Description|
|:---------------|:--------|:----------|
|key|string|Key of the property to be retrieved.|
|value|_variable_|Value of the property to be set (created or updated). Datatype can vary: `string` or `number` or `date` or `boolean`|

Note: The datatype is implied based on the JavaScript type of the value and properly set on the `property` object.

###### Returns

Property object just set or created. 

###### Examples

This example shows setting of a built-in property.

```js
Excel.run(function (ctx) { 
	var author = ctx.workbook.builtInProperties.set("author", "John Doe") 
	return ctx.sync().then(function() {
			console.log(author.value);
	});
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```

This example show the setting of document property.

```js
Excel.run(function (ctx) { 
	var today = new Date();
	var reviewDate = ctx.workbook.customProperties.set("review-date", today) 
	return ctx.sync().then(function() {
			console.log(reviewDate.value);
	});
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
	var property = ctx.workbook.customProperties.getItem("review-date");	
	property.delete();
	return ctx.sync(); 
}).catch(function(error) {
		console.log("Error: " + error);
		if (error instanceof OfficeExtension.Error) {
			console.log("Debug info: " + JSON.stringify(error.debugInfo));
		}
});
```




## Give us your feedback

Your feedback is important to us. 
* Check out the docs and let us know about any questions and issues you find in them by [submitting an issue](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository. Make sure you state the version+build number of the client you are using, and provide repro steps, console output, and error messages in any issue you open.  
* Let us know about your programming experience, what you would like to see in future versions, code samples, etc. Use [this site](http://officespdev.uservoice.com/) for entering your suggestions and ideas.
* If you find an issue with the documentation, please create a GitHub issue by clicking the **Issues** tab above. We also encourage you to fork, make the fix, and do a pull request of your proposed changes.
