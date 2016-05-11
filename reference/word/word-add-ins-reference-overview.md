# Word JavaScript API reference overview

Word provides a rich set of APIs for your application to interact with document content and metadata. Use these APIs to create a compelling experience that integrate with and extend Word. Import and export content, assemble new documents from different data sources, and integrate with document workflows to create custom document solutions.

## Shared versus Word-specific API

There are two JavaScript APIs that you can use to interact with the objects and metadata in a Word document. The first APIs were introduced with Office 2013. These APIs are called the shared OfficeJS APIs as many of the objects can be used in add-ins hosted by two or more Office clients. Use the [Word](http://dev.office.com/reference/add-ins/javascript-api-for-office?product=word) filter on [dev.office.com](dev.office.com) to get a view of the shared APIs that can be used by Word.

Starting with Word 2016 for both Mac and Windows, there is a new strongly-typed JavaScript object model for creating Word add-ins. This new object model gives you access to familiar objects like [body](../../reference/word/body.md), [content controls](../../reference/word/contentcontrol.md), [inline pictures](../../reference/word/inlinepicture.md), [paragraphs](../../reference/word/paragraph.md), and more. This is the preferred API as it provides Word specific objects.

We recommend that you start with the Word-specific APIs as the new object model is easier to use. Use the shared API when you need to:
* Target Word 2013 or Word Online.
* Perform initial actions for the application.
* Check the supported requirement set.
* Access metadata, settings, and environmental information for the document.
* Bind to sections in a document and capture events.
* Use custom XML parts.
* Open a dialog box.

## Manifest

The manifest declares where the add-in is hosted, how it is displayed, which permissions it needs, add-in commands, and other information. Learn more about how you can customize the [add-in manifests](../overview/add-in-manifests.md).

You have a few options for publishing Word add-in manifests. Read about how you can [publish your Office add-in](../publish/publish.md) to a network share, to an app catalog, or to the Office store.

## Get the API

You can reference OfficeJS from the following locations:
* https://appsforoffice.microsoft.com/lib/1/hosted/office.js
* https://appsforoffice.microsoft.com/lib/beta/hosted/office.js

Visual Studio has template projects that will provide OfficeJS. You can also use [nuget to get OfficeJS](https://www.nuget.org/packages/Microsoft.Office.js/).

Get the TypeScript definitions: ```typings install office-js --ambient```

## Understand the API

The shared and the Word-specific APIs are quite a bit different. The shared API uses namespaces to organize the objects and methods for interacting with Word.  It uses callbacks to perform the asynchronous calls to Word. The Word-specific API has strongly typed object definitions for interacting with Word. It uses promises to perform the asynchronous calls to Word.

### Running your add-in

Let's take a look at what you'll need when you run your add-in. All add-ins should have an Office.initialize event handler.  Read [Understanding the API](../develop/understanding-the-javascript-api-for-office.md) for more information about add-in initialization.

Starting with Word 2016, your Word add-in executes by passing a function into the Word.run() method. The function passed into the run method must have a context argument. This [context object](../../reference/word/requestcontext.md) is different than the context object you get from the Office object, although it is used for the same purpose which is to interact with the Word runtime environment. The context object provides access to the Word-specific JavaScript object model. Let's take a look at the comments and code of a basic Word add-in:

**Example 1. Initialization and execution of a Word add-in**

```javascript
    (function () {
        "use strict";

        // The initialize event handler is run each time the page is loaded.
        Office.initialize = function (reason) {

            // Checks for the DOM to load using the jQuery ready function.
            $(document).ready(function () {
                // Set your initialization code. You can use the reason
                // argument to determine how the add-in was loaded.
                // You can also load saved settings from the Office object
                // or use the shared OfficeJS object model.
            });
        };

        // Run a batch operation against the Word-specific object model.
        // Use the context argument to get access to the Word document.
        Word.run(function (context) {

            // Create a proxy object for the document.
            var thisDocument = context.document;
        })
    })();
```

The above example shows the basic code needed to create a Word add-in. It initializes the Office object and contains a run method for interacting with the Word document.

### Proxy objects

The Word-specufuc JavaScript object model is loosely coupled with the objects in Word. The Word JavaScript objects are proxy objects for the real objects in a Word document. All actions taken on proxy objects are not realized in Word, and the state of the Word document is not realized in the proxy objects, until the document state has been synchronized. The document state is synchronized when context.sync() is run. The sync() method essentially runs the set of commands in queue for each proxy object and then returns a promise.  Example 2 shows the creation of a proxy body object and a queued command to load the text property on the proxy body object, and then the synchronization of the body in the Word document with the body proxy object.

**Example 2. Synchronization of the document body with the body proxy object.**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        // The body object hasn't been set with any property values.
        var body = context.document.body;

        // Queue a command to load the text property for the proxy document body object.
        context.load(body, 'text');

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });
    })
```

### Command queue

The Word proxy objects have methods for accessing and updating the object model. These methods are executed sequentially in the order in which they were queued in the batch. A batch of commands is formed before a context.sync() call is made. All of the commands queued in all of the objects that use the context will be executed.

In example 3, we demonstrate how the queue of commands works. When context.sync() is called, the first thing that happens is that the [command to load](../../reference/word/loadoption.md) the body text is executed in Word. Then, the command to insert text into the body on Word occurs. The results are then returned to the body proxy object. The value of the body.text property in the Word JavaScript will be the value of the Word document body <u>before</u> the text was inserted into Word document.

**Example 3. Executing a batch of commands.**

```javascript
    // Run a batch operation against the Word object model.
    Word.run(function (context) {

        // Create a proxy object for the document body.
        var body = context.document.body;

        // Queue a command to load the text in the proxy body object.
        context.load(body, 'text');

        // Queue a command to insert text into the end of the Word document body.
        body.insertText('This is text inserted after loading the body.text property',
                        Word.InsertLocation.end);

        // Synchronize the document state by executing the queued commands,
        // and return a promise to indicate task completion.
        return context.sync().then(function () {
            console.log("Body contents: " + body.text);
        });
    })
```

## OpenSpec

Learn about the new Word APIs by reading our Open API specifications. We now share our pre-release, and preview API information so that you can review and provide feedback on our APIs. Read more about [OpenSpec](../../reference/openspec.md) to learn which new features are in the pipeline for the Word JavaScript APIs.

## Give us your feedback

Your feedback is important to us.

* Check out the docs and let us know about any questions and issues you find in them by [submitting an issue](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.
* Let us know about your programming experience, what you would like to see in future versions, code samples, etc. Use [this site](http://officespdev.uservoice.com/) for entering your suggestions and ideas.
* Need help with this API? Post questions to [StackOverflow with the office-js](http://stackoverflow.com/questions/tagged/office-js) tag.

## Additional resources

* [Word add-ins](word-add-ins.md)
* [Office Add-ins Overview](../overview/office-add-ins.md)
* [Word add-ins on GitHub](https://github.com/OfficeDev?utf8=%E2%9C%93&query=Word)
* [Snippet Explorer for Word](http://officesnippetexplorer.azurewebsites.net/#/snippets/word)