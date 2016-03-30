
# Office Add-ins platform overview

Office Add-ins enable you to extend Office clients such as Word, Excel, PowerPoint, and Outlook using web technologies like HTML, CSS, and JavaScript. 

You can use Office Add-ins to: 


-  **Add new functionality to Office clients** - For example, augment Word, Excel, PowerPoint, and Outlook by interacting with Office documents and mail items, bring external data into Office, process Office documents, expose third-party functionality into Office clients, and much more. 
    
-  **Create new rich, interactive objects that can be embedded into Office documents** - For example, maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations; email messages, meeting requests, or appointments.
    
**Office Add-ins run across multiple versions of Office** including Office for Windows Desktop, Office Online, Office for the Mac and Office for the IPad.

>**Note:** For a high-level view of where Office Add-ins are currently supported, see the [Office Add-in host and platform availability](http://dev.office.com/add-in-availability) page. 

To make your add-in available to users, you can publish it on the Office Store or to an on-premises add-in catalog. 

To try out some add-ins, you can install the following add-ins from the Office Store.


|**Office product**|**Add-in**|
|:-----|:-----|
|Excel|[Bing Maps](https://store.office.com/bing-maps-WA102957661.aspx?assetid=WA102957661)|
|Outlook|[Package Tracker](https://store.office.com/package-tracker-WA104162083.aspx?assetid=WA104162083)|
|PowerPoint|[Khan Content from Microsoft](https://store.office.com/khan-content-from-microsoft-WA104320031.aspx?assetid=WA104320031)|
|Word|[Translator](https://store.office.com/translator-WA104124372.aspx?assetid=WA104124372)|


## Anatomy of an Office Add-in


The basic components of an Office Add-in are an XML manifest file and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. When your add-in is ready for your customers, you [publish](../publish/publish.md) it an on-premises add-in catalog or [submit it to the Office Store](http://msdn.microsoft.com/library/ff075782-1303-4517-91cc-b3d730e9b9ae%28Office.15%29.aspx). Your web application needs to be hosted on a web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md).


**Manifest + webpage = an Office Add-in**
![Manifest plus webpage equals Office Add-in](../../images/DK2_AgaveOverview01.png)

###Manifest


The manifest specifies settings and capabilities of the add-in, such as the following:
    
- The add-in's display name, description, ID, version, and default locale.
    
- How the add-in integrates with Office: 
      - For add-ins that extend Word/Excel/PowerPoint/Outlook - The native Extension Points the add-in uses to expose functionality, such as buttons on the ribbon. 
      - For add-ins that create new embedable objects - The URL of the default page that is loaded for the object.
       
    
- The permission level and data access requirements for the add-in.
    
For more information, see [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md).


###Web app

The minimal version of a compliant web app is a static HTML webpage. The page can be hosted any web server, or web hosting service, such as [Microsoft Azure](../publish/host-an-office-add-in-on-microsoft-azure.md). It is responsibility of the Add-in developer to host their web app on their service of choice.  

The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but doesn't interact with either the Office document or any other Internet resource. However, because it is a web application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.net, PHP, or Node.js).Then to interact with Office clients and documents, you can use Office.js, which is a [JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md) that we provide to you. 


**Components of a Hello World Office Add-in**

![Components of a Hello World add-in](../../images/DK2_AgaveOverview07.png)


### Creating an Office Add-in with Visual Studio

The most powerful way to build an Office Add-in is to use the  **Add-in for Office** project template in Visual Studio. Visual Studio creates a complete solution that contains all the files that you need to begin testing your add-in in Office immediately. Visual Studio provides a full range of features to make it easy for you to develop and test Office Add-ins. For more information, see:


- [Create and debug Office Add-ins in Visual Studio](../get-started/create-and-debug-office-add-ins-in-visual-studio.md)

### Creating an Office Add-in with a text editor

If want to use your favorite text editor to create an Office Add-in, see the following article for information about how to get started:

    
- [Create an Office add-in using any editor](../get-started/create-an-office-add-in-using-any-editor.md)
    

### JavaScript API for Office

The JavaScript API for Office contains objects and members for building add-ins and interacting with Office content and web services.

For more information about the JavaScript API for Office, see [Understanding the JavaScript API for Office](../../docs/develop/understanding-the-javascript-api-for-office.md) and the [JavaScript API for Office](../../reference/javascript-api-for-office.md) reference.
    

The Word and Excel JavaScript APIs provide host-specific object models that you can use in an Office Add-in. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for Word or Excel. To learn more about these APIs, see the [Word add-ins](../word/word-add-ins-programming-overview.md) and [Excel add-ins](../excel/excel-add-ins-javascript-programming-overview.md) overview topics.


## Types of Office Add-ins

You can create the following types of Office Add-ins:
 
- Word, Excel, and PowerPoint add-ins
- Excel and PowerPoint content add-ins
- Outlook add-ins

### Word, Excel, and PowerPoint add-ins 
You can **add new functionality** to Office clients using add-ins. Use add-in commands to extend the Office UI. For example, you can add buttons for your add-ins on the ribbon or selected contextual menus, allowing users to easily access their add-ins within Office. Commands buttons can launch the different actions such as showing a pane with a custom HTML or executing a JavaScript function. We recommend that you [watch this Channel9 video](https://channel9.msdn.com/Events/Visual-Studio/Connect-event-2015/316) for a deeper overview of this feature.

**Add-in with commands running in Excel Desktop**
![Add-in commands](../../images/addincommands1.png)

**Add-in with commands running in Excel Online**
![Add-in commands](../../images/addincommands2.png)

You can define your commands in your add-in manifest and the Office platform takes care of interpreting them into native UI. To get started, check out these [samples on GitHub](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/), and see [Add-in commands for Excel, Word and PowerPoint](../design/add-in-commands-for-excel-and-word-preview.md)

>**Note:** Clients that do not support add-in commands yet will run your add-in as a **Taskpane** using the DefaultUrl provided on the manifest. The add-in can then be launched using the "My Add-ins" menu from the Insert Tab.

### Excel and PowerPoint content add-ins 

Content add-ins integrate web-based features as **objects that are embedded inside documents**. Content add-ins let you integrate rich, web-based data visualizations, embedded media (such as a YouTube video player or a picture gallery), as well as other external content.

**Content add-in**

![In content add-in](../../images/DK2_AgaveOverview05.png)

To try out a content add-in in Excel 2013 or Excel Online, install the [Bing Maps](https://store.office.com/bing-maps-WA102957661.aspx?assetid=WA102957661) add-in.

### Outlook Add-ins

Outlook add-ins can extend the Office ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment in a read scenario - the user viewing a received item - or in a compose scenario - the user replying or creating a new item. 

Outlook add-ins can access contextual information from the item, such as address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification on the various supporting host applications, including Outlook, Outlook for Mac, Outlook Web App, and OWA for Devices, to provide a seamless experience on the desktop, web, and tablet and mobile devices.

To learn more, see [Outlook add-ins](../outlook/overview.md)

 >**Note**  Outlook add-ins require a minimum version of Exchange 2013 or Exchange Online to host the user's mailbox. POP and IMAP email accounts aren't supported.

**Outlook add-in with command buttons on the ribbon**

![Add-in Command](../../images/41e46a9c-19ec-4ccc-98e6-a227283623d1.png)

**Outlook add-in contextual**

![Contextual add-in](../../images/DK2_AgaveOverview06.png)

To try out an Outlook add-in in Outlook, Outlook for Mac, or Outlook Web App, install the [Package Tracker](https://store.office.com/package-tracker-WA104162083.aspx?assetid=WA104162083) add-in.



## What can an Office Add-in do?

An Office Add-in can do pretty much anything a webpage can do inside the browser, such as the following:

- Extend Office Native UI by creating custom ribbon buttons and tabs.

- Provide an interactive UI and custom logic through HTML and JavaScript.
    
- Use JavaScript frameworks such as jQuery, Angular, and many others.
    
- Connect to REST endpoints and web services via HTTP and AJAX.
    
- Run server-side code or logic, if the page is implemented using a server-side scripting language such as ASP or PHP.
    

In addition to the regular capabilities of a webpage, Office Add-ins can interact with the Office application and an add-in user's content through a [JavaScript API](../../docs/develop/understanding-the-javascript-api-for-office.md) that the Office Add-ins infrastructure provides. 
    

## Additional resources


- [Design guidelines for Office Add-ins](../../docs/design/add-in-design.md)
    
- [Publish your Office Add-in](../publish/publish.md)
    
