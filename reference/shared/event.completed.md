

# event.completed
The callback that the add-in invokes to let Office know that the operation is done.

****

|||
|:-----|:-----|
|**Hosts:** Excel, Outlook, PowerPoint, Word|**Add-in type:** Content, task pane, Outlook|
|**Available in [requirement sets](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Mailbox, (Also available in the Shared Office JavaScript API outside of any requirement set.)|
|**Last changed in Mailbox**|1.3|
|**Applicable Outlook modes**|Read and Compose|



```js
event.completed();
```


## Parameters

None

## Support details

A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**|Y|Y|Y|
|**Outlook**|See **Outlook support deails** below. |||
|**PowerPoint**|Y|Y|Y|
|**Word**|Y|Y|Y|

|||
|:-----|:-----|
|**Available in requirement sets**|Mailbox, (Also available in the Shared Office JavaScript API outside of any requirement set.)|
|**Minimum permission level**|[Restricted](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|



## Outlook support details

A capital Y in the following table indicates that this property is supported in the corresponding Outlook host application. An empty cell indicates that the Outlook host application doesn't support this property.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).

 **Important:** Add-in commands and the APIs associated with them currently work only in Outlook in [Office 2016 Preview](https://products.office.com/office-2016-preview) on Windows desktop.

**Supported hosts, by platform**

| |**Office for Windows desktop**|**Office Online(in browser)**|**OWA for Devices**|
|:-----|:-----|:-----|:-----|
|**Outlook**|Y|||

|||
|:-----|:-----|
|**Available in requirement sets**|Mailbox|
|**Minimum permission level**|[ReadWriteItem](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|
|**Add-in types**|Outlook|
|**Last changed in Mailbox**|1.3|
|**Applicable Outlook modes**|Read and Compose|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history


|**Version**|**Changes**|
|:-----|:-----|
|Mailbox 1.3|Introduced|
