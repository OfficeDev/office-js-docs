
# officeTheme.controlForegroundColor property
Gets the Office theme control foreground color.





|||
|:-----|:-----|
|**Hosts:**|Excel, Outlook, PowerPoint, Word|
|**Minimum permission level**|[Restricted](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)|
|**Available in [Requirement set](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)**|Not in a set|
|**Added in**|1.3|

```js
var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;
```


## Return value

A hex color triplet


## Remarks

The colors returned correspond to the values of the Office theme selected by the user with  ** File** > **Office Account** > **Office Theme** UI, which is applied across all Office host applications.


## Support details


A capital Y in the following matrix indicates that this method is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this method.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](https://docs.microsoft.com/office/dev/add-ins/concepts/requirements-for-running-office-add-ins).


**Supported hosts, by platform**


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|**OWA for Devices**|
|:-----|:-----|:-----|:-----|:-----|
|**Excel**|Y||||
|**Outlook**|Y||||
|**PowerPoint**|Y||||
|**Word**|Y||||



|||
|:-----|:-----|
|**Minimum permission level**|[Restricted](https://docs.microsoft.com/office/dev/add-ins/develop/requesting-permissions-for-api-use-in-content-and-task-pane-add-ins)|
|**Add-in types**|Content, task pane, Outlook|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history



|**Version**|**Changes**|
|:-----|:-----|
|1.3|Introduced|
