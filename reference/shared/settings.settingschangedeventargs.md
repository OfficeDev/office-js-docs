# Settings.settingschangedeventargs object
Provides information about the settings that raised the [settingsChanged event](settings.settingschangedevent.md).

|||
|:-----|:-----|
|**Hosts:**|Access, Excel, Word|
|**Available in [Requirement set](../../docs/overview/specify-office-hosts-and-api-requirements.md)**|Settings|
|**Last changed in**|1.0|

```js
Office.EventType.SettingsChanged
```

## Members

**Properties**

|**Name**|**Description**|
|:-----|:-----|
|**[settings](settings.settingschangedeventargs.setting.md)**|Gets a **Settings** object that represents the settings that raised the settingsChanged event.|
|**[type](settings.settingschangedeventargs.type.md)**|Get an **EventType** enumeration value that identifies the kind of event that was raised.|

## Remarks

To add an event handler for the  **settingsChanged** event, use the [addHandlerAsync](settings.addhandlerasync.md) method of the **Settings** object.

The  **settingsChanged** event fires only when your add-in's script calls the **Settings.saveAsync** method to persist the in-memory copy of the settings into the document file. The **settingsChanged** event is not triggered when the [Settings.set](settings.set.md) or [Settings.remove](settings.remove.md) methods are called.

The  **settingsChanged** event was designed to let you to handle potential conflicts when two or more users are attempting to save settings at the same time when your add-in is used in a shared (co-authored) document.


 >**Important**:  Your add-in's code can register a handler for the  **settingsChanged** event when the add-in is running with any Excel client, but the event will fire only when the add-in is loaded with a spreadsheet that is opened in Excel Online, _and_ more than one user is editing the spreadsheet (co-authoring). Therefore, effectively the **settingsChanged** event is supported only in Excel Online in co-authoring scenarios.



## Support details


A capital Y in the following matrix indicates that this object is supported in the corresponding Office host application. An empty cell indicates that the Office host application doesn't support this object.

For more information about Office host application and server requirements, see [Requirements for running Office Add-ins](../../docs/overview/requirements-for-running-office-add-ins.md).


||**Office for Windows desktop**|**Office Online (in browser)**|**Office for iPad**|
|:-----|:-----|:-----|:-----|
|**Excel**||Y||
|**Word**|Y|Y||

|||
|:-----|:-----|
|**Available in requirement sets**|Settings|
|**Minimum permission level**|Restricted|
|**Add-in types**|Content, task pane|
|**Library**|Office.js|
|**Namespace**|Office|

## Support history

|**Version**|**Changes**|
|:-----|:-----|
|1.0|Introduced|
