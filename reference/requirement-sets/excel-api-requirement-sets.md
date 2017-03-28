# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Excel add-ins run across multiple versions of Office, including Office 2016 for Windows, Office for iPad, Office for Mac, and Office Online. The following table lists the Excel requirement sets, the Office host applications that support that requirement set, and the build versions or number for those applications. 

> **Note**: Any API that is listed as **beta** is not ready for production usage. They are made available so that developers can try them out in test and development environments. They are not meant to be used against production/business critical documents. 

> For the requirement sets that are marked as *Beta*, use the specified (or later) version of the Office software and use the Beta library of the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entires not listed as *Beta* are generally available and you can continue to use Production CDN library: https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  Requirement set  |  Office 2016 for Windows*  |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| ExcelApi 1.6 **Beta**  | Version 1702 (Build TBD) or later| Coming soon |  Coming soon| March 2017 | Coming soon|
| ExcelApi 1.5 **Beta**  | Version 1702 (Build TBD) or later| Coming soon |  Coming soon| March 2017 | Coming soon|
| ExcelApi 1.4 | Version 1701 (Build 7870.2024) or later| Coming soon |  Coming soon| January 2017 | Coming soon|
| ExcelApi 1.3  | Version 1608 (Build 7369.2055) or later| 1.27 or later |  15.27 or later| September 2016 | Version 1608 (Build 7601.6800) or later|
| ExcelApi 1.2  | Version 1601 (Build 6741.2088) or later | 1.21 or later | 15.22 or later| January 2016 ||
| ExcelApi 1.1  | Version 1509 (Build 4266.1001) or later | 1.19 or later | 15.20 or later| January 2016 ||

> **Note**: The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1 requirement set.

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [What version of Office am I using?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Where you can find the version and build number for an Office 365 client application](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Office Online Server overview](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## Runtime requirement support check

During the runtime, add-ins can check if a particular host supports an API requirement set by doing the following-check: 

```js
if (Office.context.requirements.isSetSupported('ExcelApi', 1.3) === true) {
  /// perform actions
}
else {
  /// provide alternate flow/logic
}
```

## Manifest based requirement support check

Use the Requirements element in the add-in manifest to specify critical requirement sets or API members that your add-in must use. If the Office host or platform doesn't support the requirement sets or API members specified in the Requirements element, the add-in won't run in that host or platform, and won't display in My Add-ins. Instead, we recommend that you make your add-in available on all platforms of an Office host, such as Excel for Windows, Excel Online, and Excel for iPad. To make your add-in available on all Office hosts and platforms, use runtime checks instead of the Requirements element.

The following code example shows an add-in that loads in all Office host applications that support ExcelApi requirement set, version 1.3.

```xml
<Requirements>
   <Sets DefaultMinVersion="1.3">
      <Set Name="ExcelApi" MinVersion="1.3"/>
   </Sets>
</Requirements>
```

## Office common API requirement sets
For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).
