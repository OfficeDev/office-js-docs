## Excel JavaScript API Open Specification

_Applies to: Excel 2016, Excel Online, Excel for iOS, Excel for Mac_

This page describes the new set of Excel JavaScript APIs that are planned for future releases. It consists of APIs that are in `beta` stage and others that are still in the design phase. Please review and provide your feedback. One of the best ways you can provide input is to open a new issue in GitHub using the links available below.

_**Note**: The following features are still under design and review phase and not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._

> **Note**: Any API that is listed as **beta** is not ready for production usage. They are made available so that developers can try them out in test and development environments. They are not meant to be used against production/business critical documents. 

> For the requirement sets that are marked as *Beta*, use the specified (or later) version of Office and use the Beta library of the CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js. Entries not listed as *Beta* are generally available and you can continue to use Production CDN library: https://appsforoffice.microsoft.com/lib/1/hosted/office.js

## New APIs in Beta

* __Events__: Support events of Chart Object. For details, see [New Event API features](/reference/new-events.md)

* __Charts__: Get Chart ID and type, more properies, formats and position for Axis, control the styles and behaviors of Chart with advanced properties, control the styles and formats for DataLabel, customerize Ero bars and Legned, set the TrendlineLabel formats and positions, set the plot area of Charts etc.

* __Calculate__: Set the calculation to Auto/Manual temporarily.

## Upcoming APIs to be previewed

* __Events__: Add onChanged and onSelectionChanged under WorksheetCollection object. Add onAdded/onDeleted event under Table object.

* __Calculate__: Get the calculate related info. 

* __AutoFilter__: Create and apply an AutoFilter to a range in worksheet.

* __Print layout and page setup__: Control the print style of worksheet, setup the page size and layout.


## Give us feedback!

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know your thoughts by submitting [issues](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.

For API support, you can post questions to the community on [StackOverflow](http://stackoverflow.com/questions/tagged/office-js) and tag them with [office-js].
