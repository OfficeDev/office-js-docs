## Introduction to Excel JS APIs 1.3 

_Applies to: Excel 2016, Excel Online, Excel for iOS, Excel for Mac

Below links describes the new set of Excel JavaScript APIs that are being planned for the next release (Requirement Set 1.3). Please review and provide your feedback. One of the best ways of providing your input is by opening new issue in GitHub using the links available below. 

_**Note**: below listed features are still under design and review phase and hence not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._


# Pivot Table
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[pivotTable](/excel/resources/pivottable.md)|_Property_ > name|Name of the pivot table.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=pivotTable-name)|
|[pivotTable](/excel/resources/pivottable.md)|_Method_ > [refresh()](/excel/resources/pivottable.md#refresh)|Refreshes the pivot table.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTable-refresh)|
|[pivotTableCollection](/excel/resources/pivottablecollection.md)|_Property_ > count|Returns the number of pivot tables in the workbook or worksheet. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=pivotTableCollection-count)|
|[pivotTableCollection](/excel/resources/pivottablecollection.md)|_Property_ > items|A collection of pivotTable objects. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=pivotTableCollection-items)|
|[pivotTableCollection](/excel/resources/pivottablecollection.md)|_Method_ > [getItem(key: string)](/excel/resources/pivottablecollection.md#getitemkey-string)|Gets a pivot table by Name.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-getItem)|
|[pivotTableCollection](/excel/resources/pivottablecollection.md)|_Method_ > [getItemAt(index: number)](/excel/resources/pivottablecollection.md#getitematindex-number)|Gets a pivot table based on its position in the collection.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-getItemAt)|
|[pivotTableCollection](/excel/resources/pivottablecollection.md)|_Method_ > [getItemOrNull(key: string)](/excel/resources/pivottablecollection.md#getitemornullkey-string)|Gets a pivot table by Name. If the pivot table does not exist, the return object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-getItemOrNull)|
|[pivotTableCollection](/excel/resources/pivottablecollection.md)|_Method_ > [refreshAll()](/excel/resources/pivottablecollection.md#refreshall)|Refreshes all the pivot tables in the collection.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-refreshAll)|

# Range
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[range](/excel/resources/range.md)|_Relationship_ > visibleView|The visible part of the range. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=range-visibleView)|
|[range](/excel/resources/range.md)|_Method_ > [getImage(height: number, width: number, fittingMode: ImageFittingMode)](/excel/resources/range.md#getimageheight-number-width-number-fittingmode-string)|Renders the range as a base64-encoded image by scaling the range to fit the specified dimensions.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getImage)|
|[range](/excel/resources/range.md)|_Method_ > [getExpandedRange(deltaRows: number, deltaColumns: number)](/excel/resources/range.md#getexpandedrangedeltarows-number-deltacolumns-number)|Gets a Range object similar to the current Range object, but with its bottom right corner expanded by some number of rows and columns.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getExpandedRange)|
|[range](/excel/resources/range.md)|_Method_ > [getColumnsAfter(count: number)](/excel/resources/range.md#getcolumnsaftercount-number)|Gets X many columns to the right of the current Range object.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getColumnsAfter)|
|[range](/excel/resources/range.md)|_Method_ > [getColumnsBefore(count: number)](/excel/resources/range.md#getcolumnsbeforecount-number)|Gets X many columns to the left of the current Range object.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getColumnsBefore)|
|[range](/excel/resources/range.md)|_Method_ > [getRowsAfter(count: number)](/excel/resources/range.md#getrowsaftercount-number)|Gets X many rows to the right of the current Range object.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getRowsAfter)|
|[range](/excel/resources/range.md)|_Method_ > [getRowsBefore(count: number)](/excel/resources/range.md#getrowsbeforecount-number)|Gets X many rows to the left of the current Range object.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getRowsBefore)|
|[range](/excel/resources/range.md)|_Method_ > [getIntersectionOrNull(anotherRange: Range or string)](/excel/resources/range.md#getintersectionornullanotherrange-range-or-string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getIntersectionOrNull)|

# RangeView
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[rangeView](/excel/resources/rangeview.md)|_Property_ > address|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-address)|
|[rangeView](/excel/resources/rangeview.md)|_Property_ > addressLocal|Represents range reference for the specified range in the language of the user. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-addressLocal)|
|[rangeView](/excel/resources/rangeview.md)|_Property_ > cellCount|Number of cells in the range. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-cellCount)|
|[rangeView](/excel/resources/rangeview.md)|_Property_ > formulas|Represents the formula in A1-style notation.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-formulas)|
|[rangeView](/excel/resources/rangeview.md)|_Property_ > formulasLocal|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-formulasLocal)|
|[rangeView](/excel/resources/rangeview.md)|_Property_ > numberFormat|Represents Excel's number format code for the given cell.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-numberFormat)|
|[rangeView](/excel/resources/rangeview.md)|_Property_ > rowCount|Returns the total number of rows in the range. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-rowCount)|
|[rangeView](/excel/resources/rangeview.md)|_Property_ > text|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-text)|
|[rangeView](/excel/resources/rangeview.md)|_Property_ > values|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-values)|
|[rangeView](/excel/resources/rangeview.md)|_Relationship_ > rows|Represents a collection of all the rows in the RangeView. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-rows)|
|[rangeView](/excel/resources/rangeview.md)|_Relationship_ > valueTypes|Represents the type of data of each cell. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-valueTypes)|
|[rangeViewCollection](/excel/resources/rangeviewcollection.md)|_Property_ > count|Returns the number of RangeView in this collection. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeViewCollection-count)|
|[rangeViewCollection](/excel/resources/rangeviewcollection.md)|_Property_ > items|A collection of RangeView objects. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeViewCollection-items)|
|[rangeViewCollection](/excel/resources/rangeviewcollection.md)|_Method_ > [getItemAt(index: number)](/excel/resources/rangeviewcollection.md#getitematindex-number)|Gets a rangeView of a certain row based on its position in the collection.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-rangeViewCollection-getItemAt)|
|[rangeViewCollection](/excel/resources/rangeviewcollection.md)|_Method_ > [load(param: object)](/excel/resources/rangeviewcollection.md#loadparam-object)|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-rangeViewCollection-load)|

# Table
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[table](/excel/resources/table.md)|_Property_ > highlightFirstColumn|Indicates whether the first column contains special formatting.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-highlightFirstColumn)|
|[table](/excel/resources/table.md)|_Property_ > highlightLastColumn|Indicates whether the last column contains special formatting.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-highlightLastColumn)|
|[table](/excel/resources/table.md)|_Property_ > showBandedColumns|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-showBandedColumns)|
|[table](/excel/resources/table.md)|_Property_ > showBandedRows|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-showBandedRows)|
|[table](/excel/resources/table.md)|_Property_ > showFilterButton|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=table-showFilterButton)|
|[table](/excel/resources/table.md)|_Method_ > [getImage(height: number, width: number, fittingMode: ImageFittingMode)](/excel/resources/table.md#getimageheight-number-width-number-fittingmode-imagefittingmode)|Renders the table as a base64-encoded image by scaling the table to fit the specified dimensions.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-table-getImage)|
|[tableCollection](/excel/resources/tablecollection.md)|_Method_ > [getItemOrNull(key: number or string)](/excel/resources/tablecollection.md#getitemornullkey-number-or-string)|Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-tableCollection-getItemOrNull)|
|[tableColumnCollection](/excel/resources/tablecolumncollection.md)|_Method_ > [getItemOrNull(key: number or string)](/excel/resources/tablecolumncollection.md#getitemornullkey-number-or-string)|Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-tableColumnCollection-getItemOrNull)|
|[tableRowCollection](/excel/resources/tablerowcollection.md)|_Method_ > [insertRows(index: number, values: (boolean or string or number)[][])](/excel/resources/tablerowcollection.md#insertrowsindex-number-values-boolean-or-string-or-number)|Inserts multiple new rows to the table. Return the first row that is added.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-tableRowCollection-addMultipleRows)|

# Workbook
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[workbook](/excel/resources/workbook.md)|_Relationship_ > pivotTables|Returns sheet protection object for a worksheet. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=workbook-pivotTables)|

# Worksheet
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[worksheet](/excel/resources/worksheet.md)|_Relationship_ > pivotTables|Returns sheet protection object for a worksheet. Read-only.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=worksheet-pivotTables)|
|[worksheet](/excel/resources/worksheet.md)|_Relationship_ > visibility|The Visibility of the worksheet.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=worksheet-visibility)|
|[worksheetCollection](/excel/resources/worksheetcollection.md)|_Method_ > [getItemOrNull(key: string)](/excel/resources/worksheetcollection.md#getitemornullkey-string)|Gets a worksheet object using its Name or ID. If the worksheet does not exist, the returned object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-worksheetCollection-getItemOrNull)|

# Binding
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[bindingCollection](/excel/resources/bindingcollection.md)|_Method_ > [add(range: Range or string, bindingType: string, id: string)](/excel/resources/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|Add a new binding to a particular Range.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-add)|
|[bindingCollection](/excel/resources/bindingcollection.md)|_Method_ > [addFromNamedItem(name: string, bindingType: string, id: string)](/excel/resources/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|Add a new binding based on a named item in the workbook.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-addFromNamedItem)|
|[bindingCollection](/excel/resources/bindingcollection.md)|_Method_ > [addFromSelection(bindingType: string, id: string)](/excel/resources/bindingcollection.md#addfromselectionbindingtype-string-id-string)|Add a new binding based on the current selection.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-addFromSelection)|
|[bindingCollection](/excel/resources/bindingcollection.md)|_Method_ > [getItemOrNull(id: string)](/excel/resources/bindingcollection.md#getitemornullid-string)|Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-getItemOrNull)|

# Chart
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[chartCollection](/excel/resources/chartcollection.md)|_Method_ > [getItemOrNull(name: string)](/excel/resources/chartcollection.md#getitemornullname-string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-chartCollection-getItemOrNull)|

# Named Item
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[namedItemCollection](/excel/resources/nameditemcollection.md)|_Method_ > [getItemOrNull(name: string)](/excel/resources/nameditemcollection.md#getitemornullname-string)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|[Go](https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-namedItemCollection-getItemOrNull)|





## Your feedback is important to us 


* Check out the new features and let us know about any comments, questions and issues by following the feedback links provided all through the docs or by [submitting an issue](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository. 
* Let us know about your programming experience, what you would like to see in future versions, code samples, etc. 


