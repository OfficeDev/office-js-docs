# Binding
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[bindingCollection](/resource/bindingcollection.md)|_Method_ > [add(range: Range or string, bindingType: string, id: string)](/resource/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|Add a new binding to a particular Range.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-add)|
|[bindingCollection](/resource/bindingcollection.md)|_Method_ > [addFromNamedItem(name: string, bindingType: string, id: string)](/resource/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|Add a new binding based on a named item in the workbook.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-addFromNamedItem)|
|[bindingCollection](/resource/bindingcollection.md)|_Method_ > [addFromSelection(bindingType: string, id: string)](/resource/bindingcollection.md#addfromselectionbindingtype-string-id-string)|Add a new binding based on the current selection.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-addFromSelection)|
|[bindingCollection](/resource/bindingcollection.md)|_Method_ > [getItemOrNull(id: string)](/resource/bindingcollection.md#getitemornullid-string)|Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-bindingCollection-getItemOrNull)|

# Chart
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[chartCollection](/resource/chartcollection.md)|_Method_ > [getItemOrNull(name: string)](/resource/chartcollection.md#getitemornullname-string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-chartCollection-getItemOrNull)|

# Named Item
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[namedItemCollection](/resource/nameditemcollection.md)|_Method_ > [getItemOrNull(name: string)](/resource/nameditemcollection.md#getitemornullname-string)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-namedItemCollection-getItemOrNull)|

# Pivot Table
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[pivotTable](/resource/pivottable.md)|_Property_ > name|Name of the pivot table.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=pivotTable-name)|
|[pivotTable](/resource/pivottable.md)|_Method_ > [refresh()](/resource/pivottable.md#refresh)|Refreshes the pivot table.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTable-refresh)|
|[pivotTableCollection](/resource/pivottablecollection.md)|_Property_ > count|Returns the number of pivot tables in the workbook or worksheet. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=pivotTableCollection-count)|
|[pivotTableCollection](/resource/pivottablecollection.md)|_Property_ > items|A collection of pivotTable objects. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=pivotTableCollection-items)|
|[pivotTableCollection](/resource/pivottablecollection.md)|_Method_ > [getItem(key: string)](/resource/pivottablecollection.md#getitemkey-string)|Gets a pivot table by Name.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-getItem)|
|[pivotTableCollection](/resource/pivottablecollection.md)|_Method_ > [getItemAt(index: number)](/resource/pivottablecollection.md#getitematindex-number)|Gets a pivot table based on its position in the collection.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-getItemAt)|
|[pivotTableCollection](/resource/pivottablecollection.md)|_Method_ > [getItemOrNull(key: string)](/resource/pivottablecollection.md#getitemornullkey-string)|Gets a pivot table by Name. If the pivot table does not exist, the return object's isNull property will be true.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-getItemOrNull)|
|[pivotTableCollection](/resource/pivottablecollection.md)|_Method_ > [refreshAll()](/resource/pivottablecollection.md#refreshall)|Refreshes all the pivot tables in the collection.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-pivotTableCollection-refreshAll)|

# Range
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[range](/resource/range.md)|_Relationship_ > visibleView|The visible part of the range. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=range-visibleView)|
|[range](/resource/range.md)|_Method_ > [getImage(height: number, width: number, fittingMode: ImageFittingMode)](/resource/range.md#getimageheight-number-width-number-fittingmode-imagefittingmode)|Renders the range as a base64-encoded image by scaling the range to fit the specified dimensions.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getImage)|
|[range](/resource/range.md)|_Method_ > [getIntersectionOrNull(anotherRange: Range or string)](/resource/range.md#getintersectionornullanotherrange-range-or-string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-range-getIntersectionOrNull)|

# RangeView
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[rangeView](/resource/rangeview.md)|_Property_ > address|Represents the range reference in A1-style. Address value will contain the Sheet reference (e.g. Sheet1!A1:B4). Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-address)|
|[rangeView](/resource/rangeview.md)|_Property_ > addressLocal|Represents range reference for the specified range in the language of the user. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-addressLocal)|
|[rangeView](/resource/rangeview.md)|_Property_ > cellCount|Number of cells in the range. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-cellCount)|
|[rangeView](/resource/rangeview.md)|_Property_ > formulas|Represents the formula in A1-style notation.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-formulas)|
|[rangeView](/resource/rangeview.md)|_Property_ > formulasLocal|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-formulasLocal)|
|[rangeView](/resource/rangeview.md)|_Property_ > numberFormat|Represents Excel's number format code for the given cell.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-numberFormat)|
|[rangeView](/resource/rangeview.md)|_Property_ > rowCount|Returns the total number of rows in the range. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-rowCount)|
|[rangeView](/resource/rangeview.md)|_Property_ > text|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-text)|
|[rangeView](/resource/rangeview.md)|_Property_ > values|Represents the raw values of the specified range. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-values)|
|[rangeView](/resource/rangeview.md)|_Relationship_ > rows|Represents a collection of all the rows in the RangeView. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-rows)|
|[rangeView](/resource/rangeview.md)|_Relationship_ > valueTypes|Represents the type of data of each cell. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeView-valueTypes)|
|[rangeViewCollection](/resource/rangeviewcollection.md)|_Property_ > count|Returns the number of RangeView in this collection. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeViewCollection-count)|
|[rangeViewCollection](/resource/rangeviewcollection.md)|_Property_ > items|A collection of RangeView objects. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=rangeViewCollection-items)|
|[rangeViewCollection](/resource/rangeviewcollection.md)|_Method_ > [getItemAt(index: number)](/resource/rangeviewcollection.md#getitematindex-number)|Gets a rangeView of a certain row based on its position in the collection.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-rangeViewCollection-getItemAt)|
|[rangeViewCollection](/resource/rangeviewcollection.md)|_Method_ > [load(param: object)](/resource/rangeviewcollection.md#loadparam-object)|Fills the proxy object created in JavaScript layer with property and object values specified in the parameter.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-rangeViewCollection-load)|

# Table
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[table](/resource/table.md)|_Property_ > highlightFirstColumn|Indicates whether the first column contains special formatting.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=table-highlightFirstColumn)|
|[table](/resource/table.md)|_Property_ > highlightLastColumn|Indicates whether the last column contains special formatting.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=table-highlightLastColumn)|
|[table](/resource/table.md)|_Property_ > showBandedColumns|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=table-showBandedColumns)|
|[table](/resource/table.md)|_Property_ > showBandedRows|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=table-showBandedRows)|
|[table](/resource/table.md)|_Property_ > showFilterButton|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=table-showFilterButton)|
|[table](/resource/table.md)|_Method_ > [getImage(height: number, width: number, fittingMode: ImageFittingMode)](/resource/table.md#getimageheight-number-width-number-fittingmode-imagefittingmode)|Renders the table as a base64-encoded image by scaling the table to fit the specified dimensions.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-table-getImage)|
|[tableCollection](/resource/tablecollection.md)|_Method_ > [getItemOrNull(key: number or string)](/resource/tablecollection.md#getitemornullkey-number-or-string)|Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-tableCollection-getItemOrNull)|
|[tableColumnCollection](/resource/tablecolumncollection.md)|_Method_ > [getItemOrNull(key: number or string)](/resource/tablecolumncollection.md#getitemornullkey-number-or-string)|Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-tableColumnCollection-getItemOrNull)|
|[tableRowCollection](/resource/tablerowcollection.md)|_Method_ > [addEmptyRows(index: number, count: number)](/resource/tablerowcollection.md#addemptyrowsindex-number-count-number)|Adds multiple empty rows to the table and return the first row that is added.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-tableRowCollection-addEmptyRows)|
|[tableRowCollection](/resource/tablerowcollection.md)|_Method_ > [addMultipleRows(index: number, values: (boolean or string or number)[][])](/resource/tablerowcollection.md#addmultiplerowsindex-number-values-boolean-or-string-or-number)|Adds multiple new rows to the table. Return the first row that is added.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-tableRowCollection-addMultipleRows)|

# Workbook
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[workbook](/resource/workbook.md)|_Relationship_ > pivotTables|Returns sheet protection object for a worksheet. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=workbook-pivotTables)|

# Worksheet
|Resources|What's new| Description| Give feedback|
|:-----|:-----|:-----|:-----|
|[worksheet](/resource/worksheet.md)|_Relationship_ > pivotTables|Returns sheet protection object for a worksheet. Read-only.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=worksheet-pivotTables)|
|[worksheet](/resource/worksheet.md)|_Relationship_ > visibility|The Visibility of the worksheet.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=worksheet-visibility)|
|[worksheetCollection](/resource/worksheetcollection.md)|_Method_ > [getItemOrNull(key: string)](/resource/worksheetcollection.md#getitemornullkey-string)|Gets a worksheet object using its Name or ID. If the worksheet does not exist, the returned object's isNull property will be true.|[Go](/resource/https://github.com/OfficeDev/office-js-docs/issues/new?title=OpenSpec-worksheetCollection-getItemOrNull)|
