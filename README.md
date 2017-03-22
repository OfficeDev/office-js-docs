## Introduction to Excel JS APIs 1.5+

_Applies to: Excel 2016, Excel Online, Excel for iOS, Excel for Mac_

Below links describes the new set of Excel JavaScript APIs that are being planned for the next releases (Requirement Set 1.5 and beyond). Please review and provide your feedback. One of the best ways of providing your input is by opening new issue in GitHub using the links available below.

_**Note**: below listed features are still under design and review phase and hence not yet available as part of the product. The final design is subject to change. Once the feature is made available, the final specification will be published as part of the master repository._

## Upcoming Excel 1.5 Release Features

### Custom XML Part

* Addition of custom XML parts collection to workbook object.
* Get custom XML part using ID
* Get a new scoped collection of custom XML parts whose namespaces match the given namespace.
* Get XML string associated with a part.
* Provide id and namespace of a part.
* Adds a new custom XML part to the workbook.
* Set entire XML part.
* Delete a custom XML part.
* Delete an attribute with the given name from the element identified by xpath.
* Query the XML content by xpath.
* Insert, update and delete attribute.

**Reference implementation:**: Please refer [here](https://github.com/mandren/Excel-CustomXMLPart-Demo) for a reference implementation that shows how custom XML parts can be used in an add-in.

### Others
* `range.getSurroundingRegion()` Returns a Range object that represents the surrounding region for this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.
* `getNextColumn()` and `getPreviousColumn()`, `getLast() on table column.
* `getActiveWorksheet()` on the workbook.
* `getRange(address: string)` off of workbook.
* `getBoundingRange(ranges: [])` Gets the smallest range object that encompasses the provided ranges. For example, the bounding range between "B2:C5" and "D10:E15" is "B2:E15".
* `getCount()` on various collections such as named item, worksheet, table, etc. to get number of items in a collection. `workbook.worksheets.getCount()`
* `getFirst()` and `getLast()` and get last on various collection such as tworksheet, able column, chart points, range view collection.
* `getNext()` and `getPrevious()` on worksheet, table column collection.
* `getRangeR1C1()` Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.

## Upcoming Excel 1.6 Release Features

### Conditional formatting

Introduces [Conditional formating](https://github.com/OfficeDev/office-js-docs/blob/ExcelJs_1.4_OpenSpec/reference/excel/conditionalformat.md) of a range. Allows follwoing types of conditional formatting:

* Color scale
* Data bar
* Icon set
* Custom

In addiiton:
* Returns the range the conditonal format is applied to.
* Removal of conditional formatting.
* Provides priority and stopifTrue capability
* Get collection of all conditional formatting on a given range.
* Clears all conditional formats active on the current specified range.



|Object| What is new| Description |Requirement set|
|:----|:----|:----|:----|
|[cellValueConditionalFormat](reference/excel/cellvalueconditionalformat.md)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[cellValueConditionalFormat](reference/excel/cellvalueconditionalformat.md)|_Relationship_ > rule|Represents the Rule object on this conditional format.|1.6|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getFirst()](reference/excel/chartpointscollection.md#getfirst)|Gets the first point in the series.|1.5|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getLast()](reference/excel/chartpointscollection.md#getlast)|Gets the last point in the series.|1.5|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getFirst()](reference/excel/chartseriescollection.md#getfirst)|Gets the first series in the collection.|1.5|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getLast()](reference/excel/chartseriescollection.md#getlast)|Gets the last series in the collection.|1.5|
|[colorScaleConditionalFormat](reference/excel/colorscaleconditionalformat.md)|_Property_ > threeColorScale|If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum). Read-only.|1.6|
|[colorScaleConditionalFormat](reference/excel/colorscaleconditionalformat.md)|_Relationship_ > criteria|The criteria of the color scale. Midpoint is optional when using a two point color scale.|1.6|
|[conditionalCellValueRule](reference/excel/conditionalcellvaluerule.md)|_Property_ > formula1|The formula, if required, to evaluate the conditional format rule on.|1.6|
|[conditionalCellValueRule](reference/excel/conditionalcellvaluerule.md)|_Property_ > formula2|The formula, if required, to evaluate the conditional format rule on.|1.6|
|[conditionalCellValueRule](reference/excel/conditionalcellvaluerule.md)|_Property_ > operator|The operator of the text conditional format. Possible values are: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqual, LessThanOrEqual.|1.6|
|[conditionalColorScaleCriteria](reference/excel/conditionalcolorscalecriteria.md)|_Relationship_ > maximum|The maximum point Color Scale Criterion.|1.6|
|[conditionalColorScaleCriteria](reference/excel/conditionalcolorscalecriteria.md)|_Relationship_ > midpoint|The midpoint Color Scale Criterion if the color scale is a 3-color scale.|1.6|
|[conditionalColorScaleCriteria](reference/excel/conditionalcolorscalecriteria.md)|_Relationship_ > minimum|The minimum point Color Scale Criterion.|1.6|
|[conditionalColorScaleCriterion](reference/excel/conditionalcolorscalecriterion.md)|_Property_ > color|HTML color code representation of the color scale color. E.g. #FF0000 represents Red.|1.6|
|[conditionalColorScaleCriterion](reference/excel/conditionalcolorscalecriterion.md)|_Property_ > formula|A number, a formula, or null (if Type is LowestValue).|1.6|
|[conditionalColorScaleCriterion](reference/excel/conditionalcolorscalecriterion.md)|_Property_ > type|What the icon conditional formula should be based on. Possible values are: Invalid, LowestValue, HighestValue, Number, Percent, Formula, Percentile.|1.6|
|[conditionalDataBarNegativeFormat](reference/excel/conditionaldatabarnegativeformat.md)|_Property_ > borderColor|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarNegativeFormat](reference/excel/conditionaldatabarnegativeformat.md)|_Property_ > fillColor|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarNegativeFormat](reference/excel/conditionaldatabarnegativeformat.md)|_Property_ > matchPositiveBorderColor|Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.|1.6|
|[conditionalDataBarNegativeFormat](reference/excel/conditionaldatabarnegativeformat.md)|_Property_ > matchPositiveFillColor|Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.|1.6|
|[conditionalDataBarPositiveFormat](reference/excel/conditionaldatabarpositiveformat.md)|_Property_ > borderColor|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarPositiveFormat](reference/excel/conditionaldatabarpositiveformat.md)|_Property_ > fillColor|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalDataBarPositiveFormat](reference/excel/conditionaldatabarpositiveformat.md)|_Property_ > gradientFill|Boolean representation of whether or not the DataBar has a gradient.|1.6|
|[conditionalDataBarRule](reference/excel/conditionaldatabarrule.md)|_Property_ > formula|The formula, if required, to evaluate the databar rule on.|1.6|
|[conditionalDataBarRule](reference/excel/conditionaldatabarrule.md)|_Property_ > type|The type of rule for the databar. Possible values are: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Property_ > priority|The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Property_ > stopIfTrue|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Property_ > type|A type of conditional format. Only one can be set at a time. Read-Only. Read-only. Possible values are: Custom, DataBar, ColorScale, IconSet.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > cellValue|Returns the cell value conditional format properties if the current conditional format is a CellValue type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > cellValueOrNullObject|Returns the cell value conditional format properties if the current conditional format is a CellValue type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > colorScale|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > colorScaleOrNullObject|Returns the ColorScale conditional format properties if the current conditional format is an ColorScale type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > custom|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > customOrNullObject|Returns the custom conditional format properties if the current conditional format is a custom type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > dataBar|Returns the data bar properties if the current conditional format is a data bar. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > dataBarOrNullObject|Returns the data bar properties if the current conditional format is a data bar. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > iconSet|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > iconSetOrNullObject|Returns the IconSet conditional format properties if the current conditional format is an IconSet type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > preset|Returns the preset criteria conditional format such as above averagebelow averageunique valuescontains blanknonblankerrornoerror properties. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > presetOrNullObject|Returns the preset criteria conditional format such as above averagebelow averageunique valuescontains blanknonblankerrornoerror properties. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > textComparison|Returns the specific text conditional format properties if the current conditional format is a text type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > textComparisonOrNullObject|Returns the specific text conditional format properties if the current conditional format is a text type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > topBottom|Returns the TopBottom conditional format properties if the current conditional format is an TopBottom type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > topBottomOrNullObject|Returns the TopBottom conditional format properties if the current conditional format is an TopBottom type. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Method_ > [delete()](reference/excel/conditionalformat.md#delete)|Deletes this conditional format.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Method_ > [getRange()](reference/excel/conditionalformat.md#getrange)|Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.|1.6|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Method_ > [getRangeOrNullObject()](reference/excel/conditionalformat.md#getrangeornullobject)|Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.|1.6|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Property_ > items|A collection of conditionalFormat objects. Read-only.|1.6|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Method_ > [add(type: string)](reference/excel/conditionalformatcollection.md#addtype-string)|Adds a new conditional format to the collection at the firsttop priority.|1.6|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Method_ > [clearAll()](reference/excel/conditionalformatcollection.md#clearall)|Clears all conditional formats active on the current specified range.|1.6|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Method_ > [getCount()](reference/excel/conditionalformatcollection.md#getcount)|Returns the number of conditional formats in the workbook. Read-only.|1.6|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Method_ > [getItemAt(index: number)](reference/excel/conditionalformatcollection.md#getitematindex-number)|Returns a conditional format at the given index.|1.6|
|[conditionalFormatRule](reference/excel/conditionalformatrule.md)|_Property_ > formula|The formula, if required, to evaluate the conditional format rule on.|1.6|
|[conditionalFormatRule](reference/excel/conditionalformatrule.md)|_Property_ > formulaLocal|The formula, if required, to evaluate the conditional format rule on in the user's language.|1.6|
|[conditionalFormatRule](reference/excel/conditionalformatrule.md)|_Property_ > formulaR1C1|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|1.6|
|[conditionalIconCriterion](reference/excel/conditionaliconcriterion.md)|_Property_ > formula|A number or a formula depending on the type.|1.6|
|[conditionalIconCriterion](reference/excel/conditionaliconcriterion.md)|_Property_ > operator|GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format. Possible values are: Invalid, GreaterThan, GreaterThanOrEqual.|1.6|
|[conditionalIconCriterion](reference/excel/conditionaliconcriterion.md)|_Relationship_ > customIcon|The custom icon for the current criterion if different from the default IconSet, else null will be returned.|1.6|
|[conditionalIconCriterion](reference/excel/conditionaliconcriterion.md)|_Relationship_ > type|What the icon conditional formula should be based on.|1.6|
|[conditionalPresetCriteriaRule](reference/excel/conditionalpresetcriteriarule.md)|_Property_ > criterion|The criterion of the conditional format. Possible values are: Invalid, Blanks, NonBlanks, Errors, NonErrors, Yesterday, Today, Tomorrow, LastSevenDays, LastWeek, ThisWeek, NextWeek, LastMonth, ThisMonth, NextMonth, AboveAverage, BelowAverage, EqualOrAboveAverage, EqualOrBelowAverage, OneStdDevAboveAverage, OneStdDevBelowAverage, TwoStdDevAboveAverage, TwoStdDevBelowAverage, ThreeStdDevAboveAverage, ThreeStdDevBelowAverage, UniqueValues, DuplicateValues.|1.6|
|[conditionalRangeBorder](reference/excel/conditionalrangeborder.md)|_Property_ > color|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalRangeBorder](reference/excel/conditionalrangeborder.md)|_Property_ > id|Represents border identifier. Read-only. Possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](reference/excel/conditionalrangeborder.md)|_Property_ > sideIndex|Constant value that indicates the specific side of the border. Read-only. Possible values are: EdgeTop, EdgeBottom, EdgeLeft, EdgeRight.|1.6|
|[conditionalRangeBorder](reference/excel/conditionalrangeborder.md)|_Property_ > style|One of the constants of line style specifying the line style for the border. Possible values are: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.6|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Property_ > count|Number of border objects in the collection. Read-only.|1.6|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Property_ > items|A collection of conditionalRangeBorder objects. Read-only.|1.6|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Relationship_ > bottom|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Relationship_ > left|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Relationship_ > right|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Relationship_ > top|Gets the top border Read-only.|1.6|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Method_ > [getItem(index: string)](reference/excel/conditionalrangebordercollection.md#getitemindex-string)|Gets a border object using its name|1.6|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Method_ > [getItemAt(index: number)](reference/excel/conditionalrangebordercollection.md#getitematindex-number)|Gets a border object using its index|1.6|
|[conditionalRangeFill](reference/excel/conditionalrangefill.md)|_Property_ > color|HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[conditionalRangeFill](reference/excel/conditionalrangefill.md)|_Method_ > [clear()](reference/excel/conditionalrangefill.md#clear)|Resets the fill.|1.6|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Property_ > bold|Represents the bold status of font.|1.6|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Property_ > color|HTML color code representation of the text color. E.g. #FF0000 represents Red.|1.6|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Property_ > italic|Represents the italic status of the font.|1.6|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Property_ > strikethrough|Represents the strikethrough status of the font.|1.6|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Property_ > underline|Type of underline applied to the font. Possible values are: None, Single, Double.|1.6|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Method_ > [clear()](reference/excel/conditionalrangefont.md#clear)|Resets the font formats.|1.6|
|[conditionalRangeFormat](reference/excel/conditionalrangeformat.md)|_Property_ > numberFormat|Represents Excel's number format code for the given range. Cleared if null is passed in.|1.6|
|[conditionalRangeFormat](reference/excel/conditionalrangeformat.md)|_Relationship_ > borders|Collection of border objects that apply to the overall conditional format range. Read-only.|1.6|
|[conditionalRangeFormat](reference/excel/conditionalrangeformat.md)|_Relationship_ > fill|Returns the fill object defined on the overall conditional format range. Read-only.|1.6|
|[conditionalRangeFormat](reference/excel/conditionalrangeformat.md)|_Relationship_ > font|Returns the font object defined on the overall conditional format range. Read-only.|1.6|
|[conditionalTextComparisonRule](reference/excel/conditionaltextcomparisonrule.md)|_Property_ > operator|The operator of the text conditional format. Possible values are: Invalid, Contains, NotContains, BeginsWith, EndsWith.|1.6|
|[conditionalTextComparisonRule](reference/excel/conditionaltextcomparisonrule.md)|_Property_ > text|The Text value of conditional format.|1.6|
|[conditionalTopBottomRule](reference/excel/conditionaltopbottomrule.md)|_Property_ > rank|The rank between 1 and 1000 for numeric ranks or 1 and 100 for percent ranks.|1.6|
|[conditionalTopBottomRule](reference/excel/conditionaltopbottomrule.md)|_Property_ > type|Format values based on the top or bottom rank. Possible values are: Invalid, TopItems, TopPercent, BottomItems, BottomPercent.|1.6|
|[customConditionalFormat](reference/excel/customconditionalformat.md)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[customConditionalFormat](reference/excel/customconditionalformat.md)|_Relationship_ > rule|Represents the Rule object on this conditional format. Read-only.|1.6|
|[customXmlPart](reference/excel/customxmlpart.md)|_Property_ > id|The custom XML part's ID. Read-only.|1.5|
|[customXmlPart](reference/excel/customxmlpart.md)|_Property_ > namespaceUri|The custom XML part's namespace URI. Read-only.|1.5|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [delete()](reference/excel/customxmlpart.md#delete)|Deletes the custom XML part.|1.5|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteAttribute(xpath: string, namespaceMappings: object, name: string)](reference/excel/customxmlpart.md#deleteattributexpath-string-namespacemappings-object-name-string)|Deletes an attribute with the given name from the element identified by xpath.|TBD|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteElement(xpath: string, namespaceMappings: object)](reference/excel/customxmlpart.md#deleteelementxpath-string-namespacemappings-object)|Deletes the element identified by xpath.|TBD|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [getXml()](reference/excel/customxmlpart.md#getxml)|Gets the custom XML part's full XML content.|1.5|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertAttribute(xpath: string, namespaceMappings: object, name: string, value: string)](reference/excel/customxmlpart.md#insertattributexpath-string-namespacemappings-object-name-string-value-string)|Inserts an attribute with the given name and value to the element identified by xpath.|TBD|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertElement(xpath: string, xml: string, namespaceMappings: object, index: number)](reference/excel/customxmlpart.md#insertelementxpath-string-xml-string-namespacemappings-object-index-number)|Inserts the given XML under the parent element identified by xpath at child position index.|TBD|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [query(xpath: string, namespaceMappings: object)](reference/excel/customxmlpart.md#queryxpath-string-namespacemappings-object)|Queries the XML content.|TBD|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [setXml(xml: string)](reference/excel/customxmlpart.md#setxmlxml-string)|Sets the custom XML part's full XML content.|1.5|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateAttribute(xpath: string, namespaceMappings: object, name: string, value: string)](reference/excel/customxmlpart.md#updateattributexpath-string-namespacemappings-object-name-string-value-string)|Updates the value of an attribute with the given name of the element identified by xpath.|TBD|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateElement(xpath: string, xml: string, namespaceMappings: object)](reference/excel/customxmlpart.md#updateelementxpath-string-xml-string-namespacemappings-object)|Updates the XML of the element identified by xpath.|TBD|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Property_ > items|A collection of customXmlPart objects. Read-only.|1.5|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [add(xml: string)](reference/excel/customxmlpartcollection.md#addxml-string)|Adds a new custom XML part to the workbook.|1.5|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [getByNamespace(namespaceUri: string)](reference/excel/customxmlpartcollection.md#getbynamespacenamespaceuri-string)|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|1.5|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [getCount()](reference/excel/customxmlpartcollection.md#getcount)|Gets the number of CustomXml parts in the collection.|1.5|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [getItem(id: string)](reference/excel/customxmlpartcollection.md#getitemid-string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [getItemOrNullObject(id: string)](reference/excel/customxmlpartcollection.md#getitemornullobjectid-string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Property_ > items|A collection of customXmlPartScoped objects. Read-only.|1.5|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getCount()](reference/excel/customxmlpartscopedcollection.md#getcount)|Gets the number of CustomXML parts in this collection.|1.5|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getItem(id: string)](reference/excel/customxmlpartscopedcollection.md#getitemid-string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getItemOrNullObject(id: string)](reference/excel/customxmlpartscopedcollection.md#getitemornullobjectid-string)|Gets a custom XML part based on its ID.|1.5|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getOnlyItem()](reference/excel/customxmlpartscopedcollection.md#getonlyitem)|If the collection contains exactly one item, this method returns it.|1.5|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getOnlyItemOrNullObject()](reference/excel/customxmlpartscopedcollection.md#getonlyitemornullobject)|If the collection contains exactly one item, this method returns it.|1.5|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Property_ > axisColor|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.6|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Property_ > axisFormat|Representation of how the axis is determined for an Excel data bar. Possible values are: Automatic, None, CellMidPoint.|1.6|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Property_ > barDirection|Represents the direction that the data bar graphic should be based on. Possible values are: Context, LeftToRight, RightToLeft.|1.6|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Property_ > showDataBarOnly|If true, hides the values from the cells where the data bar is applied.|1.6|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > lowerBoundRule|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|1.6|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > negativeFormat|Representation of all values to the left of the axis in an Excel data bar. Read-only.|1.6|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > positiveFormat|Representation of all values to the right of the axis in an Excel data bar. Read-only.|1.6|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > upperBoundRule|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|1.6|
|[iconSetConditionalFormat](reference/excel/iconsetconditionalformat.md)|_Property_ > reverseIconOrder|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|1.6|
|[iconSetConditionalFormat](reference/excel/iconsetconditionalformat.md)|_Property_ > showIconOnly|If true, hides the values and only shows icons.|1.6|
|[iconSetConditionalFormat](reference/excel/iconsetconditionalformat.md)|_Property_ > style|If set, displays the IconSet option for the conditional format. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.6|
|[iconSetConditionalFormat](reference/excel/iconsetconditionalformat.md)|_Relationship_ > criteria|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula and operator will be ignored when set.|1.6|
|[presetCriteriaConditionalFormat](reference/excel/presetcriteriaconditionalformat.md)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[presetCriteriaConditionalFormat](reference/excel/presetcriteriaconditionalformat.md)|_Relationship_ > rule|The rule of the conditional format.|1.6|
|[range](reference/excel/range.md)|_Relationship_ > conditionalFormats|Collection of ConditionalFormats that intersect the range. Read-only.|1.6|
|[range](reference/excel/range.md)|_Method_ > [getSurroundingRegion()](reference/excel/range.md#getsurroundingregion)|Returns a Range object that represents the surrounding region for this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|1.5|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getFirst()](reference/excel/rangeviewcollection.md#getfirst)|Gets the first RangeView object in the collection.|1.5|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getLast()](reference/excel/rangeviewcollection.md#getlast)|Gets the last RangeView object in the collection.|1.5|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumn()](reference/excel/tablecolumn.md#getnextcolumn)|Gets the table column that follows this one. If there are no table columns following this one, this method will throw an error.|1.5|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumnOrNullObject()](reference/excel/tablecolumn.md#getnextcolumnornullobject)|Gets the table column that follows this one. If there are no table columns following this one, this method will return a null object.|1.5|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumn()](reference/excel/tablecolumn.md#getpreviouscolumn)|Gets the table column that precedes this one. If there are no previous table columns, this method will throw an error.|1.5|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumnOrNullObject()](reference/excel/tablecolumn.md#getpreviouscolumnornullobject)|Gets the table column that precedes this one. If there are no previous table columns, this method will return a null objet.|1.5|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getFirst()](reference/excel/tablecolumncollection.md#getfirst)|Gets the first column in the table.|1.5|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getLast()](reference/excel/tablecolumncollection.md#getlast)|Gets the last column in the table.|1.5|
|[textConditionalFormat](reference/excel/textconditionalformat.md)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[textConditionalFormat](reference/excel/textconditionalformat.md)|_Relationship_ > rule|The rule of the conditional format.|1.6|
|[topBottomConditionalFormat](reference/excel/topbottomconditionalformat.md)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.6|
|[topBottomConditionalFormat](reference/excel/topbottomconditionalformat.md)|_Relationship_ > rule|The criteria of the TopBottom conditional format.|1.6|
|[workbook](reference/excel/workbook.md)|_Relationship_ > customXmlParts|Represents the collection of custom XML parts contained by this workbook. Read-only.|1.5|
|[workbook](reference/excel/workbook.md)|_Method_ > [getRange(address: string)](reference/excel/workbook.md#getrangeaddress-string)|Gets the range object specified by the address or name.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getNext(visibleOnly: bool)](reference/excel/worksheet.md#getnextvisibleonly-bool)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getNextOrNullObject(visibleOnly: bool)](reference/excel/worksheet.md#getnextornullobjectvisibleonly-bool)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getPrevious(visibleOnly: bool)](reference/excel/worksheet.md#getpreviousvisibleonly-bool)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getPreviousOrNullObject(visibleOnly: bool)](reference/excel/worksheet.md#getpreviousornullobjectvisibleonly-bool)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.|1.5|
|[worksheetCollection](reference/excel/worksheetcollection.md)|_Method_ > [getFirst(visibleOnly: bool)](reference/excel/worksheetcollection.md#getfirstvisibleonly-bool)|Gets the first worksheet in the collection.|1.5|
|[worksheetCollection](reference/excel/worksheetcollection.md)|_Method_ > [getLast(visibleOnly: bool)](reference/excel/worksheetcollection.md#getlastvisibleonly-bool)|Gets the last worksheet in the collection.|1.5|


## Give feedback

We need it, you want to give it. Feedback is much easier to give now that we're on GitHub. Check out the docs and let us know your feedback by submitting [issues](https://github.com/OfficeDev/office-js-docs/issues) directly in this repository.

For API support, you can post questions to the community on [StackOverflow](http://stackoverflow.com/questions/tagged/office-js) and tag them with [office-js].
