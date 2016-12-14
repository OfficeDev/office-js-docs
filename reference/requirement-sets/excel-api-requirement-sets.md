# Excel JavaScript API requirement sets

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md).

Excel add-ins run across multiple versions of Office, including Office 2016 for Windows, Office for iPad, Office for Mac, and Office Online. The following table lists the Excel requirement sets, the Office host applications that support that requirement set, and the build versions or number for those applications. 

|  Requirement set  |  Office 2016 for Windows*  |  Office 2016 for iPad  |  Office 2016 for Mac  | Office Online  |  Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| ExcelApi 1.5  | TBD|TBD |  TBD| TBD | TBD|
| ExcelApi 1.4  | Version 1612 or later| TBD |  TBD| January 2017 | TBD|
| ExcelApi 1.3  | Version 1608 (Build 7369.2055) or later| 1.27 or later |  15.27 or later| September 2016 | Version 1608 (Build 7601.6800) or later|
| ExcelApi 1.2  | Version 1601 (Build 6741.2088) or later | 1.21 or later | 15.22 or later| January 2016 ||
| ExcelApi 1.1  | Version 1509 (Build 4266.1001) or later | 1.19 or later | 15.20 or later| January 2016 ||

> **Note**: The build number for Office 2016 installed via MSI is 16.0.4266.1001. This version only contains the ExcelApi 1.1 requirement set.

To find out more about versions, build numbers, and Office Online Server, see:

- [Version and build numbers of update channel releases for Office 365 clients](https://technet.microsoft.com/en-us/library/mt592918.aspx)
- [What version of Office am I using?](https://support.office.com/en-us/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19?ui=en-US&rs=en-US&ad=US&fromAR=1)
- [Where you can find the version and build number for an Office 365 client application](https://technet.microsoft.com/en-us/library/mt592918.aspx#Anchor_1)
- [Office Online Server overview](https://technet.microsoft.com/en-us/library/jj219437(v=office.16).aspx)

## Office common API requirement sets
For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).


## What's new in Excel JavaScript API 1.4 and 1.5
The following are the new additions to the Excel JavaScript APIs in requirement set 1.4 and 1.5. 

|Object| What is new| Description|Feedback|
|:----|:----|:----|:----|
|[bindingCollection](reference/excel/bindingcollection.md)|_Method_ > [getCount()](reference/excel/bindingcollection.md#getcount)|Gets the number of bindings in the collection.|1.5|
|[bindingCollection](reference/excel/bindingcollection.md)|_Method_ > [getItemOrNullObject(id: string)](reference/excel/bindingcollection.md#getitemornullobjectid-string)|Gets a binding object by ID. If the binding object does not exist, will return a null object.|1.4|
|[chartCollection](reference/excel/chartcollection.md)|_Method_ > [getCount()](reference/excel/chartcollection.md#getcount)|Returns the number of charts in the worksheet.|1.5|
|[chartCollection](reference/excel/chartcollection.md)|_Method_ > [getItemOrNullObject(name: string)](reference/excel/chartcollection.md#getitemornullobjectname-string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.4|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getCount()](reference/excel/chartpointscollection.md#getcount)|Returns the number of chart points in the series.|1.5|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getFirst()](reference/excel/chartpointscollection.md#getfirst)|Gets the first point in the series.|1.5|
|[chartPointsCollection](reference/excel/chartpointscollection.md)|_Method_ > [getLast()](reference/excel/chartpointscollection.md#getlast)|Gets the last point in the series.|1.5|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getCount()](reference/excel/chartseriescollection.md#getcount)|Returns the number of series in the collection.|1.5|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getFirst()](reference/excel/chartseriescollection.md#getfirst)|Gets the first series in the collection.|1.5|
|[chartSeriesCollection](reference/excel/chartseriescollection.md)|_Method_ > [getLast()](reference/excel/chartseriescollection.md#getlast)|Gets the last series in the collection.|1.5|
|[colorScaleConditionalFormat](reference/excel/colorscaleconditionalformat.md)|_Property_ > threeColorScale|If true the color scale will have three points (minimum, midpoint, maximum), otherwise it will have two (minimum, maximum). Read-only.|1.5|
|[colorScaleConditionalFormat](reference/excel/colorscaleconditionalformat.md)|_Relationship_ > criteria|The criteria of the color scale. Midpoint is optional when using a two point color scale.|1.5|
|[conditionalColorScaleCriteria](reference/excel/conditionalcolorscalecriteria.md)|_Relationship_ > maximum|The maximum point Color Scale Criterion.|1.5|
|[conditionalColorScaleCriteria](reference/excel/conditionalcolorscalecriteria.md)|_Relationship_ > midpoint|The midpoint Color Scale Criterion if the color scale is a 3-color scale.|1.5|
|[conditionalColorScaleCriteria](reference/excel/conditionalcolorscalecriteria.md)|_Relationship_ > minimum|The minimum point Color Scale Criterion.|1.5|
|[conditionalColorScaleCriterion](reference/excel/conditionalcolorscalecriterion.md)|_Property_ > color|HTML color code representation of the color scale color. E.g. #FF0000 represents Red.|1.5|
|[conditionalColorScaleCriterion](reference/excel/conditionalcolorscalecriterion.md)|_Property_ > formula|A number, a formula, or null (if Type is LowestValue).|1.5|
|[conditionalColorScaleCriterion](reference/excel/conditionalcolorscalecriterion.md)|_Relationship_ > type|What the icon conditional formula should be based on.|1.5|
|[conditionalDataBarNegativeFormat](reference/excel/conditionaldatabarnegativeformat.md)|_Property_ > borderColor|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.5|
|[conditionalDataBarNegativeFormat](reference/excel/conditionaldatabarnegativeformat.md)|_Property_ > fillColor|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.5|
|[conditionalDataBarNegativeFormat](reference/excel/conditionaldatabarnegativeformat.md)|_Property_ > matchPositiveBorderColor|Boolean representation of whether or not the negative DataBar has the same border color as the positive DataBar.|1.5|
|[conditionalDataBarNegativeFormat](reference/excel/conditionaldatabarnegativeformat.md)|_Property_ > matchPositiveFillColor|Boolean representation of whether or not the negative DataBar has the same fill color as the positive DataBar.|1.5|
|[conditionalDataBarPositiveFormat](reference/excel/conditionaldatabarpositiveformat.md)|_Property_ > borderColor|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.5|
|[conditionalDataBarPositiveFormat](reference/excel/conditionaldatabarpositiveformat.md)|_Property_ > fillColor|HTML color code representing the fill color, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.5|
|[conditionalDataBarPositiveFormat](reference/excel/conditionaldatabarpositiveformat.md)|_Property_ > gradientFill|Boolean representation of whether or not the DataBar has a gradient.|1.5|
|[conditionalDataBarRule](reference/excel/conditionaldatabarrule.md)|_Property_ > formula|The formula, if required, to evaluate the databar rule on.|1.5|
|[conditionalDataBarRule](reference/excel/conditionaldatabarrule.md)|_Property_ > type|The type of rule for the databar. Possible values are: LowestValue, HighestValue, Number, Percent, Formula, Percentile, Automatic.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Property_ > priority|The priority (or index) within the conditional format collection that this conditional format currently exists in. Changing this also|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Property_ > stopIfTrue|If the conditions of this conditional format are met, no lower-priority formats shall take effect on that cell.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Property_ > type|A type of conditional format. Only one can be set at a time. Read-Only. Read-only. Possible values are: Custom, DataBar, ColorScale, IconSet.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > colorScale|Represents a Color Scale conditional format. Read-only.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > colorScaleOrNullObject|Represents a Color Scale conditional format. Read-only.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > custom|A custom conditional format and rule. Read-only.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > customOrNullObject|A custom conditional format and rule. Read-only.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > dataBar|Represents databars with customizable color, gradient, axis, and range format options. Read-only.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > dataBarOrNullObject|Represents databars with customizable color, gradient, axis, and range format options. Read-only.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > iconSet|Represents an IconSet conditional format. Read-only.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Relationship_ > iconSetOrNullObject|Represents an IconSet conditional format. Read-only.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Method_ > [delete()](reference/excel/conditionalformat.md#delete)|Deletes this conditional format.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Method_ > [getRange()](reference/excel/conditionalformat.md#getrange)|Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.|1.5|
|[conditionalFormat](reference/excel/conditionalformat.md)|_Method_ > [getRangeOrNullObject()](reference/excel/conditionalformat.md#getrangeornullobject)|Returns the range the conditonal format is applied to or a null object if the range is discontiguous. Read-only.|1.5|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Property_ > items|A collection of conditionalFormat objects. Read-only.|1.5|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Method_ > [add(type: string)](reference/excel/conditionalformatcollection.md#addtype-string)|Adds a new conditional format to the collection at the firsttop priority.|1.5|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Method_ > [clearAll()](reference/excel/conditionalformatcollection.md#clearall)|Clears all conditional formats active on the current specified range.|1.5|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Method_ > [getCount()](reference/excel/conditionalformatcollection.md#getcount)|Returns the number of conditional formats in the workbook. Read-only.|1.5|
|[conditionalFormatCollection](reference/excel/conditionalformatcollection.md)|_Method_ > [getItemAt(index: number)](reference/excel/conditionalformatcollection.md#getitematindex-number)|Returns a conditional format at the given index.|1.5|
|[conditionalFormatRule](reference/excel/conditionalformatrule.md)|_Property_ > formula1|The formula, if required, to evaluate the conditional format rule on.|1.5|
|[conditionalFormatRule](reference/excel/conditionalformatrule.md)|_Property_ > formula1Local|The formula, if required, to evaluate the conditional format rule on in the user's language.|1.5|
|[conditionalFormatRule](reference/excel/conditionalformatrule.md)|_Property_ > formula1R1C1|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|1.5|
|[conditionalFormatRule](reference/excel/conditionalformatrule.md)|_Property_ > formula2|The formula, if required, to evaluate the conditional format rule on.|1.5|
|[conditionalFormatRule](reference/excel/conditionalformatrule.md)|_Property_ > formula2Local|The formula, if required, to evaluate the conditional format rule on in the user's language.|1.5|
|[conditionalFormatRule](reference/excel/conditionalformatrule.md)|_Property_ > formula2R1C1|The formula, if required, to evaluate the conditional format rule on in R1C1-style notation.|1.5|
|[conditionalIconCriterion](reference/excel/conditionaliconcriterion.md)|_Property_ > formula|A number or a formula depending on the type.|1.5|
|[conditionalIconCriterion](reference/excel/conditionaliconcriterion.md)|_Relationship_ > customIcon|The custom icon for the current criterion if different from the default IconSet, else null will be returned.|1.5|
|[conditionalIconCriterion](reference/excel/conditionaliconcriterion.md)|_Relationship_ > operator|GreaterThan or GreaterThanOrEqual for each of the rule type for the Icon conditional format.|1.5|
|[conditionalIconCriterion](reference/excel/conditionaliconcriterion.md)|_Relationship_ > type|What the icon conditional formula should be based on.|1.5|
|[conditionalRangeBorder](reference/excel/conditionalrangeborder.md)|_Property_ > color|HTML color code representing the color of the border line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.5|
|[conditionalRangeBorder](reference/excel/conditionalrangeborder.md)|_Relationship_ > id|Represents border identifier. Read-only.|1.5|
|[conditionalRangeBorder](reference/excel/conditionalrangeborder.md)|_Relationship_ > sideIndex|Constant value that indicates the specific side of the border. Read-only.|1.5|
|[conditionalRangeBorder](reference/excel/conditionalrangeborder.md)|_Relationship_ > style|One of the constants of line style specifying the line style for the border. Possible values are: None, Continuous, Dash, DashDot, DashDotDot, Dot, Double, SlantDashDot.|1.5|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Property_ > count|Number of border objects in the collection. Read-only.|1.5|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Property_ > items|A collection of conditionalRangeBorder objects. Read-only.|1.5|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Relationship_ > bottom|Gets the top border Read-only.|1.5|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Relationship_ > left|Gets the top border Read-only.|1.5|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Relationship_ > right|Gets the top border Read-only.|1.5|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Relationship_ > top|Gets the top border Read-only.|1.5|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Method_ > [getItem(index: string)](reference/excel/conditionalrangebordercollection.md#getitemindex-string)|Gets a border object using its name|1.5|
|[conditionalRangeBorderCollection](reference/excel/conditionalrangebordercollection.md)|_Method_ > [getItemAt(index: number)](reference/excel/conditionalrangebordercollection.md#getitematindex-number)|Gets a border object using its index|1.5|
|[conditionalRangeFill](reference/excel/conditionalrangefill.md)|_Property_ > color|HTML color code representing the color of the fill, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.5|
|[conditionalRangeFill](reference/excel/conditionalrangefill.md)|_Method_ > [clear()](reference/excel/conditionalrangefill.md#clear)|Resets the fill.|1.5|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Property_ > bold|Represents the bold status of font.|1.5|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Property_ > color|HTML color code representation of the text color. E.g. #FF0000 represents Red.|1.5|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Property_ > italic|Represents the italic status of the font.|1.5|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Property_ > strikethrough|Represents the strikethrough status of the font.|1.5|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Relationship_ > underline|Type of underline applied to the font.|1.5|
|[conditionalRangeFont](reference/excel/conditionalrangefont.md)|_Method_ > [clear()](reference/excel/conditionalrangefont.md#clear)|Resets the font formats.|1.5|
|[conditionalRangeFormat](reference/excel/conditionalrangeformat.md)|_Property_ > numberFormat|Represents Excel's number format code for the given range. Cleared if null is passed in.|1.5|
|[conditionalRangeFormat](reference/excel/conditionalrangeformat.md)|_Relationship_ > borders|Collection of border objects that apply to the overall conditional format range. Read-only.|1.5|
|[conditionalRangeFormat](reference/excel/conditionalrangeformat.md)|_Relationship_ > fill|Returns the fill object defined on the overall conditional format range. Read-only.|1.5|
|[conditionalRangeFormat](reference/excel/conditionalrangeformat.md)|_Relationship_ > font|Returns the font object defined on the overall conditional format range. Read-only.|1.5|
|[customConditionalFormat](reference/excel/customconditionalformat.md)|_Relationship_ > format|Returns a format object, encapsulating the conditional formats font, fill, borders, and other properties. Read-only.|1.5|
|[customConditionalFormat](reference/excel/customconditionalformat.md)|_Relationship_ > rule|Represents the Rule object on this conditional format. Read-only.|1.5|
|[customXmlPart](reference/excel/customxmlpart.md)|_Property_ > id|The custom XML part's ID. Read-only.|1.4|
|[customXmlPart](reference/excel/customxmlpart.md)|_Property_ > namespaceUri|The custom XML part's namespace URI. Read-only.|1.4|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [delete()](reference/excel/customxmlpart.md#delete)|Deletes the custom XML part.|1.4|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteAttribute(xpath: string, namespaceMappings: object, name: string)](reference/excel/customxmlpart.md#deleteattributexpath-string-namespacemappings-object-name-string)|Deletes an attribute with the given name from the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [deleteElement(xpath: string, namespaceMappings: object)](reference/excel/customxmlpart.md#deleteelementxpath-string-namespacemappings-object)|Deletes the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [getXml()](reference/excel/customxmlpart.md#getxml)|Gets the custom XML part's full XML content.|1.4|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertAttribute(xpath: string, namespaceMappings: object, name: string, value: string)](reference/excel/customxmlpart.md#insertattributexpath-string-namespacemappings-object-name-string-value-string)|Inserts an attribute with the given name and value to the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [insertElement(xpath: string, xml: string, namespaceMappings: object, index: number)](reference/excel/customxmlpart.md#insertelementxpath-string-xml-string-namespacemappings-object-index-number)|Inserts the given XML under the parent element identified by xpath at child position index.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [query(xpath: string, namespaceMappings: object)](reference/excel/customxmlpart.md#queryxpath-string-namespacemappings-object)|Queries the XML content.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [setXml(xml: string)](reference/excel/customxmlpart.md#setxmlxml-string)|Sets the custom XML part's full XML content.|1.4|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateAttribute(xpath: string, namespaceMappings: object, name: string, value: string)](reference/excel/customxmlpart.md#updateattributexpath-string-namespacemappings-object-name-string-value-string)|Updates the value of an attribute with the given name of the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPart](reference/excel/customxmlpart.md)|_Method_ > [updateElement(xpath: string, xml: string, namespaceMappings: object)](reference/excel/customxmlpart.md#updateelementxpath-string-xml-string-namespacemappings-object)|Updates the XML of the element identified by xpath.|ApiSetAttribute.Spec|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Property_ > items|A collection of customXmlPart objects. Read-only.|1.4|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [add(xml: string)](reference/excel/customxmlpartcollection.md#addxml-string)|Adds a new custom XML part to the workbook.|1.4|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [getByNamespace(namespaceUri: string)](reference/excel/customxmlpartcollection.md#getbynamespacenamespaceuri-string)|Gets a new scoped collection of custom XML parts whose namespaces match the given namespace.|1.4|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [getCount()](reference/excel/customxmlpartcollection.md#getcount)|Gets the number of CustomXml parts in the collection.|1.4|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [getItem(id: string)](reference/excel/customxmlpartcollection.md#getitemid-string)|Gets a custom XML part based on its ID.|1.4|
|[customXmlPartCollection](reference/excel/customxmlpartcollection.md)|_Method_ > [getItemOrNullObject(id: string)](reference/excel/customxmlpartcollection.md#getitemornullobjectid-string)|Gets a custom XML part based on its ID.|1.4|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Property_ > items|A collection of customXmlPartScoped objects. Read-only.|1.4|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getCount()](reference/excel/customxmlpartscopedcollection.md#getcount)|Gets the number of CustomXML parts in this collection.|1.4|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getItem(id: string)](reference/excel/customxmlpartscopedcollection.md#getitemid-string)|Gets a custom XML part based on its ID.|1.4|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getItemOrNullObject(id: string)](reference/excel/customxmlpartscopedcollection.md#getitemornullobjectid-string)|Gets a custom XML part based on its ID.|1.4|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getOnlyItem()](reference/excel/customxmlpartscopedcollection.md#getonlyitem)|If the collection contains exactly one item, this method returns it.|1.4|
|[customXmlPartScopedCollection](reference/excel/customxmlpartscopedcollection.md)|_Method_ > [getOnlyItemOrNullObject()](reference/excel/customxmlpartscopedcollection.md#getonlyitemornullobject)|If the collection contains exactly one item, this method returns it.|1.4|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Property_ > axisColor|HTML color code representing the color of the Axis line, of the form #RRGGBB (e.g. "FFA500") or as a named HTML color (e.g. "orange").|1.5|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Property_ > showDataBarOnly|If true, hides the values from the cells where the data bar is applied.|1.5|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > axisFormat|Representation of how the axis is determined for an Excel data bar.|1.5|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > barDirection|Represents the direction that the data bar graphic should be based on.|1.5|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > lowerBoundRule|The rule for what consistutes the lower bound (and how to calculate it, if applicable) for a data bar.|1.5|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > negativeFormat|Representation of all values to the left of the axis in an Excel data bar. Read-only.|1.5|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > positiveFormat|Representation of all values to the right of the axis in an Excel data bar. Read-only.|1.5|
|[dataBarConditionalFormat](reference/excel/databarconditionalformat.md)|_Relationship_ > upperBoundRule|The rule for what constitutes the upper bound (and how to calculate it, if applicable) for a data bar.|1.5|
|[iconSetConditionalFormat](reference/excel/iconsetconditionalformat.md)|_Property_ > reverseIconOrder|If true, reverses the icon orders for the IconSet. Note that this cannot be set if custom icons are used.|1.5|
|[iconSetConditionalFormat](reference/excel/iconsetconditionalformat.md)|_Property_ > showIconOnly|If true, hides the values and only shows icons.|1.5|
|[iconSetConditionalFormat](reference/excel/iconsetconditionalformat.md)|_Property_ > style|If set, displays the IconSet option for the conditional format. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.5|
|[iconSetConditionalFormat](reference/excel/iconsetconditionalformat.md)|_Relationship_ > criteria|An array of Criteria and IconSets for the rules and potential custom icons for conditional icons. Note that for the first criterion only the custom icon can be modified, while type, formula and operator will be ignored when set.|1.5|
|[namedItem](reference/excel/nameditem.md)|_Property_ > comment|Represents the comment associated with this name.|1.4|
|[namedItem](reference/excel/nameditem.md)|_Property_ > scope|Indicates whether the name is scoped to the workbook or to a specific worksheet. Read-only. Possible values are: Equal, Greater, GreaterEqual, Less, LessEqual, NotEqual.|1.4|
|[namedItem](reference/excel/nameditem.md)|_Relationship_ > worksheet|Returns the worksheet on which the named item is scoped to. Throws an error if the items is scoped to the workbook instead. Read-only.|1.4|
|[namedItem](reference/excel/nameditem.md)|_Relationship_ > worksheetOrNullObject|Returns the worksheet on which the named item is scoped to. Returns a null object if the item is scoped to the workbook instead. Read-only.|1.4|
|[namedItem](reference/excel/nameditem.md)|_Method_ > [delete()](reference/excel/nameditem.md#delete)|Deletes the given name.|1.4|
|[namedItem](reference/excel/nameditem.md)|_Method_ > [getRange()](reference/excel/nameditem.md#getrange)|Returns the range object that is associated with the name. Throws an error if the named item's type is not a range.|1.4|
|[namedItem](reference/excel/nameditem.md)|_Method_ > [getRangeOrNullObject()](reference/excel/nameditem.md#getrangeornullobject)|Returns the range object that is associated with the name. Returns a null object if the named item's type is not a range.|1.4|
|[namedItemCollection](reference/excel/nameditemcollection.md)|_Method_ > [add(name: string, reference: Range or string, comment: string)](reference/excel/nameditemcollection.md#addname-string-reference-range-or-string-comment-string)|Adds a new name to the collection of the given scope.|1.4|
|[namedItemCollection](reference/excel/nameditemcollection.md)|_Method_ > [addFormulaLocal(name: string, formula: string, comment: string)](reference/excel/nameditemcollection.md#addformulalocalname-string-formula-string-comment-string)|Adds a new name to the collection of the given scope using the user's locale for the formula.|1.4|
|[namedItemCollection](reference/excel/nameditemcollection.md)|_Method_ > [getCount()](reference/excel/nameditemcollection.md#getcount)|Gets the number of named items in the collection.|1.5|
|[namedItemCollection](reference/excel/nameditemcollection.md)|_Method_ > [getItemOrNullObject(name: string)](reference/excel/nameditemcollection.md#getitemornullobjectname-string)|Gets a nameditem object using its name. If the nameditem object does not exist, will return a null object.|1.4|
|[pivotTableCollection](reference/excel/pivottablecollection.md)|_Method_ > [getCount()](reference/excel/pivottablecollection.md#getcount)|Gets the number of pivot tables in the collection.|1.5|
|[pivotTableCollection](reference/excel/pivottablecollection.md)|_Method_ > [getItemOrNullObject(name: string)](reference/excel/pivottablecollection.md#getitemornullobjectname-string)|Gets a PivotTable by name. If the PivotTable does not exist, will return a null object.|1.4|
|[range](reference/excel/range.md)|_Relationship_ > conditionalFormats|Collection of ConditionalFormats that intersect the range. Read-only.|1.5|
|[range](reference/excel/range.md)|_Method_ > [getColumnsAfter(count: number)](reference/excel/range.md#getcolumnsaftercount-number)|Gets a certain number of columns to the right of the current Range object.|ApiSet.PolyfillableDownTo1_1, 1.3|
|[range](reference/excel/range.md)|_Method_ > [getColumnsBefore(count: number)](reference/excel/range.md#getcolumnsbeforecount-number)|Gets a certain number of columns to the left of the current Range object.|ApiSet.PolyfillableDownTo1_1, 1.3|
|[range](reference/excel/range.md)|_Method_ > [getIntersectionOrNullObject(anotherRange: Range or string)](reference/excel/range.md#getintersectionornullobjectanotherrange-range-or-string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.4|
|[range](reference/excel/range.md)|_Method_ > [getResizedRange(deltaRows: number, deltaColumns: number)](reference/excel/range.md#getresizedrangedeltarows-number-deltacolumns-number)|Gets a Range object similar to the current Range object, but with its bottom-right corner expanded (or contracted) by some number of rows and columns.|ApiSet.PolyfillableDownTo1_1, 1.3|
|[range](reference/excel/range.md)|_Method_ > [getRowsAbove(count: number)](reference/excel/range.md#getrowsabovecount-number)|Gets a certain number of rows above the current Range object.|ApiSet.PolyfillableDownTo1_1, 1.3|
|[range](reference/excel/range.md)|_Method_ > [getRowsBelow(count: number)](reference/excel/range.md#getrowsbelowcount-number)|Gets a certain number of rows below the current Range object.|ApiSet.PolyfillableDownTo1_1, 1.3|
|[range](reference/excel/range.md)|_Method_ > [getSurroundingRegion()](reference/excel/range.md#getsurroundingregion)|Returns a Range object that represents the surrounding region for this range. A surrounding region is a range bounded by any combination of blank rows and blank columns relative to this range.|1.5|
|[range](reference/excel/range.md)|_Method_ > [getUsedRangeOrNullObject(valuesOnly: bool)](reference/excel/range.md#getusedrangeornullobjectvaluesonly-bool)|Returns the used range of the given range object. If there are no used cells within the range, this function will return a null object.|1.4|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getCount()](reference/excel/rangeviewcollection.md#getcount)|Gets the number of RangeView objects in the collection.|1.5|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getFirst()](reference/excel/rangeviewcollection.md#getfirst)|Gets the first RangeView object in the collection.|1.5|
|[rangeViewCollection](reference/excel/rangeviewcollection.md)|_Method_ > [getLast()](reference/excel/rangeviewcollection.md#getlast)|Gets the last RangeView object in the collection.|1.5|
|[setting](reference/excel/setting.md)|_Property_ > key|Returns the key that represents the id of the Setting. Read-only.|1.4|
|[setting](reference/excel/setting.md)|_Property_ > value|Represents the value stored for this setting.|1.4|
|[setting](reference/excel/setting.md)|_Method_ > [delete()](reference/excel/setting.md#delete)|Deletes the setting.|1.4|
|[settingCollection](reference/excel/settingcollection.md)|_Property_ > items|A collection of setting objects. Read-only.|1.4|
|[settingCollection](reference/excel/settingcollection.md)|_Method_ > [add(key: string, value: (any)[])](reference/excel/settingcollection.md#addkey-string-value-any)|Sets or adds the specified setting to the workbook.|1.4|
|[settingCollection](reference/excel/settingcollection.md)|_Method_ > [getCount()](reference/excel/settingcollection.md#getcount)|Gets the number of Settings in the collection.|1.5|
|[settingCollection](reference/excel/settingcollection.md)|_Method_ > [getItem(key: string)](reference/excel/settingcollection.md#getitemkey-string)|Gets a Setting entry via the key.|1.4|
|[settingCollection](reference/excel/settingcollection.md)|_Method_ > [getItemOrNullObject(key: string)](reference/excel/settingcollection.md#getitemornullobjectkey-string)|Gets a Setting entry via the key. If the Setting does not exist, will return a null object.|1.4|
|[tableCollection](reference/excel/tablecollection.md)|_Method_ > [getCount()](reference/excel/tablecollection.md#getcount)|Gets the number of tables in the collection.|1.5|
|[tableCollection](reference/excel/tablecollection.md)|_Method_ > [getItemOrNullObject(key: number or string)](reference/excel/tablecollection.md#getitemornullobjectkey-number-or-string)|Gets a table by Name or ID. If the table does not exist, will return a null object.|1.4|
|[tableColumn](reference/excel/tablecolumn.md)|_Property_ > name|Represents the name of the table column.|1.1, 1.1 for getting the name; 1.4 for setting it.|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumn()](reference/excel/tablecolumn.md#getnextcolumn)|Gets the table column that follows this one. If there are no table columns following this one, this method will throw an error.|1.5|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getNextColumnOrNullObject()](reference/excel/tablecolumn.md#getnextcolumnornullobject)|Gets the table column that follows this one. If there are no table columns following this one, this method will return a null object.|1.5|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumn()](reference/excel/tablecolumn.md#getpreviouscolumn)|Gets the table column that precedes this one. If there are no previous table columns, this method will throw an error.|1.5|
|[tableColumn](reference/excel/tablecolumn.md)|_Method_ > [getPreviousColumnOrNullObject()](reference/excel/tablecolumn.md#getpreviouscolumnornullobject)|Gets the table column that precedes this one. If there are no previous table columns, this method will return a null objet.|1.5|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [add(index: number, values: (boolean or string or number)[][], name: string)](reference/excel/tablecolumncollection.md#addindex-number-values-boolean-or-string-or-number-name-string)|Adds a new column to the table.|1.1, 1.1 requires an index smaller than the total column count; 1.4 allows index to be optional (null or -1|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getCount()](reference/excel/tablecolumncollection.md#getcount)|Gets the number of columns in the table.|1.5|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getFirst()](reference/excel/tablecolumncollection.md#getfirst)|Gets the first column in the table.|1.5|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getItemOrNullObject(key: number or string)](reference/excel/tablecolumncollection.md#getitemornullobjectkey-number-or-string)|Gets a column object by Name or ID. If the column does not exist, will return a null object.|1.4|
|[tableColumnCollection](reference/excel/tablecolumncollection.md)|_Method_ > [getLast()](reference/excel/tablecolumncollection.md#getlast)|Gets the last column in the table.|1.5|
|[tableRowCollection](reference/excel/tablerowcollection.md)|_Method_ > [add(index: number, values: (boolean or string or number)[][])](reference/excel/tablerowcollection.md#addindex-number-values-boolean-or-string-or-number)|Adds one or more rows to the table. The return object will be the top of the newly added row(s).|1.1, 1.1 for adding a single row; 1.4 allows adding of multiple rows.|
|[tableRowCollection](reference/excel/tablerowcollection.md)|_Method_ > [getCount()](reference/excel/tablerowcollection.md#getcount)|Gets the number of rows in the table.|1.5|
|[workbook](reference/excel/workbook.md)|_Relationship_ > customXmlParts|Represents the collection of custom XML parts contained by this workbook. Read-only.|1.4|
|[workbook](reference/excel/workbook.md)|_Relationship_ > settings|Represents a collection of Settings associated with the workbook. Read-only.|1.4|
|[workbook](reference/excel/workbook.md)|_Method_ > [getActiveWorksheet()](reference/excel/workbook.md#getactiveworksheet)|Gets the currently active worksheet in the workbook. This method is equivalent to "worksheets.getActiveWorksheet()".|1.5|
|[workbook](reference/excel/workbook.md)|_Method_ > [getRange(address: string)](reference/excel/workbook.md#getrangeaddress-string)|Gets the range object specified by the address or name.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Property_ > visibility|The Visibility of the worksheet. Possible values are: Visible, Hidden, VeryHidden.|1.1, 1.1 for reading visibility; 1.2 for setting it.|
|[worksheet](reference/excel/worksheet.md)|_Relationship_ > names|Collection of names scoped to the current worksheet. Read-only.|1.4|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getBoundingRange(ranges: ()[])](reference/excel/worksheet.md#getboundingrangeranges-)|Gets the smallest range object that encompasses the provided ranges. For example, the bounding range between "B2:C5" and "D10:E15" is "B2:E16".|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getNextWorksheet(visibleOnly: bool)](reference/excel/worksheet.md#getnextworksheetvisibleonly-bool)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will throw an error.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getNextWorksheetOrNullObject(visibleOnly: bool)](reference/excel/worksheet.md#getnextworksheetornullobjectvisibleonly-bool)|Gets the worksheet that follows this one. If there are no worksheets following this one, this method will return a null object.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getPreviousWorksheet(visibleOnly: bool)](reference/excel/worksheet.md#getpreviousworksheetvisibleonly-bool)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will throw an error.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getPreviousWorksheetOrNullObject(visibleOnly: bool)](reference/excel/worksheet.md#getpreviousworksheetornullobjectvisibleonly-bool)|Gets the worksheet that precedes this one. If there are no previous worksheets, this method will return a null objet.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getRangeR1C1(startRow: number, startColumn: number, rowCount: number, columnCount: number)](reference/excel/worksheet.md#getranger1c1startrow-number-startcolumn-number-rowcount-number-columncount-number)|Gets the range object beginning at a particular row index and column index, and spanning a certain number of rows and columns.|1.5|
|[worksheet](reference/excel/worksheet.md)|_Method_ > [getUsedRangeOrNullObject(valuesOnly: bool)](reference/excel/worksheet.md#getusedrangeornullobjectvaluesonly-bool)|The used range is the smallest range that encompasses any cells that have a value or formatting assigned to them. If the entire worksheet is blank, this function will return a null object.|1.4|
|[worksheetCollection](reference/excel/worksheetcollection.md)|_Method_ > [getCount()](reference/excel/worksheetcollection.md#getcount)|Gets the number of worksheets in the collection.|1.5|
|[worksheetCollection](reference/excel/worksheetcollection.md)|_Method_ > [getFirst(visibleOnly: bool)](reference/excel/worksheetcollection.md#getfirstvisibleonly-bool)|Gets the first worksheet in the collection.|1.5|
|[worksheetCollection](reference/excel/worksheetcollection.md)|_Method_ > [getItemOrNullObject(key: string)](reference/excel/worksheetcollection.md#getitemornullobjectkey-string)|Gets a worksheet object using its Name or ID. If the worksheet does not exist, will return a null object.|1.4|
|[worksheetCollection](reference/excel/worksheetcollection.md)|_Method_ > [getLast(visibleOnly: bool)](reference/excel/worksheetcollection.md#getlastvisibleonly-bool)|Gets the last worksheet in the collection.|1.5|


## What's new in Excel JavaScript API 1.3 
The following are the new additions to the Excel JavaScript APIs in requirement set 1.3. 

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[binding](../excel/binding.md)|_Method_ > [delete()](../excel/binding.md#delete)|Deletes the binding.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [add(range: Range or string, bindingType: string, id: string)](../excel/bindingcollection.md#addrange-range-or-string-bindingtype-string-id-string)|Add a new binding to a particular Range.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [addFromNamedItem(name: string, bindingType: string, id: string)](../excel/bindingcollection.md#addfromnameditemname-string-bindingtype-string-id-string)|Add a new binding based on a named item in the workbook.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [addFromSelection(bindingType: string, id: string)](../excel/bindingcollection.md#addfromselectionbindingtype-string-id-string)|Add a new binding based on the current selection.|1.3|
|[bindingCollection](../excel/bindingcollection.md)|_Method_ > [getItemOrNull(id: string)](../excel/bindingcollection.md#getitemornullid-string)|Gets a binding object by ID. If the binding object does not exist, the return object's isNull property will be true.|1.3|
|[chartCollection](../excel/chartcollection.md)|_Method_ > [getItemOrNull(name: string)](../excel/chartcollection.md#getitemornullname-string)|Gets a chart using its name. If there are multiple charts with the same name, the first one will be returned.|1.3|
|[namedItemCollection](../excel/nameditemcollection.md)|_Method_ > [getItemOrNull(name: string)](../excel/nameditemcollection.md#getitemornullname-string)|Gets a nameditem object using its name. If the nameditem object does not exist, the returned object's isNull property will be true.|1.3|
|[pivotTable](../excel/pivottable.md)|_Property_ > name|Name of the PivotTable.|1.3|
|[pivotTable](../excel/pivottable.md)|_Relationship_ > worksheet|The worksheet containing the current PivotTable. Read-only.|1.3|
|[pivotTable](../excel/pivottable.md)|_Method_ > [refresh()](../excel/pivottable.md#refresh)|Refreshes the PivotTable.|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Property_ > items|A collection of pivotTable objects. Read-only.|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Method_ > [getItem(name: string)](../excel/pivottablecollection.md#getitemname-string)|Gets a PivotTable by name.|1.3|
|[pivotTableCollection](../excel/pivottablecollection.md)|_Method_ > [getItemOrNull(name: string)](../excel/pivottablecollection.md#getitemornullname-string)|Gets a PivotTable by name. If the PivotTable does not exist, the return object's isNull property will be true.|1.3|
|[range](../excel/range.md)|_Method_ > [getIntersectionOrNull(anotherRange: Range or string)](../excel/range.md#getintersectionornullanotherrange-range-or-string)|Gets the range object that represents the rectangular intersection of the given ranges. If no intersection is found, will return a null object.|1.3|
|[range](../excel/range.md)|_Method_ > [getVisibleView()](../excel/range.md#getvisibleview)|Represents the visible rows of the current range.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > cellAddresses|Represents the cell addresses of the RangeView. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > columnCount|Returns the number of visible columns. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > formulas|Represents the formula in A1-style notation.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > formulasLocal|Represents the formula in A1-style notation, in the user's language and number-formatting locale.  For example, the English "=SUM(A1, introduced in 1.5)" formula would become "=SUMME(A1; 1,5)" in German.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > index|Returns a value that represents the index of the RangeView. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > numberFormat|Represents Excel's number format code for the given cell.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > rowCount|Returns the number of visible rows. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > text|Text values of the specified range. The Text value will not depend on the cell width. The # sign substitution that happens in Excel UI will not affect the text value returned by the API. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > valueTypes|Represents the type of data of each cell. Read-only. Possible values are: Unknown, Empty, String, Integer, Double, Boolean, Error.|1.3|
|[rangeView](../excel/rangeview.md)|_Property_ > values|Represents the raw values of the specified range view. The data returned could be of type string, number, or a boolean. Cell that contain an error will return the error string.|1.3|
|[rangeView](../excel/rangeview.md)|_Relationship_ > rows|Represents a collection of range views associated with the range. Read-only.|1.3|
|[rangeView](../excel/rangeview.md)|_Method_ > [getRange()](../excel/rangeview.md#getrange)|Gets the parent range associated with the current RangeView.|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_Property_ > items|A collection of rangeView objects. Read-only.|1.3|
|[rangeViewCollection](../excel/rangeviewcollection.md)|_Method_ > [getItemAt(index: number)](../excel/rangeviewcollection.md#getitematindex-number)|Gets a RangeView Row via it's index. Zero-Indexed.|1.3|
|[setting](../excel/setting.md)|_Property_ > key|Returns the key that represents the id of the Setting. Read-only.|1.3|
|[setting](../excel/setting.md)|_Method_ > [delete()](../excel/setting.md#delete)|Deletes the setting.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Property_ > items|A collection of setting objects. Read-only.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [getItem(key: string)](../excel/settingcollection.md#getitemkey-string)|Gets a Setting entry via the key.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [getItemOrNull(key: string)](../excel/settingcollection.md#getitemornullkey-string)|Gets a Setting entry via the key. If the Setting does not exist, the returned object's isNull property will be true.|1.3|
|[settingCollection](../excel/settingcollection.md)|_Method_ > [set(key: string, value: string)](../excel/settingcollection.md#setkey-string-value-string)|Sets or adds the specified setting to the workbook.|1.3|
|[settingsChangedEventArgs](../excel/settingschangedeventargs.md)|_Relationship_ > settingCollection|Gets the Setting object that represents the binding that raised the SettingsChanged event|1.3|
|[table](../excel/table.md)|_Property_ > highlightFirstColumn|Indicates whether the first column contains special formatting.|1.3|
|[table](../excel/table.md)|_Property_ > highlightLastColumn|Indicates whether the last column contains special formatting.|1.3|
|[table](../excel/table.md)|_Property_ > showBandedColumns|Indicates whether the columns show banded formatting in which odd columns are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](../excel/table.md)|_Property_ > showBandedRows|Indicates whether the rows show banded formatting in which odd rows are highlighted differently from even ones to make reading the table easier.|1.3|
|[table](../excel/table.md)|_Property_ > showFilterButton|Indicates whether the filter buttons are visible at the top of each column header. Setting this is only allowed if the table contains a header row.|1.3|
|[tableCollection](../excel/tablecollection.md)|_Method_ > [getItemOrNull(key: number or string)](../excel/tablecollection.md#getitemornullkey-number-or-string)|Gets a table by Name or ID. If the table does not exist, the return object's isNull property will be true.|1.3|
|[tableColumnCollection](../excel/tablecolumncollection.md)|_Method_ > [getItemOrNull(key: number or string)](../excel/tablecolumncollection.md#getitemornullkey-number-or-string)|Gets a column object by Name or ID. If the column does not exist, the returned object's isNull property will be true.|1.3|
|[workbook](../excel/workbook.md)|_Relationship_ > pivotTables|Represents a collection of PivotTables associated with the workbook. Read-only.|1.3|
|[workbook](../excel/workbook.md)|_Relationship_ > settings|Represents a collection of Settings associated with the workbook. Read-only.|1.3|
|[worksheet](../excel/worksheet.md)|_Relationship_ > pivotTables|Collection of PivotTables that are part of the worksheet. Read-only.|1.3|

## What's new in Excel JavaScript API 1.2
The following are the new additions to the Excel JavaScript APIs in requirement set 1.2. 

|Object| What's new| Description|Requirement set|
|:----|:----|:----|:----|
|[chart](../excel/chart.md)|_Property_ > id|Gets a chart based on its position in the collection. Read-only.|1.2|
|[chart](../excel/chart.md)|_Relationship_ > worksheet|The worksheet containing the current chart. Read-only.|1.2|
|[chart](../excel/chart.md)|_Method_ > [getImage(height: number, width: number, fittingMode: string)](../excel/chart.md#getimageheight-number-width-number-fittingmode-string)|Renders the chart as a base64-encoded image by scaling the chart to fit the specified dimensions.|1.2|
|[filter](../excel/filter.md)|_Relationship_ > criteria|The currently applied filter on the given column. Read-only.|1.2|
|[filter](../excel/filter.md)|_Method_ > [apply(criteria: FilterCriteria)](../excel/filter.md#applycriteria-filtercriteria)|Apply the given filter criteria on the given column.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyBottomItemsFilter(count: number)](../excel/filter.md#applybottomitemsfiltercount-number)|Apply a "Bottom Item" filter to the column for the given number of elements.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyBottomPercentFilter(percent: number)](../excel/filter.md#applybottompercentfilterpercent-number)|Apply a "Bottom Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyCellColorFilter(color: string)](../excel/filter.md#applycellcolorfiltercolor-string)|Apply a "Cell Color" filter to the column for the given color.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyCustomFilter(criteria1: string, criteria2: string, oper: string)](../excel/filter.md#applycustomfiltercriteria1-string-criteria2-string-oper-string)|Apply a "Icon" filter to the column for the given criteria strings.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyDynamicFilter(criteria: string)](../excel/filter.md#applydynamicfiltercriteria-string)|Apply a "Dynamic" filter to the column.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyFontColorFilter(color: string)](../excel/filter.md#applyfontcolorfiltercolor-string)|Apply a "Font Color" filter to the column for the given color.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyIconFilter(icon: Icon)](../excel/filter.md#applyiconfiltericon-icon)|Apply a "Icon" filter to the column for the given icon.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyTopItemsFilter(count: number)](../excel/filter.md#applytopitemsfiltercount-number)|Apply a "Top Item" filter to the column for the given number of elements.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyTopPercentFilter(percent: number)](../excel/filter.md#applytoppercentfilterpercent-number)|Apply a "Top Percent" filter to the column for the given percentage of elements.|1.2|
|[filter](../excel/filter.md)|_Method_ > [applyValuesFilter(values: ()[])](../excel/filter.md#applyvaluesfiltervalues-)|Apply a "Values" filter to the column for the given values.|1.2|
|[filter](../excel/filter.md)|_Method_ > [clear()](../excel/filter.md#clear)|Clear the filter on the given column.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > color|The HTML color string used to filter cells. Used with "cellColor" and "fontColor" filtering.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > criterion1|The first criterion used to filter data. Used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > criterion2|The second criterion used to filter data. Only used as an operator in the case of "custom" filtering.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > dynamicCriteria|The dynamic criteria from the Excel.DynamicFilterCriteria set to apply on this column. Used with "dynamic" filtering. Possible values are: Unknown, AboveAverage, AllDatesInPeriodApril, AllDatesInPeriodAugust, AllDatesInPeriodDecember, AllDatesInPeriodFebruray, AllDatesInPeriodJanuary, AllDatesInPeriodJuly, AllDatesInPeriodJune, AllDatesInPeriodMarch, AllDatesInPeriodMay, AllDatesInPeriodNovember, AllDatesInPeriodOctober, AllDatesInPeriodQuarter1, AllDatesInPeriodQuarter2, AllDatesInPeriodQuarter3, AllDatesInPeriodQuarter4, AllDatesInPeriodSeptember, BelowAverage, LastMonth, LastQuarter, LastWeek, LastYear, NextMonth, NextQuarter, NextWeek, NextYear, ThisMonth, ThisQuarter, ThisWeek, ThisYear, Today, Tomorrow, YearToDate, Yesterday.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > filterOn|The property used by the filter to determine whether the values should stay visible. Possible values are: BottomItems, BottomPercent, CellColor, Dynamic, FontColor, Values, TopItems, TopPercent, Icon, Custom.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > operator|The operator used to combine criterion 1 and 2 when using "custom" filtering. Possible values are: And, Or.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Property_ > values|The set of values to be used as part of "values" filtering.|1.2|
|[filterCriteria](../excel/filtercriteria.md)|_Relationship_ > icon|The icon used to filter cells. Used with "icon" filtering.|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_Property_ > date|The date in ISO8601 format used to filter data.|1.2|
|[filterDatetime](../excel/filterdatetime.md)|_Property_ > specificity|How specific the date should be used to keep data. For example, if the date is 2005-04-02 and the specifity is set to "month", the filter operation will keep all rows with a date in the month of april 2009. Possible values are: Year, Monday, Day, Hour, Minute, Second.|1.2|
|[formatProtection](../excel/formatprotection.md)|_Property_ > formulaHidden|Indicates if Excel hides the formula for the cells in the range. A null value indicates that the entire range doesn't have uniform formula hidden setting.|1.2|
|[formatProtection](../excel/formatprotection.md)|_Property_ > locked|Indicates if Excel locks the cells in the object. A null value indicates that the entire range doesn't have uniform lock setting.|1.2|
|[icon](../excel/icon.md)|_Property_ > index|Represents the index of the icon in the given set.|1.2|
|[icon](../excel/icon.md)|_Property_ > set|Represents the set that the icon is part of. Possible values are: Invalid, ThreeArrows, ThreeArrowsGray, ThreeFlags, ThreeTrafficLights1, ThreeTrafficLights2, ThreeSigns, ThreeSymbols, ThreeSymbols2, FourArrows, FourArrowsGray, FourRedToBlack, FourRating, FourTrafficLights, FiveArrows, FiveArrowsGray, FiveRating, FiveQuarters, ThreeStars, ThreeTriangles, FiveBoxes.|1.2|
|[range](../excel/range.md)|_Property_ > columnHidden|Represents if all columns of the current range are hidden.|1.2|
|[range](../excel/range.md)|_Property_ > formulasR1C1|Represents the formula in R1C1-style notation.|1.2|
|[range](../excel/range.md)|_Property_ > hidden|Represents if all cells of the current range are hidden. Read-only.|1.2|
|[range](../excel/range.md)|_Property_ > rowHidden|Represents if all rows of the current range are hidden.|1.2|
|[range](../excel/range.md)|_Relationship_ > sort|Represents the range sort of the current range. Read-only.|1.2|
|[range](../excel/range.md)|_Method_ > [merge(across: bool)](../excel/range.md#mergeacross-bool)|Merge the range cells into one region in the worksheet.|1.2|
|[range](../excel/range.md)|_Method_ > [unmerge()](../excel/range.md#unmerge)|Unmerge the range cells into separate cells.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Property_ > columnWidth|Gets or sets the width of all colums within the range. If the column widths are not uniform, null will be returned.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Property_ > rowHeight|Gets or sets the height of all rows in the range. If the row heights are not uniform null will be returned.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Relationship_ > protection|Returns the format protection object for a range. Read-only.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Method_ > [autofitColumns()](../excel/rangeformat.md#autofitcolumns)|Changes the width of the columns of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeFormat](../excel/rangeformat.md)|_Method_ > [autofitRows()](../excel/rangeformat.md#autofitrows)|Changes the height of the rows of the current range to achieve the best fit, based on the current data in the columns.|1.2|
|[rangeReference](../excel/rangereference.md)|_Property_ > address|Represents the visible rows of the current range.|1.2|
|[rangeSort](../excel/rangesort.md)|_Method_ > [apply(fields: SortField[], matchCase: bool, hasHeaders: bool, orientation: string, method: string)](../excel/rangesort.md#applyfields-sortfield-matchcase-bool-hasheaders-bool-orientation-string-method-string)|Perform a sort operation.|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > ascending|Represents whether the sorting is done in an ascending fashion.|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > color|Represents the color that is the target of the condition if the sorting is on font or cell color.|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > dataOption|Represents additional sorting options for this field. Possible values are: Normal, TextAsNumber.|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > key|Represents the column (or row, depending on the sort orientation) that the condition is on. Represented as an offset from the first column (or row).|1.2|
|[sortField](../excel/sortfield.md)|_Property_ > sortOn|Represents the type of sorting of this condition. Possible values are: Value, CellColor, FontColor, Icon.|1.2|
|[sortField](../excel/sortfield.md)|_Relationship_ > icon|Represents the icon that is the target of the condition if the sorting is on the cell's icon.|1.2|
|[table](../excel/table.md)|_Relationship_ > sort|Represents the sorting for the table. Read-only.|1.2|
|[table](../excel/table.md)|_Relationship_ > worksheet|The worksheet containing the current table. Read-only.|1.2|
|[table](../excel/table.md)|_Method_ > [clearFilters()](../excel/table.md#clearfilters)|Clears all the filters currently applied on the table.|1.2|
|[table](../excel/table.md)|_Method_ > [convertToRange()](../excel/table.md#converttorange)|Converts the table into a normal range of cells. All data is preserved.|1.2|
|[table](../excel/table.md)|_Method_ > [reapplyFilters()](../excel/table.md#reapplyfilters)|Reapplies all the filters currently on the table.|1.2|
|[tableColumn](../excel/tablecolumn.md)|_Relationship_ > filter|Retrieve the filter applied to the column. Read-only.|1.2|
|[tableSort](../excel/tablesort.md)|_Property_ > matchCase|Represents whether the casing impacted the last sort of the table. Read-only.|1.2|
|[tableSort](../excel/tablesort.md)|_Property_ > method|Represents Chinese character ordering method last used to sort the table. Read-only. Possible values are: PinYin, StrokeCount.|1.2|
|[tableSort](../excel/tablesort.md)|_Relationship_ > fields|Represents the current conditions used to last sort the table. Read-only.|1.2|
|[tableSort](../excel/tablesort.md)|_Method_ > [apply(fields: SortField[], matchCase: bool, method: string)](../excel/tablesort.md#applyfields-sortfield-matchcase-bool-method-string)|Perform a sort operation.|1.2|
|[tableSort](../excel/tablesort.md)|_Method_ > [clear()](../excel/tablesort.md#clear)|Clears the sorting that is currently on the table. While this doesn't modify the table's ordering, it clears the state of the header buttons.|1.2|
|[tableSort](../excel/tablesort.md)|_Method_ > [reapply()](../excel/tablesort.md#reapply)|Reapplies the current sorting parameters to the table.|1.2|
|[workbook](../excel/workbook.md)|_Relationship_ > functions|Represents Excel application instance that contains this workbook. Read-only.|1.2|
|[worksheet](../excel/worksheet.md)|_Relationship_ > protection|Returns sheet protection object for a worksheet. Read-only.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Property_ > protected|Indicates if the worksheet is protected. Read-Only. Read-only.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Relationship_ > options|Sheet protection options. Read-only.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Method_ > [protect(options: WorksheetProtectionOptions)](../excel/worksheetprotection.md#protectoptions-worksheetprotectionoptions)|Protects a worksheet. Fails if the worksheet has been protected.|1.2|
|[worksheetProtection](../excel/worksheetprotection.md)|_Method_ > [unprotect()](../excel/worksheetprotection.md#unprotect)|Unprotects a worksheet.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowAutoFilter|Represents the worksheet protection option of allowing using auto filter feature.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowDeleteColumns|Represents the worksheet protection option of allowing deleting columns.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowDeleteRows|Represents the worksheet protection option of allowing deleting rows.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowFormatCells|Represents the worksheet protection option of allowing formatting cells.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowFormatColumns|Represents the worksheet protection option of allowing formatting columns.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowFormatRows|Represents the worksheet protection option of allowing formatting rows.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowInsertColumns|Represents the worksheet protection option of allowing inserting columns.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowInsertHyperlinks|Represents the worksheet protection option of allowing inserting hyperlinks.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowInsertRows|Represents the worksheet protection option of allowing inserting rows.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowPivotTables|Represents the worksheet protection option of allowing using PivotTable feature.|1.2|
|[worksheetProtectionOptions](../excel/worksheetprotectionoptions.md)|_Property_ > allowSort|Represents the worksheet protection option of allowing using sort feature.|1.2|

## Excel JavaScript API 1.1
Excel JavaScript API 1.1 is the first version of the API. For details about the API,  see the Excel JavaScript API reference topics.  
    
## Additional resources

- [Specify Office hosts and API requirements](../../docs/overview/specify-office-hosts-and-api-requirements.md)
- [Office Add-ins XML manifest](../../docs/overview/add-in-manifests.md)
