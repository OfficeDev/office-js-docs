|Object| What is new| Description | Requirement set|
|:----|:----|:----|:----|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending of the operator.|1.8|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > formula2|Gets or sets the Formula2, i.e. maximum value or value depending of the operator.|1.8|
|[basicDataValidation](reference/excel/basicdatavalidation.md)|_Property_ > operator|The operator to use for validating the data. Possible values are: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqualTo, LessThanOrEqualTo.|1.8|
|[customDataValidation](reference/excel/customdatavalidation.md)|_Property_ > formula|Custom data validation formula, it is to create special rules, such as preventing duplicates, or limiting the total in a range of cells.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Property_ > ignoreBlanks|Ignore blanks: no data validation will be performed on blank cells, it defaults to true.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Property_ > type|Type of the data validation, see Excel.DataValidationType for details. Read-only. Possible values are: Invalid, None, WholeNumber, Decimal, List, Date, Time, TextLength, Inconsistent, MixedCriteria.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > errorAlert|Error alert when user enters invalid data.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > prompt|Prompt when users select a cell.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Relationship_ > rule|Data Validation rule that contains different type of data validation criteria.|1.8|
|[dataValidation](reference/excel/datavalidation.md)|_Method_ > [clear()]((reference/excel/datavalidation.md#clear)|Clears the data validation from the current range.|1.8|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > message|Represents error alert message.|1.8|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > showAlert|It determines show error alert dialog or not when users enter invalid data, it defaults to true.|1.8|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > style|Represents Data validation alert type, please see Excel.DataValidation lertStyle for details. Possible values are: Invalid, Stop, Warning, Information.|1.8|
|[dataValidationErrorAlert](reference/excel/datavalidationerroralert.md)|_Property_ > title|Represents error alert dialog title.|1.8|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > message|Represents the message of the prompt.|1.8|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > showPrompt|It determines showing the prompt or not when user selects a cell with the data validation.|1.8|
|[dataValidationPrompt](reference/excel/datavalidationprompt.md)|_Property_ > title|Represents the title for the prompt.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > custom|Custom data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > date|Date data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > decimal|Decimal data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > list|List data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > textLength|TextLength data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > time|Time data validation criteria.|1.8|
|[dataValidationRule](reference/excel/datavalidationrule.md)|_Relationship_ > wholeNumber|WholeNumber data validation criteria.|1.8|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > formula1|Gets or sets the Formula1, i.e. minimum value or value depending of the operator.|1.8|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > formula2|Gets or sets the Formula2, i.e. maximum value or value depending of the operator.|1.8|
|[dateTimeDataValidation](reference/excel/datetimedatavalidation.md)|_Property_ > operator|The operator to use for validating the data. Possible values are: Invalid, Between, NotBetween, EqualTo, NotEqualTo, GreaterThan, LessThan, GreaterThanOrEqualTo, LessThanOrEqualTo.|1.8|
|[listDataValidation](reference/excel/listdatavalidation.md)|_Property_ > inCellDropDown|Displays the list in cell drop down or not, it defaults to true.|1.8|
|[listDataValidation](reference/excel/listdatavalidation.md)|_Property_ > source|Source of the list for data validation|1.8|
|[range](reference/excel/range.md)|_Relationship_ > dataValidation|Returns a data validation object. Read-only.|1.8|

