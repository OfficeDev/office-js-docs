# WorksheetProtectionOptions Object (JavaScript API for Excel)

Represents the options in sheet protection.

## Properties

| Property              | Type | Description                                                                       | Req. Set |
| :-------------------- | :--- | :-------------------------------------------------------------------------------- | :------- |
| allowAutoFilter       | bool | Represents the worksheet protection option of allowing using auto filter feature. | [1.2]    |
| allowDeleteColumns    | bool | Represents the worksheet protection option of allowing deleting columns.          | [1.2]    |
| allowDeleteRows       | bool | Represents the worksheet protection option of allowing deleting rows.             | [1.2]    |
| allowEditObjects      | bool | Represents the worksheet protection option of allowing editing objects.           | [1.7]    |
| allowEditScenarios    | bool | Represents the worksheet protection option of allowing editing scenarios.         | [1.7]    |
| allowFormatCells      | bool | Represents the worksheet protection option of allowing formatting cells.          | [1.2]    |
| allowFormatColumns    | bool | Represents the worksheet protection option of allowing formatting columns.        | [1.2]    |
| allowFormatRows       | bool | Represents the worksheet protection option of allowing formatting rows.           | [1.2]    |
| allowInsertColumns    | bool | Represents the worksheet protection option of allowing inserting columns.         | [1.2]    |
| allowInsertHyperlinks | bool | Represents the worksheet protection option of allowing inserting hyperlinks.      | [1.2]    |
| allowInsertRows       | bool | Represents the worksheet protection option of allowing inserting rows.            | [1.2]    |
| allowPivotTables      | bool | Represents the worksheet protection option of allowing using PivotTable feature.  | [1.2]    |
| allowSort             | bool | Represents the worksheet protection option of allowing using sort feature.        | [1.2]    |

_See property access [examples.](#property-access-examples)_

## Relationships

| Relationship  | Type                                                  | Description                                                                                           | Req. Set |
| :------------ | :---------------------------------------------------- | :---------------------------------------------------------------------------------------------------- | :------- |
| selectionMode | [ProtectionSelectionMode](protectionselectionmode.md) | Represents the worksheet protection option of selection mode. Possible values: Normal, Unlocked. None | [1.7]    |

## Methods

None

[1.2]: ../requirement-sets/excel-api-requirement-sets.md
[1.7]: ../requirement-sets/excel-api-requirement-sets.md
