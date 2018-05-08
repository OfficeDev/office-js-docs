 

# MailboxEnums

### [Office](Office.md). MailboxEnums

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|1.0|
|Applicable Outlook mode|Compose or read|

### Members

#### AttachmentType :String

Specifies an attachment's type.

AttachmentType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`File`| String|`file`|The attachment is a file.|
|`Item`| String|`item`|The attachment is an Exchange item.|
|`Cloud`| String|`cloud`|The attachment is stored in a cloud location, such as OneDrive. The `id` property of the attachment contains a URL to the file.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|1.0|
|Applicable Outlook mode|Compose or read|


#### Days :String

Specifies the day of week or type of day.

Days

##### Type:

* String

##### Properties:

| Name | Type | Value | Description |
|---|---|---|---|
|`Mon`| String |`mon`|Monday.|
|`Tue`| String |`tue`|Tuesday.|
|`Wed`| String |`wed`|Wednesday.|
|`Thu`| String |`thu`|Thursday.|
|`Fri`| String |`fri`|Friday.|
|`Sat`| String |`sat`|Saturday.|
|`Sun`| String |`sun`|Sunday.|
|`Weekday`| String |`weekday`|Week day (excludes weekend days): `Mon`, `Tue`, `Wed`, `Thu`, and `Fri`.|
|`WeekendDay`| String |`weekendDay`|Weekend day: `Sat` and `Sun`.|
|`Day`| String |`day`|Day of week.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|Applicable Outlook mode|Compose or read|


#### EntityType :String

Specifies an entity's type.

EntityType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`Address`| String|`address`|Specifies that the entity is a postal address.|
|`Contact`| String|`contact`|Specifies that the entity is a contact.|
|`EmailAddress`| String|`emailAddress`|Specifies that the entity is SMTP email address.|
|`MeetingSuggestion`| String|`meetingSuggestion`|Specifies that the entity is a meeting suggestion.|
|`PhoneNumber`| String|`phoneNumber`|Specifies that the entity is US phone number.|
|`TaskSuggestion`| String|`taskSuggestion`|Specifies that the entity is a task suggestion.|
|`URL`| String|`url`|Specifies that the entity is an Internet URL.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|1.0|
|Applicable Outlook mode|Compose or read|


#### ItemNotificationMessageType :String

Specifies the notification message type for an appointment or message.

ItemNotificationMessageType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`ProgressIndicator`| String|`progressIndicator`|The notificationMessage is a progress indicator.|
|`InformationalMessage`| String|`informationalMessage`|The notificationMessage is an informational message.|
|`ErrorMessage`| String|`errorMessage`|The notificationMessage is an error message.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|1.3|
|Applicable Outlook mode|Compose or read|


#### ItemType :String

Specifies an item's type.

ItemType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`Message`| String|`message`|An email, meeting request, meeting response, or meeting cancellation.|
|`Appointment`| String|`appointment`|An appointment item.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|1.0|
|Applicable Outlook mode|Compose or read|


#### Month :String

Specifies the month.

Month

##### Type:

* String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`Jan`| String |`jan`|January or 1st month of the year.|
|`Feb`| String |`feb`|February or 2nd month of the year.|
|`Mar`| String |`mar`|March or 3rd month of the year.|
|`Apr`| String |`apr`|April or 4th month of the year.|
|`May`| String |`may`|May or 5th month of the year.|
|`Jun`| String |`jun`|June or 6th month of the year.|
|`Jul`| String |`jul`|July or 7th month of the year.|
|`Aug`| String |`aug`|August or 8th month of the year.|
|`Sep`| String |`sep`|September or 9th month of the year.|
|`Oct`| String |`oct`|October or 10th month of the year.|
|`Nov`| String |`nov`|November or 11th month of the year.|
|`Dec`| String |`dec`|December or 12th month of the year.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|Applicable Outlook mode|Compose or read|


#### RecipientType :String

Specifies the type of recipient for an appointment.

RecipientType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`Other`| String|`other`|The recipient is not one of the other recipient types.|
|`DistributionList`| String|`distributionList`|The recipient is a distribution list containing a list of email addresses.|
|`User`| String|`user`|The recipient is an SMTP email address that is on the Exchange server.|
|`ExternalUser`| String|`externalUser`|The recipient is an SMTP email address that is not on the Exchange server.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|1.1|
|Applicable Outlook mode|Compose or read|


#### RecurrenceTimeZone :String

Specifies the time zone applied to the recurrence.

RecurrenceTimeZone

##### Type:

* String

##### Properties:

| Name | Type | Value | Description |
|---|---|---|---|
|`AfghanistanStandardTime`|String|`Afghanistan Standard Time`|Afghanistan Standard Time|
|`AlaskanStandardTime`|String|`Alaskan Standard Time`|Alaskan Standard Time|
|`AleutianStandardTime`|String|`Aleutian Standard Time`|Aleutian Standard Time|
|`AltaiStandardTime`|String|`Altai Standard Time`|Altai Standard Time|
|`ArabStandardTime`|String|`Arab Standard Time`|Arab Standard Time|
|`ArabianStandardTime`|String|`Arabian Standard Time`|Arabian Standard Time|
|`ArabicStandardTime`|String|`Arabic Standard Time`|Arabic Standard Time|
|`ArgentinaStandardTime`|String|`Argentina Standard Time`|Argentina Standard Time|
|`AstrakhanStandardTime`|String|`Astrakhan Standard Time`|Astrakhan Standard Time|
|`AtlanticStandardTime`|String|`Atlantic Standard Time`|Atlantic Standard Time|
|`AUSCentralStandardTime`|String|`AUS Central Standard Time`|Australia Central Standard Time|
|`AusCentralW_StandardTime`|String|`Aus Central W. Standard Time`|Australia Central West Standard Time|
|`AUSEasternStandardTime`|String|`AUS Eastern Standard Time`|AUS Eastern Standard Time|
|`AzerbaijanStandardTime`|String|`Azerbaijan Standard Time`|Azerbaijan Standard Time|
|`AzoresStandardTime`|String|`Azores Standard Time`|Azores Standard Time|
|`BahiaStandardTime`|String|`Bahia Standard Time`|Bahia Standard Time|
|`BangladeshStandardTime`|String|`Bangladesh Standard Time`|Bangladesh Standard Time|
|`BelarusStandardTime`|String|`Belarus Standard Time`|Belarus Standard Time|
|`BougainvilleStandardTime`|String|`Bougainville Standard Time`|Bougainville Standard Time|
|`CanadaCentralStandardTime`|String|`Canada Central Standard Time`|Canada Central Standard Time|
|`CapeVerdeStandardTime`|String|`Cape Verde Standard Time`|Cape Verde Standard Time|
|`CaucasusStandardTime`|String|`Caucasus Standard Time`|Caucasus Standard Time|
|`CenAustraliaStandardTime`|String|`Cen. Australia Standard Time`|Central Australia Standard Time|
|`CentralAmericaStandardTime`|String|`Central America Standard Time`|Central America Standard Time|
|`CentralAsiaStandardTime`|String|`Central Asia Standard Time`|Central Asia Standard Time|
|`CentralBrazilianStandardTime`|String|`Central Brazilian Standard Time`|Central Brazilian Standard Time|
|`CentralEuropeStandardTime`|String|`Central Europe Standard Time`|Central Europe Standard Time|
|`CentralEuropeanStandardTime`|String|`Central European Standard Time`|Central European Standard Time|
|`CentralPacificStandardTime`|String|`Central Pacific Standard Time`|Central Pacific Standard Time|
|`CentralStandardTime`|String|`Central Standard Time`|Central Standard Time|
|`CentralStandardTime_Mexico`|String|`Central Standard Time (Mexico)`|Central Standard Time (Mexico)|
|`ChathamIslandsStandardTime`|String|`Chatham Islands Standard Time`|Chatham Islands Standard Time|
|`ChinaStandardTime`|String|`China Standard Time`|China Standard Time|
|`CubaStandardTime`|String|`Cuba Standard Time`|Cuba Standard Time|
|`DatelineStandardTime`|String|`Dateline Standard Time`|Dateline Standard Time|
|`E_AfricaStandardTime`|String|`E. Africa Standard Time`|East Africa Standard Time|
|`E_AustraliaStandardTime`|String|`E. Australia Standard Time`|East Australia Standard Time|
|`E_EuropeStandardTime`|String|`E. Europe Standard Time`|East Europe Standard Time|
|`E_SouthAmericaStandardTime`|String|`E. South America Standard Time`|East South America Standard Time|
|`EasterIslandStandardTime`|String|`Easter Island Standard Time`|Easter Island Standard Time|
|`EasternStandardTime`|String|`Eastern Standard Time`|Eastern Standard Time|
|`EasternStandardTime_Mexico`|String|`Eastern Standard Time (Mexico)`|Eastern Standard Time (Mexico)|
|`EgyptStandardTime`|String|`Egypt Standard Time`|Egypt Standard Time|
|`EkaterinburgStandardTime`|String|`Ekaterinburg Standard Time`|Ekaterinburg Standard Time|
|`FijiStandardTime`|String|`Fiji Standard Time`|Fiji Standard Time|
|`FLEStandardTime`|String|`FLE Standard Time`|FLE Standard Time|
|`GeorgianStandardTime`|String|`Georgian Standard Time`|Georgian Standard Time|
|`GMTStandardTime`|String|`GMT Standard Time`|GMT Standard Time|
|`GreenlandStandardTime`|String|`Greenland Standard Time`|Greenland Standard Time|
|`GreenwichStandardTime`|String|`Greenwich Standard Time`|Greenwich Standard Time|
|`GTBStandardTime`|String|`GTB Standard Time`|GTB Standard Time|
|`HaitiStandardTime`|String|`Haiti Standard Time`|Haiti Standard Time|
|`HawaiianStandardTime`|String|`Hawaiian Standard Time`|Hawaiian Standard Time|
|`IndiaStandardTime`|String|`India Standard Time`|India Standard Time|
|`IranStandardTime`|String|`Iran Standard Time`|Iran Standard Time|
|`IsraelStandardTime`|String|`Israel Standard Time`|Israel Standard Time|
|`JordanStandardTime`|String|`Jordan Standard Time`|Jordan Standard Time|
|`KaliningradStandardTime`|String|`Kaliningrad Standard Time`|Kaliningrad Standard Time|
|`KamchatkaStandardTime`|String|`Kamchatka Standard Time`|Kamchatka Standard Time|
|`KoreaStandardTime`|String|`Korea Standard Time`|Korea Standard Time|
|`LibyaStandardTime`|String|`Libya Standard Time`|Libya Standard Time|
|`LineIslandsStandardTime`|String|`Line Islands Standard Time`|Line Islands Standard Time|
|`LordHoweStandardTime`|String|`Lord Howe Standard Time`|Lord Howe Standard Time|
|`MagadanStandardTime`|String|`Magadan Standard Time`|Magadan Standard Time|
|`MagallanesStandardTime`|String|`Magallanes Standard Time`|Magallanes Standard Time|
|`MarquesasStandardTime`|String|`Marquesas Standard Time`|Marquesas Standard Time|
|`MauritiusStandardTime`|String|`Mauritius Standard Time`|Mauritius Standard Time|
|`MidAtlanticStandardTime`|String|`Mid-Atlantic Standard Time`|Mid-Atlantic Standard Time|
|`MiddleEastStandardTime`|String|`Middle East Standard Time`|Middle East Standard Time|
|`MontevideoStandardTime`|String|`Montevideo Standard Time`|Montevideo Standard Time|
|`MoroccoStandardTime`|String|`Morocco Standard Time`|Morocco Standard Time|
|`MountainStandardTime`|String|`Mountain Standard Time`|Mountain Standard Time|
|`MountainStandardTime_Mexico`|String|`Mountain Standard Time (Mexico)`|Mountain Standard Time (Mexico)|
|`MyanmarStandardTime`|String|`Myanmar Standard Time`|Myanmar Standard Time|
|`N_CentralAsiaStandardTime`|String|`N. Central Asia Standard Time`|North Central Asia Standard Time|
|`NamibiaStandardTime`|String|`Namibia Standard Time`|Namibia Standard Time|
|`NepalStandardTime`|String|`Nepal Standard Time`|Nepal Standard Time|
|`NewZealandStandardTime`|String|`New Zealand Standard Time`|New Zealand Standard Time|
|`NewfoundlandStandardTime`|String|`Newfoundland Standard Time`|Newfoundland Standard Time|
|`NorfolkStandardTime`|String|`Norfolk Standard Time`|Norfolk Standard Time|
|`NorthAsiaEastStandardTime`|String|`North Asia East Standard Time`|North Asia East Standard Time|
|`NorthAsiaStandardTime`|String|`North Asia Standard Time`|North Asia Standard Time|
|`NorthKoreaStandardTime`|String|`North Korea Standard Time`|North Korea Standard Time|
|`OmskStandardTime`|String|`Omsk Standard Time`|Omsk Standard Time|
|`PacificSAStandardTime`|String|`Pacific SA Standard Time`|Pacific SA Standard Time|
|`PacificStandardTime`|String|`Pacific Standard Time`|Pacific Standard Time|
|`PacificStandardTimeMexico`|String|`Pacific Standard Time (Mexico)`|Pacific Standard Time (Mexico)|
|`PakistanStandardTime`|String|`Pakistan Standard Time`|Pakistan Standard Time|
|`ParaguayStandardTime`|String|`Paraguay Standard Time`|Paraguay Standard Time|
|`RomanceStandardTime`|String|`Romance Standard Time`|Romance Standard Time|
|`RussiaTimeZone10`|String|`Russia Time Zone 10`|Russia Time Zone 10|
|`RussiaTimeZone11`|String|`Russia Time Zone 11`|Russia Time Zone 11|
|`RussiaTimeZone3`|String|`Russia Time Zone 3`|Russia Time Zone 3|
|`RussianStandardTime`|String|`Russian Standard Time`|Russian Standard Time|
|`SAEasternStandardTime`|String|`SA Eastern Standard Time`|SA Eastern Standard Time|
|`SAPacificStandardTime`|String|`SA Pacific Standard Time`|SA Pacific Standard Time|
|`SAWesternStandardTime`|String|`SA Western Standard Time`|SA Western Standard Time|
|`SaintPierreStandardTime`|String|`Saint Pierre Standard Time`|Saint Pierre Standard Time|
|`SakhalinStandardTime`|String|`Sakhalin Standard Time`|Sakhalin Standard Time|
|`SamoaStandardTime`|String|`Samoa Standard Time`|Samoa Standard Time|
|`SaratovStandardTime`|String|`Saratov Standard Time`|Saratov Standard Time|
|`SEAsiaStandardTime`|String|`SE Asia Standard Time`|Southeast Asia Standard Time|
|`SingaporeStandardTime`|String|`Singapore Standard Time`|Singapore Standard Time|
|`SouthAfricaStandardTime`|String|`South Africa Standard Time`|South Africa Standard Time|
|`SriLankaStandardTime`|String|`Sri Lanka Standard Time`|Sri Lanka Standard Time|
|`SudanStandardTime`|String|`Sudan Standard Time`|Sudan Standard Time|
|`SyriaStandardTime`|String|`Syria Standard Time`|Syria Standard Time|
|`TaipeiStandardTime`|String|`Taipei Standard Time`|Taipei Standard Time|
|`TasmaniaStandardTime`|String|`Tasmania Standard Time`|Tasmania Standard Time|
|`TocantinsStandardTime`|String|`Tocantins Standard Time`|Tocantins Standard Time|
|`TokyoStandardTime`|String|`Tokyo Standard Time`|Tokyo Standard Time|
|`TomskStandardTime`|String|`Tomsk Standard Time`|Tomsk Standard Time|
|`TongaStandardTime`|String|`Tonga Standard Time`|Tonga Standard Time|
|`TransbaikalStandardTime`|String|`Transbaikal Standard Time`|Transbaikal Standard Time|
|`TurkeyStandardTime`|String|`Turkey Standard Time`|Turkey Standard Time|
|`TurksAndCaicosStandardTime`|String|`Turks And Caicos Standard Time`|Turks And Caicos Standard Time|
|`UlaanbaatarStandardTime`|String|`Ulaanbaatar Standard Time`|Ulaanbaatar Standard Time|
|`USEasternStandardTime`|String|`US Eastern Standard Time`|United States Eastern Standard Time|
|`USMountainStandardTime`|String|`US Mountain Standard Time`|United States Mountain Standard Time|
|`UTC`|String|`UTC`|Coordinated Universal Time (UTC)|
|`UTCPLUS12`|String|`UTC+12`|Coordinated Universal Time (UTC) + 12 hours|
|`UTCPLUS13`|String|`UTC+13`|Coordinated Universal Time (UTC) + 13 hours|
|`UTCMINUS02`|String|`UTC-02`|Coordinated Universal Time (UTC) - 2 hours|
|`UTCMINUS08`|String|`UTC-08`|Coordinated Universal Time (UTC) - 8 hours|
|`UTCMINUS09`|String|`UTC-09`|Coordinated Universal Time (UTC) - 9 hours|
|`UTCMINUS11`|String|`UTC-11`|Coordinated Universal Time (UTC) - 11 hours|
|`VenezuelaStandardTime`|String|`Venezuela Standard Time`|Venezuela Standard Time|
|`VladivostokStandardTime`|String|`Vladivostok Standard Time`|Vladivostok Standard Time|
|`W_AustraliaStandardTime`|String|`W. Australia Standard Time`|West Australia Standard Time|
|`W_CentralAfricaStandardTime`|String|`W. Central Africa Standard Time`|West Central Africa Standard Time|
|`W_EuropeStandardTime`|String|`W. Europe Standard Time`|West Europe Standard Time|
|`W_MongoliaStandardTime`|String|`W. Mongolia Standard Time`|West Mongolia Standard Time|
|`WestAsiaStandardTime`|String|`West Asia Standard Time`|West Asia Standard Time|
|`WestBankStandardTime`|String|`West Bank Standard Time`|West Bank Standard Time|
|`WestPacificStandardTime`|String|`West Pacific Standard Time`|West Pacific Standard Time|
|`YakutskStandardTime`|String|`Yakutsk Standard Time`|Yakutsk Standard Time|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|Applicable Outlook mode|Compose or read|


#### RecurrenceType :String

Specifies the type of recurrence.

RecurrenceType

##### Type:

* String

##### Properties:

| Name | Type | Value | Description |
|---|---|---|---|
|`Daily`| String |`daily`|Daily.|
|`Weekday`| String |`weekday`|Weekday.|
|`Weekly`| String |`weekly`|Weekly.|
|`Monthly`| String |`monthly`|Monthly.|
|`Yearly`| String |`yearly`|Yearly.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|Applicable Outlook mode|Compose or read|


#### ResponseType :String

Specifies the type of response to a meeting invitation.

ResponseType

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`None`| String|`none`|There has been no response from the attendee.|
|`Organizer`| String|`organizer`|The attendee is the meeting organizer.|
|`Tentative`| String|`tentative`|The meeting request was tentatively accepted by the attendee.|
|`Accepted`| String|`accepted`|The meeting request was accepted by the attendee.|
|`Declined`| String|`declined`|The meeting request was declined by the attendee.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|1.0|
|Applicable Outlook mode|Compose or read|


#### RestVersion :String

Specifies the version of the REST API that corresponds to a REST-formatted item ID. 

RestVersion

##### Type:

*   String

##### Properties:

|Name| Type| Value | Description|
|---|---|---|---|
|`v1_0`| String|`v1.0`|Version 1.0.|
|`v2_0`| String|`v2.0`|Version 2.0.|
|`Beta`| String|`beta`|Beta.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|1.3|
|Applicable Outlook mode|Compose or read|


#### WeekNumber :String

Specifies the week of the month.

WeekNumber

##### Type:

* String

##### Properties:

| Name | Type | Value | Description |
|---|---|---|---|
|`First`| String |`first`|First week of the month.|
|`Second`| String |`second`|Second week of the month.|
|`Third`| String |`third`|Third week of the month.|
|`Fourth`| String |`fourth`|Fourth week of the month.|
|`Fifth`| String |`fifth`|Fifth week of the month.|

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)|Preview|
|Applicable Outlook mode|Compose or read|