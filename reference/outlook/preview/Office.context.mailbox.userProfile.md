

# userProfile

### [Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

### Members

####  accountType :String

> **Note:** This member is currently only supported in Outlook 2016 for Mac, build # and greater.

Gets the account type of the user associated with the mailbox. The possible values are listed in the following table.

| Value | Description |
|-------|-------------|
| `enterprise` | The mailbox is on an on-premises Exchange server. |
| `gmail` | The mailbox is associated with a Gmail account. |
| `office365` | The mailbox is associated with an Office 365 work or school account. |
| `outlookCom` | The mailbox is associated with a personal Outlook.com account. |

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| preview|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
console.log(Office.context.mailbox.userProfile.accontType);
```

####  displayName :String

Gets the user's display name.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  emailAddress :String

Gets the user's SMTP email address.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  timeZone :String

Gets the user's default time zone.

##### Type:

*   String

##### Requirements

|Requirement| Value|
|---|---|
|[Minimum mailbox requirement set version](../tutorial-api-requirement-sets.md)| 1.0|
|[Minimum permission level](../../../docs/outlook/understanding-outlook-add-in-permissions.md)| ReadItem|
|Applicable Outlook mode| Compose or read|

##### Example

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```