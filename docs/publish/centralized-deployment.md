# Publish Office Add-ins using Centralized Deployment via the Office 365 admin center

The Office 365 admin center makes it easy for an administrator to deploy Office Add-ins to users and groups within their organization. Add-ins deployed via the admin center are available to users in their Office applications right away, with no client configuration required. You can use Centralized Deployment to deploy internal add-ins as well as add-ins provided by ISVs.

The Office 365 admin center currently supports the following scenarios:

- Centralized Deployment of new and updated add-ins to individuals, groups, or an organization.
- Deployment to multiple platforms, including Windows and Office Online, with Mac coming soon.
- Deployment to English language and worldwide tenants.
- Deployment of cloud-hosted add-ins.
- Deployment of add-ins that are hosted within a firewall.
- Deployment of Office Store add-ins.
- Automatic installation of an add-in for users when they launch the Office application.
- Automatic removal of an add-in for users if the admin turns off or deletes the add-in, or if users are removed from Azure Active Directory or from a group to which the add-in has been deployed.

Centralized Deployment is the recommended way for an Office 365 admin to deploy Office add-ins within an organization, provided that the organization meets all requirements for using Centralized Deployment. For information on how to determine if your organization can use Centralized Deployment, see [Determine if Centralized Deployment of add-ins works for your Office 365 organization](https://support.office.com/en-us/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-B4527D49-4073-4B43-8274-31B7A3166F92).

>**Note:** In an on-premises environment with no connection to Office 365, or to deploy SharePoint add-ins or Office Add-ins that target Office 2013, use a [SharePoint add-in catalog](publish-task-pane-and-content-add-ins-to-an-add-in-catalog.md). To deploy COM/VSTO add-ins, use ClickOnce or Windows Installer, as described in [Deploying an Office solution](https://msdn.microsoft.com/en-us/library/bb386179.aspx).

## Publish an Office Add-in via Centralized Deployment

Before you begin, confirm that your organization meets all requirements for using Centralized Deployment, as described in [Determine if Centralized Deployment of add-ins works for your Office 365 organization](https://support.office.com/en-us/article/Determine-if-Centralized-Deployment-of-add-ins-works-for-your-Office-365-organization-B4527D49-4073-4B43-8274-31B7A3166F92).

If your organization meets all requirements, complete the following steps to publish an Office Add-in via Centralized Deployment:

1. Sign in to Office 365 with your work or school account.
2. Select the app launcher icon in the upper-left and choose **Admin**.
3. In the navigation menu, choose **Settings** > **Services & add-ins**.
4. If you see a message on the top of the page announcing the new Office 365 admin center, choose the message to go to the Admin Center Preview (see [About the Office 365 admin center](https://support.office.com/en-ie/article/About-the-Office-365-admin-center-758befc4-0888-4009-9f14-0d147402fd23)).
5. Choose **Upload Add-in** at the top of the page. 
6. Choose one of the following options on the **Centralized Deployment** page:

    - **I want to add an Add-in from the Office Store.**
    - **I have the manifest file (.xml) on this device.** For this option, choose **Browse** to locate the manifest file (.xml) that you want to use.
    - **I have a URL for the manifest file.** For this option, type the manifest's URL in the field provided.

    ![New Add-In dialog in Office 365 admin center](../../images/b3abd42f-63d8-4a5f-8893-d1ae38f4e9b2.png)

7.	Choose **Next**.

8.	If you selected the option to add an Add-in from the Office Store, select the add-in. Notice that you can view available add-ins via categories of **Suggested for you**, **Rating**, or **Name**. You may only add free add-ins from the Office Store; adding paid add-ins isn't currently supported.

    ![Select an Add-In dialog in Office 365 admin center](../../images/2a8de1f4-03b0-4ab6-aa99-4451ee30a64c.png)

9. The add-in is now enabled. On the page for the add-in, its status is **On**, like that shown for the Power BI Tiles add-in in the screenshot below. In the **Who has access** section, choose **Edit** to assign the add-in to users and/or groups.

    ![Power BI Tiles add-in page in Office 365 admin center](../../images/0faa60e8-1e71-4ed1-bbc1-5a2f85ebf981.png)

10.	On the **Edit who has access page**, choose either **Everyone** or **Specific Users/Groups**. Use the Search box to find the users and/or groups to whom you want to deploy the add-in.

    ![Edit who has access page in Office 365 admin center](../../images/46571963-5938-4c7d-b60e-a3ad06758ddf.png)

    >**Note:** For single sign-on (SSO) add-ins, the users and groups assigned will also be shared with add-ins that share the same Azure App ID. Any changes to user assignments will also apply to those add-ins. The related add-ins will be shown on this page. For SSO add-ins only, this page will display the list of Microsoft Graph permissions that the add-in requires.

11.	When finished, choose **Save**, review the add-in settings, and then choose **Close**. You now see your add-in along with other apps in Office 365.
    >**Note:** When an administrator chooses **Save**, consent is given for all users. 

    ![list of apps in Office 365 admin center](../../images/71bfd837-20bc-4517-9513-33fc70147669.png)

>**Tip:** When you deploy a new add-in to users and/or groups in your organization, consider sending them an email that describes when and how to use the add-in, and includes links to relevant Help content, FAQs, or other support resources.

## Considerations when granting access to an add-in

Admins can assign an add-in to everyone in the organization or to specific users and/or groups within the organization. The following list describes the implications of each option:

**Everyone**: As the name implies, this option assigns the add-in to every user in the tenant. Use this option sparingly and only for add-ins that are truly universal to your organization.

**Users**: If you assign an add-in to individual users, you'll need to update the Central Deployment settings for the add-in each time you want to assign it additional users. Likewise, you'll need to update the Central Deployment settings for the add-in each time you want to remove a user's access to the add-in.

**Groups**: If you assign an add-in to a group, users who are added to the group will automatically be assigned the add-in. Likewise, when a user is removed from a group, the user automatically loses access to the add-in. In either case, no additional action is required from the Office 365 admin.

In general, for ease of maintenance, we recommend assigning add-ins by using groups whenever possible. However, in situations where you want to restrict add-in access to a very small number of users, it may be practical enough to assign the add-in to specific users. 

## Add-in states

The following table describes the differnt states of an add-in.

|State|How the state occurs|Impact|
|-----|--------------------|------|
|**Active**|Admin uploaded the add-in and assigned it to users and/or groups.|Users and/or groups assigned to the add-in see it in the relevant Office clients.|
|**Turned off**|Admin turned off the add-in.|Users and/or groups assigned to the add-in no longer have access to it. If the add-in state is changed from **Turned off** to **Active**, the users and groups will have access to it again.|
|**Deleted**|Admin deleted the add-in.|Users and/or groups assigned the add-in no longer have access to it.|

## End user experience with add-ins

If the add-in supports add-in commands, the commands will appear on the Office application ribbon for all users to whom the add-in is deployed. 

If the add-in does not support add-in commands, users can add it from the **My Add-ins** button by doing the following:

1.	In Word 2016, Excel 2016, or PowerPoint 2016, choose **Insert** > **My Add-ins**.
2.	Choose the **Admin Managed** tab in the add-in window.
3.	Choose the add-in, and then choose **Add**. 

