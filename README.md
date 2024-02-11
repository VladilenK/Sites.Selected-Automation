# Sites.Selected-Automation

This is to support a blog article and provide some ideas on Sites.Selected SharePoint and graph API permissions automated provisioning.

## Scenario
You administer Microsoft 365 SharePoint Online. Part of your daily activities is providing Microsoft Graph and SharePoint Sites.Selected API permissions to other users (developers).

In Aug/Sep 2023 Microsoft pushed an update that prevents site collection admins to create or update an Azure Access Control (ACS) principal (that was the way most of developers used to get Client Id and Client secret to access SharePoint site). So your users are probably getting something like Your SharePoint tenant admin doesn’t allow site collection admins to create or update an Azure Access Control (ACS) principal message attempting to create or update SharePoint App-only principal at AppRegNew.aspx or AppInv.aspx pages. Here are more details on the issue.

Microsoft and MVPs shared some technique how to provide Sites.Selected API permissions, but dealing with scripts manually, elevating individual permissions every time you need to run the script – it all takes time and not very efficient. More and more devs are reaching you on the app. So you want to automate this process.

## Solution
### Solution architecture
My way to automate it includes:

- SharePoint list as a frontend
here you can accept intake requests, organize approval workflow and display automation results
- Azure Function App as a backend
here will be your PowerShell script hosted that runs on scheduled basis and takes care of actual permissions provisioning



Blog article: [Sites.Selected permissions provisioning automation](https://vladilen.com/office-365/automating-sites-selected-permissions-provisioning/)

Video part 1: [SharePoint Sites.Selected Permissions - Requesting](https://www.youtube.com/watch?v=RyYOeKnR7f0&t=2s)
Video part 2: [SharePoint Sites.Selected Permissions - Automation](https://www.youtube.com/watch?v=n5Q93c82xyA)
