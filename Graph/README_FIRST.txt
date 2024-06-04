**************************************************
* Graph API - obtaining Client ID and Secret Key *
**************************************************

1. Login to Microsoft Azure console
Navigate to Azure Active Directory* and click on App Registration from the left 
panel (the first item under management section).

(*) alternatively navigate to this URL, redirecting to Azure AD B2C:
https://portal.azure.com/#view/Microsoft_AAD_B2CAdmin/TenantManagementMenuBlade/~/overview


2. Register a new App
Once you clicked on "New Registration" button (from the header of the App 
Registration tab) you can get a name to your app. You should restrict it for 
internal accounts only (the option described as "Single Tenant").

3. Create a Secret Key
From the overview of the new app created you can create a secret by clicking on 
the link "Add a Certificate or a Secret"

4. Check authorizations and scopes
Please check the right permission(s) before integrates the app credentials into 
a new script. An overview of them can be obtained from the Graph Explorer 
platform, clicking on "Modify permissions" tab:
https://developer.microsoft.com/en-us/graph/graph-explorer
