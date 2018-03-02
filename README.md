# Rapid7AlertsChecker
C# application that reads emails from a mailbox in Office 365, parses them and sends them to Slack as notifications.

Steps needed to set up Azure AD Application so it can read specified mailbox securely.

1. Go to Azure AD Admin Center - https://aad.portal.azure.com
2. Open App Registrations
3. Click New Application Registration
4. Enter in the Name, select 'Native' as Application Type and type https://localhost for Redirect URI
5. Click on the new app and select All Settings
6. Click Required Permissions
7. Remove any preconfigured APIs you do not need access to (such as Windows Azure Active Directory)
8. Click Add
9. Select the API scopes you wish to grant the application and save (in this case grant MS Graph API access with Read User mail and Read and Write User Mail delegate access)
10. Click Grant Permissions along the top
11. Click Enterprise Applications along the left
12. Select All Applications
13. Find the new application from the list and select it
14. Click Properties
15. After everything loads switch User Assignment Required? To Yes and Visible To Users? to No
16. Create a new user account in Azure AD that will gain access to the app if one does not already exist. Ensure that the account has licensing in Office 365 so that a mailbox is generated for it. 
17. Open PowerShell
18. Import the AzureAD module (https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0)
19. Connect to AzureAD using the Connect-AzureAD cmdlet (https://docs.microsoft.com/en-us/powershell/azure/active-directory/install-adv2?view=azureadps-2.0)
20. Run the following commands (modify to what you need)
	$user = Get-AzureADUser -ObjectId username@domain.com
	$sp = Get-AzureADServicePrincipal -Filter "displayname eq '{New App's Name Goes Here}'"
	New-AzureADUserAppRoleAssignment -ObjectId $user.ObjectId -PrincipalId $user.ObjectId -ResourceId $sp.ObjectId -Id ([Guid]::Empty)
21. Verify in the web GUI that the user shows up in the applications Overview as a tile
22. Go back to the new Azure AD App and copy the Application ID (aka Client ID) and the Tenant ID (can be found in the endpoints list or in the tenant properties)

References:
https://docs.microsoft.com/en-us/azure/active-directory/active-directory-coreapps-assign-user-azure-portal
https://serverfault.com/questions/845293/how-to-assign-users-to-a-native-app-in-azure-ad
https://docs.microsoft.com/en-us/azure/active-directory/application-access-assignment-how-to-add-assignment#assign-a-user-directly-as-an-administrator
	
