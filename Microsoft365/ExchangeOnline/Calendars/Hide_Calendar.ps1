<#
By default everyone in the organization is able to see anyone else's calendar, 
due to the "default" level of permission on the Calendar folder set to "AvailabilityOnly". 
Which is basically "read-only" access and everyone should be able to add the Calendar to their Outlook. 
That is, unless you have modified the default permissions.
Now, in most occasions you would want to change the Default level of permissions to LimitedDetails, 
as it will show some additional information about the appointments in the Calendar. You can do this by:
#>
Set-MailboxFolderPermission room:\calendar -User Default -AccessRights LimitedDetails
