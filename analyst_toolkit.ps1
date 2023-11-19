# Caution: This script is a homegrown tool that is meant to be used only by the author (me). 
# The author is not responsible for any unintended results by anyone else running this code. 
# That being said please tell me where this can be improved and I will surely respond approperatly. 

# Abstract: Running this script will give you a list of options to choose from. Select the approperate
# integer for the respective task you want performed. 


function Get-NullOrNahh 
{
    param([Array] $arr, [string] $name)

    if ($arr -ne $null)
    {
        $myType = ($arr.GetType()).BaseType.Name
        if ($myType -eq "Array" )
        {
            if ($arr.Count -ne 0 -and $arr.Length -ne 0)
            {
                Write-Host -ForegroundColor DarkGreen "Found items for the search named $($name)`n"
                return $true
            }
            else { Write-Host -ForegroundColor Red "Array is empty for the search named $($name)`n"; return $false }
        } else { Write-Host -ForegroundColor Red "Not an array for the search named $($name)`n"; return $false }
    }
}

try
{
    Connect-AzureAD

    $users = Get-AzureADUser
    foreach ($user in $users) 
    { 
        # Path variavle definitions - Azure
        $outputAzMemPath = "C:\Users\$($env:USERNAME)\Downloads\$($user.UserPrincipalName)-az_memberships.csv"
        $outputAzOwnedObjectPath = "C:\Users\$($env:USERNAME)\Downloads\$($user.UserPrincipalName)-az_owned_obj.csv"
        $outputAzLicensePath = "C:\Users\$($env:USERNAME)\Downloads\$($user.UserPrincipalName)-az_license.csv"
        $outputAzAuditPath = "C:\Users\$($env:USERNAME)\Downloads\$($user.UserPrincipalName)-az_audit.csv"
        $outputAzLoginPath = "C:\Users\$($env:USERNAME)\Downloads\$($user.UserPrincipalName)-az_signin.csv"

        Write-Host -ForegroundColor Yellow "`n`nMemberships for $($user.DisplayName)"  
        $azMems = Get-AzureADUserMembership -All $true -ObjectId $user.ObjectId 
        if (Get-NullOrNahh $azMems "Azure Memberships") { $azMems | Export-Csv -Path $outputAzMemPath -NoTypeInformation }
        
        Write-Host -ForegroundColor Yellow "Owned objects for $($user.DisplayName)"
        $azOwnObjects = Get-AzureADUserOwnedObject -All $true -ObjectId $user.ObjectId 
        if (Get-NullOrNahh $azOwnObjects "Owned Objects") { $azOwnObjects | Export-Csv -Path $outputAzOwnedObjectPath -NoTypeInformation }

        Write-Host -ForegroundColor Yellow "Licenses for $($user.DisplayName)"
        $azLicenses = Get-AzureADUserLicenseDetail -ObjectId $user.ObjectId 
        if (Get-NullOrNahh $azLicenses "Licenses") { $azLicenses | Select-Object SkuPartNumber | Export-Csv -Path $outputAzLicensePath -NoTypeInformation }

        Write-Host -ForegroundColor Yellow "audit log entries for $($user.DisplayName)"
        $azAuditLog = Get-AzureADAuditDirectoryLogs -Filter "initiatedBy/user/UserPrincipalName eq '$($user.UserPrincipalName)'" 
        if (Get-NullOrNahh $azAuditLog "Audit Log") { $azAuditLog | Export-Csv -Path $outputAzAuditPath -NoTypeInformation }
        
        Write-Host -ForegroundColor Yellow "Logins for $($user.DisplayName)"
        $azSignInLog = Get-AzureADAuditSignInLogs -Filter "userPrincipalName eq '$($user.UserPrincipalName)'" 
        if (Get-NullOrNahh $azSignInLog "Sign-In Log") { $azSignInLog | Export-Csv -Path $outputAzLoginPath -NoTypeInformation }
    }
}
catch
{
    Write-Host $_
}
finally
{
    Disconnect-AzureAD
}



try 
{
    Connect-ExchangeOnline

    $mailboxes = Get-Mailbox
    foreach ($mailbox in $mailboxes)
    {
        # Path variavle definitions - Azure
        $outputMailRecipientPermissionsTrustee = "C:\Users\$($env:USERNAME)\Downloads\$($mailbox.UserPrincipalName)-mail_recipient_permissions_trustee.csv"
        $outputMailRecipientPermissions = "C:\Users\$($env:USERNAME)\Downloads\$($mailbox.UserPrincipalName)-mail_recipient_permissions.csv"
        $outputMailDelegates = "C:\Users\$($env:USERNAME)\Downloads\$($mailbox.UserPrincipalName)-mail_delegates.csv"
        $outputMailInboxRules = "C:\Users\$($env:USERNAME)\Downloads\$($mailbox.UserPrincipalName)-mail_inbox_rules.csv"
        $outputMailStats = "C:\Users\$($env:USERNAME)\Downloads\$($mailbox.UserPrincipalName)-mail_stats.csv"
        $outputMailAudit = "C:\Users\$($env:USERNAME)\Downloads\$($mailbox.UserPrincipalName)-mail_audit.csv"

        # lists the recipients for whom the user has SendAs permission. The user can send messages that appear to come directly from the recipients.
        Write-Host -ForegroundColor Cyan "`n`nDelegate access on other mailboxes for $($mailbox.DisplayName)";
        $mailboxRecipientPermissionsTrustee = Get-RecipientPermission -Trustee $mailbox.Identity 
        if (Get-NullOrNahh $mailboxRecipientPermissionsTrustee "Recipient Permissions Trustee") { $mailboxRecipientPermissionsTrustee | Export-Csv -Path $outputMailRecipientPermissionsTrustee -NoTypeInformation }

        # This example lists the users who have SendAs permission on the mailbox. The users listed can send messages that appear to come directly from the mailbox.
        Write-Host -ForegroundColor Cyan "Delegate send-as access on mailbox for $($mailbox.DisplayName)";
        $mailboxRecipientPermissions = Get-RecipientPermission $mailbox.Identity 
        if (Get-NullOrNahh $mailboxRecipientPermissions "Recipient Permissions") { $mailboxRecipientPermissions | Export-Csv -Path $outputMailRecipientPermissions -NoTypeInformation }

        # Use the Get-MailboxPermission cmdlet to retrieve permissions on a mailbox.
        Write-Host -ForegroundColor Cyan "Delegate access on $($mailbox.DisplayName)"; 
        $mailboxPermissions = Get-MailboxPermission -Identity $mailbox.Identity 
        if (Get-NullOrNahh $mailboxPermissions "Mailbox Permissions") { $mailboxPermissions | Export-Csv -Path $outputMailDelegates -NoTypeInformation }

        # Retrieves mailbox audit log entries for actions performed by Admin, Owner, and Delegate login types. A maximum of 2,000 log entries are returned
        Write-Host -ForegroundColor Cyan "Mailboxes audit logs for $($mailbox.DisplayName)"; 
        $mailboxAudit = Search-MailboxAuditLog -ShowDetails -Identity $mailbox.UserPrincipalName -LogonTypes Admin,Delegate,Owner -StartDate 11/01/2023 -EndDate 11/19/2023 -ResultSize 5000
        if (Get-NullOrNahh $mailboxAudit "Mailbox Audit Log") { $mailboxAudit | Export-Csv -Path $outputMailAudit -NoTypeInformation }

        # Retrieves all Inbox rules for the mailbox
        Write-Host -ForegroundColor Cyan "Mailboxes inbox rules for $($mailbox.DisplayName)"; 
        $mailboxInboxRules = Get-InboxRule -Mailbox $mailbox.UserPrincipalName -IncludeHidden 
        if (Get-NullOrNahh $mailboxInboxRules "Inbox Rules") { $mailboxInboxRules | Export-Csv -Path $outputMailInboxRules -NoTypeInformation }
        
        # Retrieves all the information about the user
        Write-Host -ForegroundColor Cyan "Mailboxes statistics for $($mailbox.DisplayName)"; 
        $mailboxStats = Get-MailboxFolderStatistics -Identity $mailbox.Identity 
        if (Get-NullOrNahh $mailboxStats "Sign-In Log") { $mailboxStats | Export-Csv -Path $outputMailStats -NoTypeInformation }
    }
}
catch 
{
    Write-Host $_
}
finally 
{
    Disconnect-ExchangeOnline
}