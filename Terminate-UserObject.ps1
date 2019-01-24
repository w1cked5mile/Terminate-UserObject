<#
.Synopsis
Terminate-UserObject is used to secure terminated employees user objects.
.Description
Terminate-UserObject was written to accomdate an AD Forest infrastructure and uses Get-ADForest to enumerate the domains available.
It will validate a user for termination from any domain in the forest and accomodates user objects that have a 
manager in another domain.  This script has been tested in an AD Forest with three domains and approximately
900 User Objects. Results may vary with larger scale AD Forests. Script supports MFA or non-MFA MS Online 
and Exchange Online connectivity using the -MFA switch.
Actions:
    Disable user object
    Remove user object from all groups
    Clear user object attributes
    Set user object description to termination date
    Set Out of Office messages
    Set delegate authority on the mailbox to terminated employee's manager
    Migrate regular mailbox to shared mailbox
    Removes Office 365 Licenses
    Hide user object in the GAL
    Move user object to the "Terminated" OU (NOTE: OU must be at the root of the AD domain and named "Terminated")
.Parameter TargetUser
User Logon Name of User-Object to be terminated
.Parameter Delegate
Alternate delegate User-Object for mailbox and out of office replies (Optional)
.Parameter MFA
Required if using multi-factor authentication with MS Online or Exchange Online
.Parameter SkipOOO
Skip processing the Out of Office replies
.Example
Terminate-UserObject -TargetUser Minnie.Mouse
NOTE: Will request credentials twice for login to MS Online and Exchange Online with MFA
Output:
Found Minnie.Mouse on DC1.SUB1.AD.CONTOSO.COM
Ready to terminate:
User:           Minnie Mouse
User Email:     minnie.mouse@contoso.com
User DC:        DC1.SUB1.AD.CONTOSO.COM
User Company:   Contoso
Manger:         Mickey.Mouse
Manager Email:  mickey.mouse@contoso.com
Manager DC:     DC1.SUB1.AD.CONTOSO.COM
---
Disabled Minnie Mouse
Minnie Mouse removed from CN=ACL_Sales_READ,OU=File Share Access,OU=Groups,DC=SUB1,DC=AD,DC=CONTOSO,DC=com
Minnie Mouse removed from CN=ACL_Public_READ,OU=File Share Access,OU=Groups,DC=SUB1,DC=AD,DC=CONTOSO,DC=com
Minnie Mouse removed from all groups.
Cleared attributes for Minnie Mouse.
Minnie Mouse out of office replies set.
---
Mickey.Mouse@contoso.com added as a Full Rights delegate for Minnie Mouse mailbox.
Minnie Mouse mailbox migrated to shared.
Minnie Mouse hidden from Global Address List.
Minnie Mouse moved to OU=Terminated,DC=SUB1,DC=AD,DC=CONTOSO,DC=COM.
Minnie Mouse has been terminated.
.Example
Terminate-UserObject -TargetUser Mickey.Mouse -MFA
NOTE: Will request credentials for logging into MS Online and Exchange Online
#>

#Requires -Modules ActiveDirectory, MSOnline 

Param (
    [Parameter (Mandatory = $true)]
    [string]$TargetUser,
    [Parameter (Mandatory = $False)]
    [string]$Delegate,
    [Parameter (Mandatory = $False)]
    [switch]$SkipOOO,
    [Parameter (Mandatory = $false)]
    [switch]$MFA
)

Function Get-UserDC($Target) {
    $AllDomains = (Get-ADForest).Domains 
    foreach ($Domain in $AllDomains) {
        $TargetDC = (Get-ADDomainController -Server $Domain).Hostname
        Try {
            Get-ADUser -Server $TargetDC -Identity $Target -ErrorAction Stop > $null
            $TargetDC
            Return
        }
        catch {}
    }
}

Function Get-PrimarySMTP ($Target) {
    foreach ($address in $Target.ProxyAddresses) {
        if (($address.Length -gt 5) -and ($address.SubString(0, 5) -ceq 'SMTP:')) {
            $TargetEmail += $address.SubString(5)
            $TargetEmail
            Return
        }
    }
}

Function Disable-User {
    
    # Disable User Object

    if ($User.enabled -eq $true) {
        Disable-ADAccount -Server $TargetDC -identity $User.SAMAccountName
        Write-log -LogEntry "Disabled $($User.DisplayName)" -LogPath $LogPath
    }
    else {
        Write-log -LogEntry "$($User.DisplayName) is already disabled" -LogPath $LogPath
    }
}

Function Clear-UserGroups {
    # Remove User Object from all groups

    if ($null -eq $User.MemberOf) {
        Write-log -LogEntry "$($User.DisplayName) is not a member of any groups" -LogPath $LogPath
    }
    else {
        $Groups = $User.MemberOf
        foreach ($Group in $Groups) {
            $User.MemberOf | Remove-ADGroupMember -Server $TargetDC -confirm:$false -Members $User.SAMAccountName
            Write-log -LogEntry "$($User.DisplayName) removed from $Group" -LogPath $LogPath
        }
        Write-log -LogEntry  "$($User.DisplayName) removed from all groups." -LogPath $LogPath
    }
}

Function Clear-UserAttribs {
    # Set description and clear other attributes

    Set-ADUser -Server $TargetDC -Identity $User -Description "User Object terminated $termDate"
    Set-ADUser -Server $TargetDC -Identity $User -Clear st, street, streetaddress, postalCode, telephoneNumber, title, l, company, co, c, department, physicalDeliveryOfficeName
    Set-ADUser -Server $TargetDC -Identity $User -HomePhone $null
    Set-ADUser -Server $TargetDC -Identity $User -OfficePhone $null
    Set-ADUser -Server $TargetDC -Identity $User -MobilePhone $null
    Set-ADUser -Server $TargetDC -Identity $User -Clear "extensionAttribute1"
    Set-ADUser -Server $TargetDC -Identity $User -Clear "extensionAttribute10"
    Write-log -LogEntry "Cleared attributes for $($User.DisplayName)." -LogPath $LogPath
}

Function Set-UserMailbox {
    if ($SkipOOO) {
        Write-log -LogEntry  "Skipping assignment of Out of Office replies." -LogPath $LogPath
    } 
    else {
        
        #Define Out of Office replies for internal and external emails

        $FullName = $User.GivenName + " " + $User.Surname #Terminated Users Full Name

        If ($User.Manager) {
            $IntOOO = "Effective $termDate, $FullName is no longer a representative of $($User.Company).  Please contact $($Manager.DisplayName) at $ManagerEmail for assistance with your inquiry." #Internal Out of Office Message
            $ExtOOO = "Effective $termDate, $FullName is no longer a representative of $($User.Company).  Please contact $($Manager.DisplayName) at $ManagerEmail for assistance with your inquiry." #External Out of Office Message
        }
        else {
            $IntOOO = "Effective $termDate, $FullName is no longer a representative of $($User.Company).  Please contact $($User.Company) for assistance with your inquiry." #Internal Out of Office Message
            $ExtOOO = "Effective $termDate, $FullName is no longer a representative of $($User.Company).  Please contact $($User.Company) for assistance with your inquiry." #External Out of Office Message
        }

        # Set Out of Office messages

        Set-MailboxAutoReplyConfiguration -identity $User.UserPrincipalName -AutoReplyState Enabled -InternalMessage $IntOOO -ExternalMessage $ExtOOO
        Write-log -LogEntry Write-Host "$FullName out of office replies set." -LogPath $logp
    }
    
    If ($User.Manager) {
        Add-MailboxPermission -identity $FullName -User $Manager.UserPrincipalName -AccessRights fullaccess -AutoMapping $true
        Write-log -LogEntry "$($Manager.UserPrincipalName) added as a Full Rights delegate for $FullName mailbox." -LogPath $LogPath
    }
    else {
        Write-log -LogEntry "$FullName has no manager defined." -LogPath $LogPath
    }

    # Migrate mailbox to shared mailbox

    Set-Mailbox $User.UserPrincipalName -Type Shared
    Write-log -LogEntry "$FullName mailbox migrated to shared." -LogPath $LogPath

    # Hide terminated user from the Global Address List

    Set-ADUser $User -replace @{msExchHideFromAddressLists = $True}
    Write-log -LogEntry "$FullName hidden from Global Address List." -LogPath $LogPath
}
Function Move-UserTermOU {

    # Calculate Terminated OU

    $termOU = "OU=Terminated,"
    $OUArray = $TargetDC.Split("{.}")
    
    For ($i = 1; $i -lt $OUArray.Count; $i++) {
        If ($i -ne ($OUArray.Count - 1)) {
            $termOU = $termOU + "DC=" + $OUArray[$i] + ","
        }
        else {
            $termOU = $termOU + "DC=" + $OUArray[$i]
        }
    }

    # Move User Object to the Turnover OU

    Move-ADObject -Server $TargetDC -Identity $User -TargetPath $termOU
    Write-log -LogEntry "$($User.DisplayName) moved to $termOU." -LogPath $LogPath
}

Function Remove-O365Licenses {
    # Remove all Office 365 Licenses assigned to user object

    (Get-MsolUser -UserPrincipalName $User.UserPrincipalName).Licenses.AccountSkuId |
        ForEach-Object {
        Set-MsolUserLicense -UserPrincipalName $User.UserPrincipalName -RemoveLicenses $_
    }
}

# Write Log to text file

function Write-Log {
    param(
        [string]$LogEntry, 
        [string]$LogPath
    )
    
    ((Get-Date).ToString() + " - " + $LogEntry) >> $LogPath;
}

# Begin Main Script

$termDate = Get-Date -Format d
$LogPath = ".\Logs\" + $TargetUser + ".txt"

$LogEntry = "Starting termination process for $TargetUser"
Write-log -LogEntry $LogEntry -LogPath $LogPath

# Search for and populate User information

If ($TargetUser) {
    $TargetDC = Get-UserDC $TargetUser
    
    If ($TargetDC) {
        $LogEntry = "Found $TargetUser on $TargetDC"
        Write-Log -LogEntry $LogEntry -LogPath $LogPath
    }
    else {
        $LogEntry = "$TargetUser not found in local domains"
        Write-Log -LogEntry $LogEntry -LogPath $LogPath
        Write-Host $LogEntry -ForegroundColor Red
        Break
    }

    $User = Get-ADUser -server $TargetDC -Identity $TargetUser -Properties DisplayName, UserPrincipalName, HomeDirectory, MemberOf, Manager, mail, msExchHideFromAddressLists, Company 
}

If ($Delegate) {
    $ManagerDC = Get-UserDC $Delegate
    $Manager = Get-ADUser -server $ManagerDC -Identity $Delegate -Properties displayName, UserPrincipalName, ProxyAddresses
    $ManagerEmail = Get-PrimarySMTP $Manager
}
elseIf ($user.Manager) {
    $ManagerDC = Get-UserDC $User.Manager
    $Manager = Get-ADUser -server $ManagerDC -Identity $User.Manager -Properties displayName, UserPrincipalName, ProxyAddresses
    $ManagerEmail = Get-PrimarySMTP $Manager
}

Write-Log -LogEntry "Ready to terminate:" -LogPath $LogPath
Write-Log -LogEntry "User:           $($User.displayName)" -LogPath $LogPath
Write-Log -LogEntry "User Email:     $($User.Mail)" -LogPath $LogPath
Write-Log -LogEntry "User DC:        $TargetDC" -LogPath $LogPath
Write-Log -LogEntry "User Company:   $($User.Company)" -LogPath $LogPath
If ($ManagerEmail) {
    Write-Log -LogEntry "Manger:         $($Manager.displayName)" -LogPath $LogPath
    Write-Log -LogEntry "Manager Email:  $ManagerEmail" -LogPath $LogPath
    Write-Log -LogEntry "Manager DC:     $ManagerDC" -LogPath $LogPath
}

If ($MFA) {
    # Connect to MSOnline Services with MFA

    Import-Module MSOnline 
    Connect-MsolService

    # Connect to Exchange Online with MFA

    Import-Module $((Get-ChildItem -Path $($env:LOCALAPPDATA + "\Apps\2.0\") -Filter Microsoft.Exchange.Management.ExoPowershellModule.dll -Recurse ).FullName|Where-Object {$_ -notmatch "_none_"}|Select-Object -First 1)
    $EXOSession = New-ExoPSSession
    Import-PSSession $EXOSession
}
else { 
    $O365Credentials = get-credential

    # Connect to MSOnline Services

    Import-Module MSOnline 
    Connect-MsolService -Credential $O365Credentials

    # Connect to Exchange Online
    
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $O365Credentials -Authentication Basic -AllowRedirection
    Import-PSSession $Session
}

# Perform tasks against User-Object

Disable-User

Clear-UserGroups

Clear-UserAttribs

if (get-mailbox -Identity $User.UserPrincipalName) {
    # Perform Email and O365 Operations

    Set-UserMailbox
}
else {
    Write-log -LogEntry "$($User.DisplayName) does not have a mailbox assigned." -LogPath $LogPath
}

if ((get-msoluser -UserPrincipalName $User.UserPrincipalName).licenses.servicestatus) {
    # If user-object has licenses assigned remove them

    Remove-O365Licenses
}
else {
    Write-log -LogEntry "$($User.DisplayName) does not have licenses assigned." -LogPath $LogPath
}

Move-UserTermOU

Write-log -LogEntry "$($User.displayname) has been terminated." -LogPath $LogPath

# Disconnect open PSSessions

Get-PSSession | Remove-PSSession