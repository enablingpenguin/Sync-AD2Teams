## =====Overview=====================================================================================
## PowerShell script to sync AAD group ("SG") and Office 365 group ("O365 Group")
## SG membership takes precedence (primary)- users in the SG are added to the O365 Group ("replica"),
## and users not in the SG are removed from the O365 Group
##
## Please read important notes in the accompanying blog post. https://aka.ms/SyncGroupsScript
##
## This script probably requires more hardening against various situations including:
##  - nested groups
##  - different types of groups
##  - Unicode email aliases
##  - and more
## ==================================================================================================             
##
## =====Author Info==================================================================================
## Dan Stevenson, Microsoft Corporation, Taipei, Taiwan
## Email (and Teams): dansteve@microsoft.com
## Twitter: @danspot
## LinkedIn: https://www.linkedin.com/in/dansteve/
## ==================================================================================================             
## 
## ======MIT License=================================================================================             
## Copyright 2018 Microsoft Corporation
## Permission is hereby granted, free of charge, to any person obtaining a copy of this software
## and associated documentation files (the "Software"), to deal in the Software without restriction,
## including without limitation the rights to use, copy, modify, merge, publish, distribute,
## sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
## furnished to do so, subject to the following conditions:
##
## The above copyright notice and this permission notice shall be included in all copies or
## substantial portions of the Software.
##
## THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
## BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
## NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
## DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
## OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
## ==================================================================================================             
##
## =====Version History==============================================================================             
## [date]       [version]       [notes]
## --------------------------------------------------------------------------------
## 8/22/18      0.3             fixed 2 bugs: adding members from an array not a string, and closing the Exchange session
## 8/16/18      0.2             mostly debugged end-to-end, including removing O365 members who are notin the SG
## 7/27/18      0.1             initial draft script, just working notes
## ==================================================================================================             
##

Start-Transcript -Path C:\TeamsSyncTool\log.txt

function Send-Email {
    $from = "{FromEMAIL}" 
    $to = "{TOEMAIL}"
    $smtp = "smtp.office365.com" 
    $sub = "Teams Sync Log" 
    $body = "Log Attached"
    $secpasswd = ConvertTo-SecureString "So*iCfB5txKt9D" -AsPlainText -Force
    $mycreds = New-Object System.Management.Automation.PSCredential($from, $secpasswd)
    Send-MailMessage -UseSsl -Attachments C:\TeamsSyncTool\log.txt -To $to -From $from -Subject $sub -Body $body -Credential $mycreds -SmtpServer $smtp -DeliveryNotificationOption Never -BodyAsHtml
}

function Get-JDMsolGroupMember { 
<#
.SYNOPSIS
    The function enumerates Azure AD Group members with the support for nested groups.
.EXAMPLE
    Get-JDMsolGroupMember 6d34ab03-301c-4f3a-8436-98f873ec121a
.EXAMPLE
    Get-JDMsolGroupMember -ObjectId  6d34ab03-301c-4f3a-8436-98f873ec121a -Recursive
.EXAMPLE
    Get-MsolGroup -SearchString "Office 365 E5" | Get-JDMsolGroupMember -Recursive
.NOTES
    Author   : Johan Dahlbom, johan[at]dahlbom.eu
    Blog     : 365lab.net 
    The script are provided “AS IS” with no guarantees, no warranties, and it confer no rights.
#>
 
    param(
        [CmdletBinding(SupportsShouldProcess=$true)]
        [Parameter(Mandatory=$true, ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [ValidateScript({Get-MsolGroup -ObjectId $_})]
        $ObjectId,
        [switch]$Recursive
    )
    begin {
        $MSOLAccountSku = Get-MsolAccountSku -ErrorAction Ignore -WarningAction Ignore
        if (-not($MSOLAccountSku)) {
            throw "Not connected to Azure AD, run Connect-MsolService"
        }
    } 
    process {
        Write-Verbose -Message "Enumerating group members in group $ObjectId"
        $UserMembers = @(Get-MsolGroupMember -GroupObjectId $ObjectId -MemberObjectTypes User -All)
        if ($PSBoundParameters['Recursive']) {
            $GroupsMembers = Get-MsolGroupMember -GroupObjectId $ObjectId -MemberObjectTypes Group -All
            if ($GroupsMembers) {
                Write-Verbose -Message "$ObjectId have $($GroupsMembers.count) group(s) as members, enumerating..."
                $GroupsMembers | ForEach-Object -Process {
                    Write-Verbose "Enumerating nested group $($_.Displayname) ($($_.ObjectId))"
                    $UserMembers += Get-JDMsolGroupMember -Recursive -ObjectId $_.ObjectId 
                }
            }
        }
        Return ($UserMembers | Sort-Object -Property EmailAddress -Unique) 
         
    }
    end {
     
    }
}

# get credentials and login as Exchange admin and PS Session (remember to close session later)
$securePassword=ConvertTo-SecureString "{PASSWORD}" -AsPlainText -Force
$ExchangeCred = New-Object System.Management.Automation.PSCredential("{AUTHEMAIL}", $securePassword)
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $ExchangeCred -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking


# get credentials and login as AAD admin
#$AADCred = Get-Credential -Message "AAD admin login"
Connect-MsolService -Credential $ExchangeCred

$IgnoreUsers = @("{Comma Seperated List to emails to ignore}")

$TeamsToSync = "C:\TeamsSyncTool\teamstosync.csv"
$GroupList = Import-Csv $TeamsToSync

$GroupList | ForEach-Object {

	# get name of O365 Group to look up
	$O365GroupName = $_.o365name

	# get name of Group to look up
	$GroupName = $_.adname

	# get O365 Group ID
	### you can do this via AAD as well: $O365Group = Get-MsolGroup | Where-Object {$_.DisplayName -eq $O365GroupName}
	#$O365Group = Get-UnifiedGroup -Identity $O365GroupName
	$O365Group = Get-MsolGroup | Where-Object {$_.DisplayName -eq $O365GroupName}
	$O365GroupID = $O365Group.ObjectID.ToString()
    Write-Host ""
	Write-Host "Syncing $O365GroupName to $GroupName"
	Write-Host "O365 Group ID: $O365GroupID"


	# get list of O365 Group members
	### you can do this via AAD as well: $O365GroupMembers = Get-MsolGroupMember -GroupObjectId $O365GroupID
	$O365GroupMembers = Get-UnifiedGroupLinks -Identity $O365GroupID -LinkType members -resultsize unlimited

	# get Group ID
	$Group = Get-MsolGroup | Where-Object {$_.DisplayName -eq $GroupName}
	$GroupID = $Group.ObjectId
	Write-Host "Group ID: $GroupID"

	# get list of Group members
	$GroupMembers = Get-JDMsolGroupMember $GroupID -Recursive

	# loop through all Group members and add them to a list
	# might be more efficient (from a service API perspective) to have an inner foreach 
	# loop that verifies the user is not in the O365 Group
	Write-Host "Loading list of Group members"
	$GroupMembersToAdd = New-Object System.Collections.ArrayList
	foreach ($GroupMember in $GroupMembers) 
	{
			$memberType = $GroupMember.GroupMemberType
			if ($memberType -eq 'User') {
					$memberEmail = $GroupMember.EmailAddress
					$GroupMembersToAdd.Add($memberEmail) | Out-Null
			}
	}

	# add all the Group members to the O365 Group
	# this is not super efficient - might be better to remove any existing members first
	# this might need to be broken into multiple calls depending on API limitations
	Write-Output "Adding Group members to O365 Group"
	Add-UnifiedGroupLinks -Identity $O365GroupID -LinkType Members -Links $GroupMembersToAdd

	# loop through the O365 Group and remove anybody who is not in the group
	Write-Output "Looking for O365 Group members who are not in Group"
	$O365GroupMembersToRemove = New-Object System.Collections.ArrayList
	foreach ($O365GroupMember in $O365GroupMembers) {
			$userFound = 0
			foreach ($emailAddress in $O365GroupMember.EmailAddresses) {
	# trim the protocol ("SMTP:")
					$emailAddress = $emailAddress.substring($emailAddress.indexOf(":")+1,$emailAddress.length-$emailAddress.indexOf(":")-1)
					if ($GroupMembersToAdd.Contains($emailAddress)) { $userFound = 1 }
			}
			if ($userFound -eq 0 -and $IgnoreUsers -notcontains $O365GroupMember.PrimarySmtpAddress) { $O365GroupMembersToRemove.Add($O365GroupMember) | Out-Null }
	}


	if ($O365GroupMembersToRemove.Count -eq 0) {
			Write-Host "No Users to Remove"
	} else {
	# remove members
			Write-host -ForegroundColor Magenta "Removing $O365GroupMembersToRemove"
					foreach ($memberToRemove in $O365GroupMembersToRemove) {
					Remove-UnifiedGroupLinks -Identity $O365GroupID -LinkType Members -Links $memberToRemove.name -Confirm:$false
			}
	}

}

# close the Exchange session
Remove-PSSession $Session

Stop-Transcript
Send-Email