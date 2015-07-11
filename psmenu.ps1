<######################################################################
Multi Layered Powershell Menu 

Written By Dan Arnold
V1.0

######################################################################>

<#$LiveCred = Get-Credential

Import-Module MSOnline
connect-msolservice -Credential $LiveCred

$session=New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell/ -Credential $LiveCred -Authentication Basic -AllowRedirection
Import-PSSession $Session
#>

$xAppName    = "PowershellMenu"
[BOOLEAN]$global:xExitSession=$false
function LoadMenuSystem(){
	[INT]$xMenu1=0
	[INT]$xMenu2=0
	[BOOLEAN]$xValidSelection=$false
	[INT]$xExecute=0

	while ( $xMenu1 -lt 1 -or $xMenu1 -gt 5 ){
		CLS;

		#… Present the Menu Options
		Write-Host "`nPowershell System Administration Menu`n" -ForegroundColor Magenta;
		Write-Host "`tPlease select the admin area`n" -Fore Cyan;
		Write-Host "`t`t1. Get Mailbox Info" -Fore Cyan;
		Write-Host "`t`t2. Mailbox SMTP, Safe Senders, & Blocked Senders" -Fore Cyan;
		Write-Host "`t`t3. Mailbox Sharing, Licensing, Forwarding, & Delegation" -Fore Cyan;
		Write-Host "`t`t4. Complex Processes" -Fore Cyan;
		Write-Host "`t`t5. Quit and exit`n" -Fore Cyan;
		#… Retrieve the response from the user
		[int]$xMenu1 = Read-Host "`tEnter Menu Option Number";
		if( $xMenu1 -lt 1 -or $xMenu1 -gt 5 ){
			Write-Host "`tPlease select one of the options available.`n" -Fore Red;start-Sleep -Seconds 2;
		}
	}
	Switch ($xMenu1){    #… User has selected a valid entry.. load next menu
		1 {
			while ($true) {
				while ( $xMenu2 -lt 1 -or $xMenu2 -gt 9 ){
					CLS;
					# Present the Menu Options
					Write-Host "`nPowershell System Administration Menu`n" -Fore Magenta;
					Write-Host "`tPlease select the Mailbox Information you need`n" -Fore Cyan;
					Write-Host "`t`t1. Get Mailbox Delegate permissions" -Fore Cyan;
					write-host "`t`t2. Get Forwarding Information for a Mailbox" -Fore Cyan;
					write-host "`t`t3. Get Mailbox Size" -Fore Cyan;
					write-host "`t`t4. Does user have an Exchange license?" -Fore Cyan;
					write-host "`t`t5. Is the AD user Enabled or Disabled?" -Fore Cyan;
					write-host "`t`t6. Get AD user information" -Fore Cyan;
					write-host "`t`t7. Get a list of Inbox rules" -Fore Cyan;
					write-host "`t`t8. Get a Global list of an email that exists in Exchange" -Fore Cyan;
					Write-Host "`t`t9. Go to Main Menu`n" -Fore Cyan;
					[int]$xMenu2 = Read-Host "`tEnter Menu Option Number";
					if( $xMenu2 -lt 1 -or $xMenu2 -gt 9 ){
						Write-Host "Please select one of the options available.`n" -Fore Red;start-Sleep -Seconds 2;
					}
				}
				Switch ($xMenu2){
					1{ Write-Host "`n`tYou Selected Option 1 – Get Mailbox Delegate permissions`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter mailbox name";
							Write-Host "`n`tYou entered: $mailbox";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Get-MailboxPermission –identity $mailbox | where {($_.IsInherited -eq $False) -and -not ($_.User -like "NT AUTHORITY\SELF") -and -not ($_.User -like '*Discovery Management*')} | Select Identity,User,AccessRights | FT Identity, User, AccessRights -AutoSize;
							#Get-MailboxPermission –identity $mailbox | where {($_.IsInherited -eq $False) -and -not ($_.User -like "NT AUTHORITY\SELF") -and -not ($_.User -like '*Discovery Management*')} | FT Identity, User, AccessRights -AutoSize;
							Read-Host "`tEnter to return to menu";
						}
					}
					2{ Write-Host "`n`tYou Selected Option 2 - Get Forwarding Information for a Mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox: $mailbox";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Get-Mailbox -identity $mailbox | Where {$_.ForwardingAddress -ne $null} | Select Name, PrimarySMTPAddress, ForwardingAddress, DeliverToMailboxAndForward | FT Name, PrimarySMTPAddress, ForwardingAddress, DeliverToMailboxAndForward -autosize;
							#Get-Mailbox -identity $mailbox | Where {$_.ForwardingAddress -ne $null} | FT Name, PrimarySMTPAddress, ForwardingAddress, DeliverToMailboxAndForward -autosize;
							Read-Host "`tEnter to return to menu";
						}
					}
					3{ Write-Host "`n`tYou Selected Option 3 - Get Mailbox Size`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox: $mailbox";		
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Get-MailboxStatistics -identity $mailbox | ft DisplayName, TotalItemSize, ItemCount -autosize;
							Read-Host "`tEnter to return to menu";
						}
					}
					4{ Write-Host "`n`tYou Selected Option 4 - Does user have an Exchange license?`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox to check: $mailbox";		
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							[string]$l = (Get-MsolUser -UserPrincipalName $mailbox"@cordelllaw.com").isLicensed;
							if ($l -eq "False") {
								Write-Host "`n`t$mailbox does Not have an Exchange License" -Fore Red;
							} else {
								Write-Host "`n`t$mailbox has an Exchange License" -Fore Green;
							}
							Read-Host "`n`tEnter to return to menu";
						}
					}
					5{ Write-Host "`n`tYou Selected Option 5 - Is the AD user Enabled or Disabled?`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox to check: $mailbox";		
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							[string]$e = "";
							$xExecute = 0;
							[string]$e = (Get-ADUser -identity $mailbox -Server Col1dc1).Enabled;
							#write-host "e=$e";
							#write-host "xExecute=$xExecute";
							if ($e -ne "True") {
								Write-Host "`n`t$mailbox is Disabled" -Fore Red;
							} else {
								Write-Host "`n`t$mailbox is Enabled" -Fore Green;
							}
							Read-Host "`n`tEnter to return to menu";
						}
					}
					6{ Write-Host "`n`tYou Selected Option 6 - Get AD user information`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox to check: $mailbox";		
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							[string]$desc = (get-aduser -identity $mailbox -Properties Description).Description;
							Write-host "`tDescription = $desc";
							[string]$office = (get-aduser -identity $mailbox -Properties Office).Office;
							Write-host "`tOffice = $office";
							[string]$displayName = (get-aduser -identity $mailbox -Properties DisplayName).DisplayName;
							Write-host "`tdisplayName = $displayName";
							[string]$givenName = (get-aduser -identity $mailbox -Properties GivenName).GivenName;
							Write-host "`tgivenName = $givenName";
							[string]$surName = (get-aduser -identity $mailbox -Properties SurName).SurName;
							Write-host "`tsurName = $surName";
							[string]$officePhone = (get-aduser -identity $mailbox -Properties OfficePhone).OfficePhone; # same as ipPhone
							Write-host "`tofficePhone = $officePhone";
							[string]$ipPhone = (get-aduser -identity $mailbox -Properties ipPhone).ipPhone;  # same as office phone
							Write-host "`tipPhone = $ipPhone";
							[string]$mobilePhone = (get-aduser -identity $mailbox -Properties mobilePhone).mobilePhone;  
							Write-host "`tmobilePhone = $mobilePhone";
							[string]$fax = (get-aduser -identity $mailbox -Properties Fax).Fax;
							Write-host "`tfax = $fax";
							[string]$email = (get-aduser -identity $mailbox -Properties EmailAddress).EmailAddress;
							Write-host "`temail = $email";
							[string]$street = (get-aduser -identity $mailbox -Properties StreetAddress).StreetAddress;
							Write-host "`tstreet = $street";
							[string]$city = (get-aduser -identity $mailbox -Properties City).City;
							Write-host "`tcity = $city";
							[string]$state = (get-aduser -identity $mailbox -Properties State).State;
							Write-host "`tstate = $state";
							[string]$zip = (get-aduser -identity $mailbox -Properties PostalCode).PostalCode;
							Write-host "`tzip = $zip";
							[string]$country = (get-aduser -identity $mailbox -Properties Co).Co;
							Write-host "`tcountry = $country";
							[string]$samAcct = (get-aduser -identity $mailbox -Properties SamAccountName).SamAccountName;
							Write-host "`tsamAcct = $samAcct";
							[string]$scriptPath = (get-aduser -identity $mailbox -Properties ScriptPath).ScriptPath;
							Write-host "`tscriptPath = $scriptPath";
							[string]$pwNeverExpires = (get-aduser -identity $mailbox -Properties PasswordNeverExpires).PasswordNeverExpires;
							Write-host "`tpwNeverExpires = $pwNeverExpires";
							[string]$pwNotRequired = (get-aduser -identity $mailbox -Properties PasswordNotRequired).PasswordNotRequired;
							Write-host "`tpwNotRequired = $pwNotRequired";
							[string]$homeDrive = (get-aduser -identity $mailbox -Properties HomeDrive).HomeDrive;
							Write-host "`thomeDrive = $homeDrive";
							[string]$homeDir = (get-aduser -identity $mailbox -Properties HomeDirectory).HomeDirectory;
							Write-host "`thomeDir = $homeDir";
							[string]$title = (get-aduser -identity $mailbox -Properties Title).Title;  # same as description
							Write-host "`ttitle = $title";
							[string]$desc = (get-aduser -identity $mailbox -Properties Description).Description;  # same as title
							Write-host "`tdesc = $desc";
							[string]$company = (get-aduser -identity $mailbox -Properties Company).Company;
							Write-host "`tcompany = $company";
							#[string]$container = (get-aduser -identity $mailbox -Properties Container).Container;
							#Write-host "`tcompany = $company";
							$groups = (([ADSISEARCHER]"samaccountname=$mailbox").Findone().Properties.memberof -replace '^CN=([^,]+).+$','$1');
							$i = 0;
							while ($groups[$i] -ne $null) {
								Write-host "`tGroup: " -nonewline;
								Write-host $groups[$i++];
							}
							
							#if ($e -eq "False") {
							#	Write-Host "`n`t$mailbox is Disabled" -Fore Red;
							#} else {
							#	Write-Host "`n`t$mailbox is Enabled" -Fore Green;
							#}
							Read-Host "`n`tEnter to return to menu";
						}
					}
					7{ Write-Host "`n`tYou Selected Option 7 - Get a listing of inbox rules for a mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox: $mailbox";	
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Get-InboxRule -Mailbox $mailbox | FT Name -autosize;
							Read-Host "`tEnter to return to menu";
						}
					}
					8{ Write-Host "`n`tYou Selected Option 8 - Get a Global list of an email that exists in Exchange`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$searchquery = Read-Host "`tEnter the search query (i.e. Subject: Cordell Connections - April 28, 2015)";
							Write-Host "`n`tYou entered search string: $searchquery";					 	
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Get-Mailbox | Search-Mailbox –SearchQuery $searchquery –TargetMailbox itadmin –TargetFolder “DeleteReport” –LogOnly –LogLevel Full;
							Read-Host "`tEnter to return to menu";
						}
					}
					default { Write-Host "`n`tYou Selected Option 9 – Go to Main Menu`n" -Fore Yellow; start-Sleep -Seconds 1;break;}
				}
				if ($xMenu2 -lt 9) {$xMenu2=99;} else {break;}
			}
		}
		2 {
			while ($true) {
				while ( $xMenu2 -lt 1 -or $xMenu2 -gt 7 ){
					CLS;
					# Present the Menu Options
					Write-Host "`nPowershell System Administration Menu`n" -Fore Magenta;
					Write-Host "`tPlease select the Group administration task you require`n" -Fore Cyan;
					Write-Host "`t`t1. Set the primary SMTP address for a Mailbox" -Fore Cyan;
					write-host "`t`t2. Add an email to Safe Senders for all mailboxes" -Fore Cyan;
					write-host "`t`t3. Add an email to Safe Senders for a single mailbox" -Fore Cyan;
					write-host "`t`t4. Add an email to Blocked Senders for a single mailbox" -Fore Cyan;
					write-host "`t`t5. Hide mailbox from GAL" -Fore Cyan;
					write-host "`t`t6. Add mailbox to GAL" -Fore Cyan;
					Write-Host "`t`t7. Go to Main Menu`n" -Fore Cyan;
					[int]$xMenu2 = Read-Host "`tEnter Menu Option Number";
					if( $xMenu2 -lt 1 -or $xMenu2 -gt 7 ){
						Write-Host "Please select one of the options available.`n" -Fore Red;start-Sleep -Seconds 2;
					}
				}
				Switch ($xMenu2){
					1{ Write-Host "`n`tYou Selected Option 1 – Set the primary SMTP address for a mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter mailbox name (i.e. darnold)";
							Write-Host "`n`tYou entered: $mailbox";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-Mailbox –identity $mailbox -WindowsEmailAddress "$mailbox@cordelllaw.com";
							Read-Host "`tEnter to return to menu";
						}
					}
					2{ Write-Host "`n`tYou Selected Option 2 – Add an email to Safe Senders for all mailboxes`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$safe = Read-Host "`tEnter the Safe Sender email (i.e. dan.arnold22@gmail.com)";
							Write-Host "`n`tYou entered: $safe";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Get-Mailbox | Set-MailboxJunkEmailConfiguration -TrustedSendersAndDomains $safe;
							Read-Host "`tEnter to return to menu";
						}
					}
					3{ Write-Host "`n`tYou Selected Option 3 – Add an email to Safe Senders list for a mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox to add a Safe Sender (i.e. darnold)";
							$safe = Read-Host "`tEnter the email of the Safe Sender (i.e. dan.arnold22@gmail.com)";
							Write-Host "`n`tYou entered mailbox: $mailbox and email: $safe";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-MailboxJunkEmailConfiguration –Identity $mailbox -TrustedSendersAndDomains @{Add=$safe};
							Read-Host "`tEnter to return to menu";
						}
					}
					4{ Write-Host "`n`tYou Selected Option 4 – Add an email to Blocked Senders list for a mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox to add a Safe Sender (i.e. darnold)";
							$safe = Read-Host "`tEnter the email of the Safe Sender (i.e. dan.arnold22@gmail.com)";
							Write-Host "`n`tYou entered mailbox: $mailbox and email: $safe";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-MailboxJunkEmailConfiguration –Identity $mailbox -BlockedSendersAndDomains @{Add=$safe};
							Read-Host "`tEnter to return to menu";
						}
					}
					5{ Write-Host "`n`tYou Selected Option 5 – Hide mailbox from GAL`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox to hide (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox: $mailbox";	
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-ADUser -Identity $mailbox -replace @{msExchHideFromAddressLists = $true};
							Read-Host "`tEnter to return to menu";
						}
					}
					6{ Write-Host "`n`tYou Selected Option 6 – Add mailbox to GAL`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox to add (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox: $mailbox";	
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-ADUser -Identity $mailbox -replace @{msExchHideFromAddressLists = $false};
							Read-Host "`tEnter to return to menu";
						}
					}
					default { Write-Host "`n`tYou Selected Option 7 – Go to Main Menu`n" -Fore Yellow; start-Sleep -Seconds 1;break;}
				}
				if ($xMenu2 -lt 7) {$xMenu2=99;} else {break;}
			}
		}
		3 {
			while ($true) {
				while ( $xMenu2 -lt 1 -or $xMenu2 -gt 13 ){
					CLS;
					# Present the Menu Options
					Write-Host "`nPowershell System Administration Menu`n" -Fore Magenta;
					Write-Host "`tPlease select the Mailbox administration task you require`n" -Fore Cyan;
					Write-Host "`t`t1. Convert a Regular Mailbox to a Shared Mailbox" -Fore Cyan;
					Write-Host "`t`t2. Convert a Shared Mailbox to a Regular Mailbox" -Fore Cyan;
					Write-Host "`t`t3. Remove a License from a Mailbox" -Fore Cyan;
					Write-Host "`t`t4. Add a license to a Mailbox" -Fore Cyan;
					Write-Host "`t`t5. Add Delegation rights to a Mailbox" -Fore Cyan;
					Write-Host "`t`t6. Remove Delegation rights from a Mailbox" -Fore Cyan;
					write-host "`t`t7. Add rmexchange and relayserviceaccount delegates" -Fore Cyan;
					Write-Host "`t`t8. Forward a Mailbox to an internal Mailbox" -Fore Cyan;
					Write-Host "`t`t9. Forward a Mailbox to an external email" -Fore Cyan;
					Write-Host "`t`t10. Remove Forward from a Mailbox to an internal Mailbox" -Fore Cyan;
					Write-Host "`t`t11. Remove Forward from a Mailbox to an external email" -Fore Cyan;
					write-host "`t`t12. Delete an email (Globally) that exists in Exchange" -Fore Cyan;	
					Write-Host "`t`t13. Go to Main Menu`n" -Fore Cyan;
					[int]$xMenu2 = Read-Host "`tEnter Menu Option Number";
					if( $xMenu2 -lt 1 -or $xMenu2 -gt 13 ){
						Write-Host "Please select one of the options available.`n" -Fore Red;start-Sleep -Seconds 2;
					}
				}
				Switch ($xMenu2){
					1{ Write-Host "`n`tYou Selected Option 1 – Convert a Regular Mailbox to a Shared Mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter Email address (i.e. darnold@cordelllaw.com)";
							Write-Host "`n`tYou entered: $mailbox";					
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Get-Mailbox $mailbox | Set-Mailbox –ProhibitSendReceiveQuota 10GB –ProhibitSendQuota 9.75GB –IssueWarningQuota 9.5GB –type shared;
							Read-Host "`tEnter to return to menu.  Please remove the Exchange user license (Menu Choice #3)";
						}
					}
					2{ Write-Host "`n`tYou Selected Option 2 – Convert a Shared Mailbox to a Regular Mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter Email address (i.e. darnold@cordelllaw.com)";
							Write-Host "`n`tYou entered: $mailbox";					
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Get-Mailbox $mailbox | Set-Mailbox –ProhibitSendReceiveQuota 50GB –ProhibitSendQuota 49GB –IssueWarningQuota 45GB –type regular;
							Read-Host "`tEnter to return to menu";
						}
					}
					3{ Write-Host "`n`tYou Selected Option 3 – Remove a License from a Mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter Email address (i.e. darnold@cordelllaw.com)";
							Write-Host "`n`tYou entered: $mailbox";					
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-MsolUserLicense -UserPrincipalName $mailbox -RemoveLicenses cordelllaw:EXCHANGESTANDARD;
							# will likely change to cordelllaw:ENTERPRISEPACK or cordelllaw:EMS
							Read-Host "`tEnter to return to menu";
						}
					}

					4{ Write-Host "`n`tYou Selected Option 4 – Add a license to a Mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter Email address (i.e. darnold@cordelllaw.com)";
							Write-Host "`n`tYou entered: $mailbox";					
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							set-msoluser -userprincipalname $mailbox -usagelocation US
							Set-MsolUserLicense -UserPrincipalName $mailbox -AddLicenses cordelllaw:EXCHANGESTANDARD;
							Read-Host "`tEnter to return to menu";
						}
					}
					5{ Write-Host "`n`tYou Selected Option 5 – Add Delegation rights to a Mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox to add delegation to (i.e. darnold)";
							$user = Read-Host "`tEnter the user to add as a delegate (i.e. sdawson)";
							Write-Host "`n`tYou entered mailbox: $mailbox and user: $user";					
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Add-mailboxpermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $true;
							Read-Host "`tEnter to return to menu";
						}
					}
					6{ Write-Host "`n`tYou Selected Option 6 – Remove Delegation rights from a Mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox to remove delegation from (i.e. darnold)";
							$user = Read-Host "`tEnter the user to remove as a delegate (i.e. sdawson)";
							Write-Host "`n`tYou entered mailbox: $mailbox and user: $user";					
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Remove-MailboxPermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All;
							Read-Host "`tEnter to return to menu";
						}
					}
					7{ Write-Host "`n`tYou Selected Option 7 – Add rmexchange and relayserviceaccount delegates`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter mailbox name to add delagates (i.e. darnold)";
							Write-Host "`n`tYou entered: $mailbox";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Add-MailboxPermission -identity $mailbox -user rmexchange -AccessRights FullAccess -InheritanceType All;
							Add-RecipientPermission -identity $mailbox -Trustee rmexchange -AccessRights SendAs;
							Add-RecipientPermission -identity $mailbox -Trustee relayserviceaccount -AccessRights SendAs;
							Read-Host "`tEnter to return to menu";
						}
					}
					8{ Write-Host "`n`tYou Selected Option 8 – Forward a Mailbox to an internal Mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter mailbox name to forward (i.e. darnold)";
							$forward = Read-Host "`tEnter the email to forward to (i.e. sdawson@cordelllaw.com)";
							Write-Host "`n`tYou entered: $mailbox and $forward";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-Mailbox –Identity $mailbox –DeliverToMailboxAndForward $true –ForwardingAddress $forward;
							Read-Host "`tEnter to return to menu";
						}
					}
					9{ Write-Host "`n`tYou Selected Option 9 – Forward a Mailbox to an external email`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter mailbox name to forward (i.e. darnold)";
							$forward = Read-Host "`tEnter the email to forward to (i.e. darnold22@yahoo.com)";
							Write-Host "`n`tYou entered: $mailbox and $forward";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-Mailbox –Identity $mailbox –DeliverToMailboxAndForward $true –ForwardingSMTPAddress $forward;
							Read-Host "`tEnter to return to menu";
						}
					}
					10{ Write-Host "`n`tYou Selected Option 10 – Remove a Forward from a Mailbox to an internal Mailbox`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter mailbox name to remove forward (i.e. darnold)";
							Write-Host "`n`tYou entered: $mailbox";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-Mailbox –Identity $mailbox –DeliverToMailboxAndForward $false –ForwardingAddress $null;
							Read-Host "`tEnter to return to menu";
						}
					}
					11{ Write-Host "`n`tYou Selected Option 11 – Remove a Forward from a Mailbox to an external email`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter mailbox name to remove forward (i.e. darnold)";
							Write-Host "`n`tYou entered: $mailbox";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Set-Mailbox –Identity $mailbox –DeliverToMailboxAndForward $false –ForwardingSMTPAddress $null;
							Read-Host "`tEnter to return to menu";
						}
					}
					12{ Write-Host "`n`tYou Selected Option 12 - Delete an email (Globally) that exists in Exchange`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$searchquery = Read-Host "`tEnter the search query (i.e. From:Dan.arnold@protonmail.com)";
							Write-Host "`n`tYou entered search string: $searchquery";					
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Write-Host "`n`tThis could take a while so please be patient" -Fore Yellow;
							Get-Mailbox | Search-Mailbox –SearchQuery $searchquery –DeleteContent -Force –TargetMailbox itadmin –TargetFolder “DeleteReport” –LogLevel Full;
							Read-Host "`tEnter to return to menu";
						}
					}
					
					default { Write-Host "`n`tYou Selected Option 13 – Go to Main Menu`n" -Fore Yellow; start-Sleep -Seconds 1;break;}
				} # end switch - 2nd level
				if ($xMenu2 -lt 13) {$xMenu2=99;} else {break;}
			} # end while
		} # end switch selection
		4 {
			while ($true) {
				while ( $xMenu2 -lt 1 -or $xMenu2 -gt 6 ){
					CLS;
					# Present the Menu Options
					Write-Host "`nPowershell System Administration Menu`n" -Fore Magenta;
					Write-Host "`tPlease select the process to run`n" -Fore Cyan;
					Write-Host "`t`t1. Convert User to Shared mailbox, remove license, add Delegate and Forwarding" -Fore Cyan;
					write-host "`t`t2. Reset User's Password" -Fore Cyan;
					write-host "`t`t3. Add a User and mailbox (coming soon)" -Fore Cyan;
					write-host "`t`t4. Remove a User from AD and associated mailbox (coming soon)" -Fore Cyan;
					write-host "`t`t5. p5" -Fore Cyan;
					Write-Host "`t`t6. Go to Main Menu`n" -Fore Cyan;
					[int]$xMenu2 = Read-Host "`tEnter Menu Option Number";
					if( $xMenu2 -lt 1 -or $xMenu2 -gt 6 ){
						Write-Host "Please select one of the options available.`n" -Fore Red;start-Sleep -Seconds 2;
					}
				}
				Switch ($xMenu2){
					1{ Write-Host "`n`tYou Selected Option 1 – Convert User to Shared mailbox, remove license, add Delegate and Forwarding`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter mailbox name";
							Write-Host "`n`tYou entered: $mailbox";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							[string]$size = (Get-MailboxStatistics -identity $mailbox).TotalItemSize;
							[decimal]$num = $size.substring(0,5);
							#[decimal]$floor = [math]::Floor($num);
							[decimal]$ceiling = [math]::Ceiling($num);
							if($ceiling -gt 9) {
								Write-host "`n`tMailbox size: $size" -Fore Red;
								Write-host "`tMAILBOX IS TOO LARGE TO CONVERT TO A SHARED MAILBOX" -Fore Red;
							} else {
								Write-host "`n`tMailbox size: $size" -Fore Green;
								Write-host "`n`tMailbox will now be converted to a Shared Mailbox and then the Exchange License will be removed." -Fore Green;
								Get-Mailbox $mailbox | Set-Mailbox –ProhibitSendReceiveQuota 10GB –ProhibitSendQuota 9.75GB –IssueWarningQuota 9.5GB –type shared -whatif;
								if ($lastexitcode -eq 1) {
									write-host "`n`tThe command Failed: Mailbox NOT converted to a Shared Mailbox" -ForegroundColor red;
								} else {
									write-host "`n`tThe command Succeeded: Mailbox converted to a Shared Mailbox" -ForegroundColor green;
									Set-MsolUserLicense -UserPrincipalName $mailbox -RemoveLicenses cordelllaw:EXCHANGESTANDARD;								
									if ($lastexitcode -eq 1) {
										write-host "`n`tThe command Failed: License NOT removed" -ForegroundColor red;
									} else {
										write-host "`n`tThe command Succeeded: License removed" -ForegroundColor green;
										$xExecute = 0;
										while ($xExecute -lt 1) {
											$user = Read-Host "`n`tEnter the user to add as a delegate (i.e. sdawson)";
											Write-Host "`n`tYou entered delegate: $user";					
											Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
											[INT]$xExecute = Read-Host "`n`tChoice";
										}
										if( $xExecute -eq 1) {
											Add-mailboxpermission -Identity $mailbox -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping $true -whatif;
											if ($lastexitcode -eq 1) {
												write-host "`n`tThe command Failed: Delegate NOT added" -ForegroundColor red;
											} else {
												write-host "`n`tThe command Succeeded: Delegate added" -ForegroundColor green;
												$xExecute = 0;
												while ($xExecute -lt 1) {
													$forward = Read-Host "`n`tEnter the email to forward to (i.e. sdawson@cordelllaw.com)";
													Write-Host "`n`tYou entered: $forward";
													Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
													[INT]$xExecute = Read-Host "`n`tChoice";
												}
												if( $xExecute -eq 1) {
													if ($forward -ne "") {
														Set-Mailbox –Identity $mailbox –DeliverToMailboxAndForward $true –ForwardingAddress $forward -whatif;
													} else {
														write-host "`n`tNo Forwarding mailbox specified" -ForegroundColor magenta;
														$lastexitcode = 1;
													}
													if ($lastexitcode -eq 1) {
														write-host "`n`tThe command Failed: Forwarding mailbox NOT added" -ForegroundColor red;
													} else {
														write-host "`n`tThe command Succeeded: Forwarding mailbox added" -ForegroundColor green;		
													}													
													#Read-Host "`tEnter to return to menu";
												}
											}
											#Read-Host "`tEnter to return to menu";
										}								
									}
								}
							}
							#$UserMailboxStats = Get-Mailbox -identity $mailbox -RecipientTypeDetails UserMailbox -ResultSize Unlimited | Get-MailboxStatistics
							#$UserMailboxStats | Add-Member -MemberType ScriptProperty -Name TotalItemSizeInBytes -Value {$this.TotalItemSize -replace "(.*\()|,| [a-z]*\)", ""}
							#$UserMailboxStats | Select-Object DisplayName,@{Name="TotalItemSize (GB)"; Expression={[math]::Round($_.TotalItemSizeInBytes/1GB,2)}} | ft -autosize
							#Write-host "`n`tMailbox size: $x" -Fore Green;					
							Read-Host "`tEnter to return to menu";
						}
					}
					2{ Write-Host "`n`tYou Selected Option 2 - Reset User's Password`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox: $mailbox";
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							$newpwd = ConvertTo-SecureString -String "P@ssw0rd" -AsPlainText -Force;
							Set-ADuser -identity $mailbox -Enabled $True;
							Set-ADAccountPassword $mailbox -NewPassword $newpwd -Reset;
							Set-ADuser -identity $mailbox -ChangePasswordAtLogon $True;
							if ($lastexitcode -eq 1) {
								write-host "`n`tThe command Failed: Password not changed." -ForegroundColor red;
							} else {
								write-host "`n`tThe command Succeeded: Password changed to [P@ssw0rd]." -ForegroundColor green;		
							}													
							Read-Host "`tEnter to return to menu";
						}
					}
					3{ Write-Host "`n`tYou Selected Option 3 - Add a User and mailbox (coming soon)`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the mailbox (i.e. darnold)";
							Write-Host "`n`tYou entered mailbox: $mailbox";		
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							#[string]$office = (get-aduser -identity $mailbox -Properties Office).Office;
							#[string]$displayName = (get-aduser -identity $mailbox -Properties DisplayName).DisplayName;
							#[string]$givenName = (get-aduser -identity $mailbox -Properties GivenName).GivenName;
							#[string]$surName = (get-aduser -identity $mailbox -Properties SurName).SurName;
							#[string]$officePhone = (get-aduser -identity $mailbox -Properties OfficePhone).OfficePhone; # same as ipPhone
							#[string]$fax = (get-aduser -identity $mailbox -Properties Fax).Fax;
							#[string]$ipPhone = (get-aduser -identity $mailbox -Properties ipPhone).ipPhone;  # same as office phone
							#[string]$email = (get-aduser -identity $mailbox -Properties EmailAddress).EmailAddress;
							#[string]$street = (get-aduser -identity $mailbox -Properties StreetAddress).StreetAddress;
							#[string]$city = (get-aduser -identity $mailbox -Properties City).City;
							#[string]$state = (get-aduser -identity $mailbox -Properties State).State;
							#[string]$zip = (get-aduser -identity $mailbox -Properties PostalCode).PostalCode;
							#[string]$country = (get-aduser -identity $mailbox -Properties Co).Co;
							#[string]$samAcct = (get-aduser -identity $mailbox -Properties SamAccountName).SamAccountName;
							#[string]$scriptPath = (get-aduser -identity $mailbox -Properties ScriptPath).ScriptPath;
							#[string]$pwNeverExpires = (get-aduser -identity $mailbox -Properties PasswordNeverExpires).PasswordNeverExpires;
							#[string]$pwNotRequired = (get-aduser -identity $mailbox -Properties PasswordNotRequired).PasswordNotRequired;
							#[string]$homeDrive = (get-aduser -identity $mailbox -Properties HomeDrive).HomeDrive;
							#[string]$homeDir = (get-aduser -identity $mailbox -Properties HomeDirectory).HomeDirectory;
							#[string]$title = (get-aduser -identity $mailbox -Properties Title).Title;  # same as description
							#[string]$desc = (get-aduser -identity $mailbox -Properties Description).Description;  # same as title
							#[string]$company = (get-aduser -identity $mailbox -Properties Company).Company;
							#$groups = (([ADSISEARCHER]"samaccountname=$mailbox").Findone().Properties.memberof -replace '^CN=([^,]+).+$','$1');
							#$i = 0;
							#while ($groups[$i] -ne $null) {
							#	Write-host "`tGroup: " -nonewline;
							#	Write-host $groups[$i++];
							#}


							Import-Csv "NewUsers.csv" | ForEach-Object {
								$userPrinc = $_."Logon Username" + "@cordelllaw.com"; write-host "user: $userPrinc";
								$scriptPath = $_."Script Path" + $_."Logon Username"; write-host "script: $scriptPath";
								New-QADUser -Name $_."Name" -ParentContainer $_."Container" -SamAccountName $_."Logon Username" -UserPassword "pass123!ForWhat" `
									 -GivenName $_."First Name" -surName $_."Last Name" -LogonScript $scriptPath -Description $_."Title" -EmailAddress $userPrinc `
									 -DisplayName $_."Name" -Office $_."Office" -OfficePhone $_."Phone" -ipPhone $_."Phone" -mobilePhone $_."mobilePhone" -Fax $_."Fax" `
									 -WhatIf; `
								Add-QADGroupMember -identity $_."Users" -Member $_."Logon Username" -WhatIf; `
								Set-QADUser -identity $_."Logon Username" `
									-UserMustChangePassword $true -WhatIf;`
							}
							Read-Host "`tEnter to return to menu";
						}
					}
					4{ Write-Host "`n`tYou Selected Option 4 - Remove a User from AD and associated mailbox (coming soon)`n" -Fore Yellow;
						$xExecute = 0;
						while ($xExecute -lt 1) {
							$mailbox = Read-Host "`tEnter the user to delete (i.e. darnold)";
							Write-Host "`n`tYou entered user: $mailbox";	
							Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
							[INT]$xExecute = Read-Host "`n`tChoice";
						}
						if( $xExecute -eq 1) {
							Remove-ADUser -identity $mailbox -Confirm:$false;
							if ($lastexitcode -eq 1) {
								write-host "`n`tThe command Failed: AD User and mailbox not deleted." -ForegroundColor red;
							} else {
								write-host "`n`tThe command Succeeded: AD User and mailbox deleted." -ForegroundColor green;		
							}													
							Read-Host "`tEnter to return to menu";
						}
					}
					5{ Write-Host "`n`tYou Selected Option 5 - p5`n" -Fore Yellow;
						$xExecute = 0;
						#while ($xExecute -lt 1) {
						#	$searchquery = Read-Host "`tEnter the search query (i.e. Subject: Cordell Connections - April 28, 2015)";
						#	Write-Host "`n`tYou entered search string: $searchquery";					 	
						#	Write-host "`n`tEnter " -nonewline;Write-host "[1]" -nonewline -Fore Yellow; Write-host " to continue or " -nonewline; Write-host "[0]" -nonewline -Fore Yellow; Write-host " to re-enter or" -nonewline; Write-host " [99]" -nonewline -Fore Yellow; Write-host " to abort and return to menu." -nonewline;
						#	[INT]$xExecute = Read-Host "`n`tChoice";
						#}
						#if( $xExecute -eq 1) {
						#	Get-Mailbox | Search-Mailbox –SearchQuery $searchquery –TargetMailbox itadmin –TargetFolder “DeleteReport” –LogOnly –LogLevel Full;
						#	Read-Host "`tEnter to return to menu";
						#}
					}
					default { Write-Host "`n`tYou Selected Option 6 – Go to Main Menu`n" -Fore Yellow; start-Sleep -Seconds 1;break;}
				}
				if ($xMenu2 -lt 6) {$xMenu2=99;} else {break;}
			}
		}
		default { $global:xExitSession=$true}
	} # end switch - 1st level
$xExitSession=$true
}
if ($xExitSession -eq 0) {LoadMenuSystem}
If ($xExitSession){
	exit-pssession    #… User quit & Exit
}else{
	.\PSMenu_v1_0.ps1    #… Loop the function

}
