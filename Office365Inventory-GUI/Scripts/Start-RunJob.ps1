Function Start-RunJob {
	Param(
		$header,
		$observable
	)
	#verify if connected
	$connected = Check-Connection
	Write-Debug "Conntect to Office 365: $($connected)"
	
	if($connected -eq $true){
		If ($header -eq 'AAD Users') {
		
			$Script:AADUsers_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]    
			$AADUsers_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all Azure Active Directory Users...Please Wait"
			})
			
			try{
				$users = get-msoluser -all -ea stop | where{$_.UserPrincipalName -notlike "*#ext#*"} | select DisplayName, FirstName, LastName, UserPrincipalName, Title, Department, Office, PhoneNumber, MobilePhone, CloudAnchor, IsLicensed, @{Name="License"; Expression = {$_.licenses.accountskuid}} 
				
				ForEach ($user in $users) { 
					If (-NOT [System.String]::IsNullOrEmpty($user)) { 
						$AADUsers_Observable.Add((
							New-Object PSObject -Property @{
								AADUser_DisplayName = $user.DisplayName
								AADUser_FirstName = $user.FirstName
								AADUser_LastName = $user.LastName	
								AADUser_UserPrincipalName = $user.UserPrincipalName
								AADUser_Title = $user.Title
								AADUser_Department = $user.Department
								AADUser_Office = $user.Office	
								AADUser_PhoneNumber = $user.PhoneNumber
								AADUser_MobilePhone = $user.MobilePhone
								AADUser_CloudAnchor = $user.CloudAnchor
								AADUser_IsLicensed = $user.IsLicensed
								AADUser_Licenses = $user.License										
							}
						))						
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.AADUsers_Listview.ItemsSource = $AADUsers_Observable
					$Global:Clients = $AADUsers_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Azure Active Directory Users."	
					$uiHash.AADUsers_Image.Source = "$pwd\Images\Check_Okay.ico"
				})
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.AADUsers_Image.Source = "$pwd\Images\Check_Error.ico"
				})			
			}		
		}
		elseif ($header -eq 'AAD Deleted Users') {
			$Script:AADDeletedUsers_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$AADDeletedUsers_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all Azure Active Directory Deleted Users...Please Wait"
			})
			
			try{
				$users = Get-MsolUser -All -ReturnDeletedUsers | select SignInName, UserPrincipalName, DisplayName, SoftDeletionTimestamp, IsLicensed, @{Name="License"; Expression = {$_.licenses.accountskuid}} 
				
				ForEach ($user in $users) { 
					If (-NOT [System.String]::IsNullOrEmpty($user)) {  
						$AADDeletedUsers_Observable.Add((
							New-Object PSObject -Property @{
								AADDeletedUser_SignInName = $user.SignInName
								AADDeletedUser_UserPrincipalName = $user.UserPrincipalName
								AADDeletedUser_DisplayName = $user.DisplayName
								AADDeletedUser_SoftDeletionTimestamp = $user.SoftDeletionTimestamp
								AADDeletedUser_IsLicensed = $user.IsLicensed
								AADDeletedUser_Licenses = $user.License
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.AADDeletedUsers_Listview.ItemsSource = $AADDeletedUsers_Observable
					$Global:Clients = $AADDeletedUsers_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Azure Active Directory Deleted Users."		
					$uiHash.AADDeletedUsers_Image.Source = "$pwd\Images\Check_Okay.ico"
				})
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.AADDeletedUsers_Image.Source = "$pwd\Images\Check_Error.ico"
				})		
			}		
		}
		elseif ($header -eq 'AAD External Users') {
			$Script:AADExternalUsers_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$AADExternalUsers_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all Azure Active Directory External Users...Please Wait"
			})
			
			try{
				$users = Get-MsolUser -all | where{$_.UserPrincipalName -like "*#ext#*"} | select SignInName, UserPrincipalName, DisplayName, WhenCreated
				
				ForEach ($user in $users) { 
					If (-NOT [System.String]::IsNullOrEmpty($user)) {  
						$AADExternalUsers_Observable.Add((
							New-Object PSObject -Property @{
								AADExternalUser_SignInName = $user.SignInName
								AADExternalUser_UserPrincipalName = $user.UserPrincipalName
								AADExternalUser_DisplayName = $user.DisplayName
								AADExternalUser_WhenCreated = $user.WhenCreated
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.AADExternalUsers_Listview.ItemsSource = $AADExternalUsers_Observable
					$Global:Clients = $AADExternalUsers_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Azure Active Directory External Users."		
					$uiHash.AADExternalUsers_Image.Source = "$pwd\Images\Check_Okay.ico"
				})
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.AADExternalUsers_Image.Source = "$pwd\Images\Check_Error.ico"
				})		
			}		
		}
		elseif ($header -eq 'AAD Contacts') {
			$Script:AADContacts_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$AADContacts_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all Azure Active Directory Contacts...Please Wait"
			})
			
			try{
				$Contacts = Get-Msolcontact -all | select DisplayName, EmailAddress
				
				ForEach ($Contact in $Contacts) { 
					If (-NOT [System.String]::IsNullOrEmpty($Contact)) {  
						$AADContacts_Observable.Add((
							New-Object PSObject -Property @{
								AADContacts_DisplayName = $Contact.DisplayName
								AADContacts_EmailAddress = $Contact.EmailAddress
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.AADContacts_Listview.ItemsSource = $AADContacts_Observable
					$Global:Clients = $AADContacts_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Azure Active Directory Contacts."	
					$uiHash.AADContacts_Image.Source = "$pwd\Images\Check_Okay.ico"		
				})
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.AADContacts_Image.Source = "$pwd\Images\Check_Error.ico"
				})	
			}		
		}
		elseif ($header -eq 'AAD Groups') {
			$Script:AADGroups_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$AADGroups_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all Azure Active Directory Groups...Please Wait"
			})
			
			try{
				$groups = get-msolGroup | select DisplayName, EmailAddress, GroupType, ValidationStatus
				
				ForEach ($group in $groups) { 
					If (-NOT [System.String]::IsNullOrEmpty($group)) {  
						$AADGroups_Observable.Add((
							New-Object PSObject -Property @{
								AADGroup_GroupType = $group.GroupType
								AADGroup_DisplayName = $group.DisplayName
								AADGroup_EmailAddress = $group.EmailAddress
								AADGroup_ValidationStatus = $group.ValidationStatus
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.AADGroups_Listview.ItemsSource = $AADGroups_Observable
					$Global:Clients = $AADGroups_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Azure Active Directory Groups."	
					$uiHash.AADGroups_Image.Source = "$pwd\Images\Check_Okay.ico"	
				})
			}
			catch{
				error-action -message "Error retrieving Azure Active Directory Groups."	
				$uiHash.AADGroups_Image.Source = "$pwd\Images\Check_Error.ico"
			}
		}
		elseif ($header -eq 'AAD Licenses') {
			$Script:AADLicenses_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$AADLicenses_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all Azure Active Licenses...Please Wait"
			})
			
			try{
				$licenses = Get-MsolAccountSku | select AccountSkuID, ActiveUnits, ConsumedUnits, LockedOutUnits
				
				ForEach ($license in $licenses) { 
					If (-NOT [System.String]::IsNullOrEmpty($license)) {  
						$AADLicenses_Observable.Add((
							New-Object PSObject -Property @{
								AADLicenses_AccountSkuID = $license.AccountSkuID
								AADLicenses_ActiveUnits = $license.ActiveUnits
								AADLicenses_ConsumedUnits = $license.ConsumedUnits
								AADLicenses_LockedOutUnits = $license.LockedOutUnits
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.AADLicenses_Listview.ItemsSource = $AADLicenses_Observable
					$Global:Clients = $AADLicenses_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Azure Active Directory Licenses."
					$uiHash.AADLicenses_Image.Source = "$pwd\Images\Check_Okay.ico"	
				})
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.AADLicenses_Image.Source = "$pwd\Images\Check_Error.ico"
				})	
			}				
		}
		elseif ($header -eq 'AAD Domains') {
			$Script:AADDomains_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$AADDomains_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				Start-Action -message "Retrieving all Azure Active Domains...Please Wait"
			})
			
			try{
				$domains = Get-MsolDomain | select Name, Status, Authentications
				
				ForEach ($domain in $Domains) { 
					If (-NOT [System.String]::IsNullOrEmpty($domain)) {  
						$AADDomains_Observable.Add((
							New-Object PSObject -Property @{
								AADDomains_Name = $domain.Name
								AADDomains_Status = $domain.Status
								AADDomains_Authentications = $domain.Authentications
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.AADDomains_Listview.ItemsSource = $AADDomains_Observable
					$Global:Clients = $AADDomains_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Azure Active Directory Domains."
					$uiHash.AADDomains_Image.Source = "$pwd\Images\Check_Okay.ico"	
				})
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.AADDomains_Image.Source = "$pwd\Images\Check_Error.ico"
				})	
			}				
		}
		elseif ($header -eq 'Exchange Mailboxes') {
			$Script:ExchangeMailboxes_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$ExchangeMailboxes_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all Exchange Mailboxes...Please Wait"
			})
			
			try{
				$ExchangeMailboxes = Get-Mailbox | sort DisplayName | select DisplayName, Alias, PrimarySMTPAddress, ArchiveStatus, UsageLocation, WhenMailboxCreated,LegacyExchangeDN,EmailAddresses,RecipienttypeDetails
				
				ForEach ($ExchangeMailbox in $ExchangeMailboxes) { 
					If (-NOT [System.String]::IsNullOrEmpty($ExchangeMailbox)) { 
						$statistics = Get-MailboxStatistics $ExchangeMailbox.alias -WarningAction:SilentlyContinue| select ItemCount, TotalItemSize, LastLogonTime
						$ExchangeMailboxes_Observable.Add((
							New-Object PSObject -Property @{
								ExchangeMailboxes_DisplayName = $ExchangeMailbox.DisplayName
								ExchangeMailboxes_Alias = $ExchangeMailbox.Alias
								ExchangeMailboxes_PrimarySMTPAddress = $ExchangeMailbox.PrimarySMTPAddress
								ExchangeMailboxes_ItemCount = $statistics.ItemCount
								ExchangeMailboxes_TotalItemSize = $statistics.TotalItemSize
								ExchangeMailboxes_ArchiveStatus = $ExchangeMailbox.ArchiveStatus
								ExchangeMailboxes_UsageLocation = $ExchangeMailbox.UsageLocation
								ExchangeMailboxes_WhenMailboxCreated = $ExchangeMailbox.WhenMailboxCreated
								ExchangeMailboxes_ProxyAddresses = $ExchangeMailbox.EmailAddresses -join ";"
								ExchangeMailboxes_LegacyExchangeDN = $ExchangeMailbox.LegacyExchangeDN
								ExchangeMailboxes_RecipientTypeDetails = $ExchangeMailbox.RecipienttypeDetails
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.ExchangeMailboxes_Listview.ItemsSource = $ExchangeMailboxes_Observable
					$Global:Clients = $ExchangeMailboxes_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Exchange Mailboxes."
					$uiHash.ExchangeMailboxes_Image.Source = "$pwd\Images\Check_Okay.ico"	
				})
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.ExchangeMailboxes_Image.Source = "$pwd\Images\Check_Error.ico"
				})	
			}				
		}
		elseif ($header -eq 'Exchange Archives') {
			$Script:ExchangeArchives_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$ExchangeArchives_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all Exchange Archives...Please Wait"
			})
			
			try{
				$ExchangeArchives = Get-Mailbox -Archive | sort DisplayName | select DisplayName, Alias, PrimarySMTPAddress, ArchiveStatus, UsageLocation, WhenMailboxCreated
				
				ForEach ($ExchangeArchive in $ExchangeArchives) { 
					If (-NOT [System.String]::IsNullOrEmpty($ExchangeArchive)) {
						$statistics = Get-MailboxStatistics $ExchangeArchive.alias -archive -WarningAction:SilentlyContinue| select ItemCount, TotalItemSize, LastLogonTime
						$ExchangeArchives_Observable.Add((
							New-Object PSObject -Property @{
								ExchangeArchives_DisplayName = $ExchangeArchive.DisplayName
								ExchangeArchives_Alias = $ExchangeArchive.Alias
								ExchangeArchives_PrimarySMTPAddress = $ExchangeArchive.PrimarySMTPAddress
								ExchangeArchives_ItemCount = $statistics.ItemCount
								ExchangeArchives_TotalItemSize = $statistics.TotalItemSize
								ExchangeArchives_ArchiveStatus = $ExchangeArchive.ArchiveStatus
								ExchangeArchives_UsageLocation = $ExchangeArchive.UsageLocation
								ExchangeArchives_WhenMailboxCreated = $ExchangeArchive.WhenMailboxCreated
								ExchangeArchives_LastLogonTime = $statistics.LastLogonTime
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.ExchangeArchives_Listview.ItemsSource = $ExchangeArchives_Observable
					$Global:Clients = $ExchangeArchives_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Exchange Archives."
					$uiHash.ExchangeArchives_Image.Source = "$pwd\Images\Check_Okay.ico"
				})
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.ExchangeArchives_Image.Source = "$pwd\Images\Check_Error.ico"
				})	
			}				
		}
		elseif ($header-eq 'Exchange Groups') {
			$Script:ExchangeGroups_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$ExchangeGroups_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all Exchange Groups...Please Wait"
			})
			
			try{
				$ExchangeGroups = Get-Group | where{$_.RecipientTypeDetails -ne "RoleGroup"} | sort DisplayName | select DisplayName, RecipientTypeDetails, @{Name="Owner"; Expression = {$_.ManagedBy}}, WindowsEmailAddress
				
				ForEach ($ExchangeGroup in $ExchangeGroups) { 
					If (-NOT [System.String]::IsNullOrEmpty($ExchangeGroup)) {
						$ExchangeGroups_Observable.Add((
							New-Object PSObject -Property @{
								ExchangeGroups_DisplayName = $ExchangeGroup.DisplayName
								ExchangeGroups_RecipientTypeDetails = $ExchangeGroup.RecipientTypeDetails
								ExchangeGroups_Owner = $ExchangeGroup.Owner
								ExchangeGroups_WindowsEmailAddress = $ExchangeGroup.WindowsEmailAddress
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.ExchangeGroups_Listview.ItemsSource = $ExchangeGroups_Observable
					$Global:Clients = $ExchangeGroups_Observable | Select -Expand DisplayName
					end-action -message "Retrieved Exchange Groups."
					$uiHash.ExchangeGroups_Image.Source = "$pwd\Images\Check_Okay.ico"	
				})
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.ExchangeGroups_Image.Source = "$pwd\Images\Check_Error.ico"
				})	
			}				
		}
		elseif ($header -eq 'SharePoint Sites') {
			$Script:SharePointSites_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$SharePointSites_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all SharePoint Sites...Please Wait"
			})
			
			try{
				$SharePointSites = get-sposite | select url, title, WebsCount, StorageUsageCurrent, Status, LocaleId, template, owner, LastContentModifiedDate
				
				ForEach ($SharePointSite in $SharePointSites) { 
					If (-NOT [System.String]::IsNullOrEmpty($SharePointSite)) {
						$SharePointSites_Observable.Add((
							New-Object PSObject -Property @{
								SharePointSites_url = $SharePointSite.url
								SharePointSites_title = $SharePointSite.title
								SharePointSites_WebsCount = $SharePointSite.WebsCount
								SharePointSites_StorageUsageCurrent = $SharePointSite.StorageUsageCurrent
								SharePointSites_status = $SharePointSite.Status
								SharePointSites_localid = $SharePointSite.LocaleId
								SharePointSites_template = $SharePointSite.template
								SharePointSites_owner = $SharePointSite.owner
								SharePointSites_LastContentModifiedDate = $SharePointSite.LastContentModifiedDate
							}
						))   					
					}
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.SharePointSites_Listview.ItemsSource = $SharePointSites_Observable
					$Global:Clients = $SharePointSites_Observable | Select -Expand DisplayName
					end-action -message "Retrieved SharePoint Sites."
					$uiHash.SharePointSites_Image.Source = "$pwd\Images\Check_Okay.ico"		
				})	
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.SharePointSites_Image.Source = "$pwd\Images\Check_Error.ico"
				})	
			}				
		}
		elseif ($header -eq 'SharePoint Webs') {
			$Script:SharePointWebs_Observable = New-Object System.Collections.ObjectModel.ObservableCollection[object]
			$SharePointWebs_Observable.Clear()
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				start-action -message "Retrieving all SharePoint Webs...Please Wait"
			})
			
			try{
				$siteCollections = Get-SPOSite
				$SPOCredentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credential.Username, $Credential.password)
				
				#loop through all site collections
				foreach ($siteCollection in $siteCollections){
					#search for webs
					$AllWebs = Get-SPOWebs -url $siteCollection.url -spocredentials $SPOCredentials
				}
				
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{   
					$uiHash.SharePointWebs_Listview.ItemsSource = $SharePointWebs_Observable
					$Global:Clients = $SharePointWebs_Observable | Select -Expand DisplayName
					end-action -message "Retrieved SharePoint Webs."
					$uiHash.SharePointWebs_Image.Source = "$pwd\Images\Check_Okay.ico"		
				})	
			}
			catch{
				$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
					error-action -message $_.exception.Message
					$uiHash.SharePointWebs_Image.Source = "$pwd\Images\Check_Error.ico"
				})	
			}				
		}
		Else {
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				error-action -message "No action available at home tab"
			})
		}
	}
	else{
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			error-action -message "Please first connect to Office 365 by using the connect button."
		})
	}
}

function Check-Connection(){
	try
	{
		Get-MsolDomain -ErrorAction Stop > $null
		
		if ((Get-PSSession).count -gt 0)
		{
			$connected = $true
		}
		else{
			$connected = $false
		}
		
	}
	catch 
	{
		$connected = $false
	}
	return $connected
}

function Get-SPOWebs(){
	param(
		$URL,
		$SPOCredentials
	)
	
	$context = New-Object Microsoft.SharePoint.Client.ClientContext($url)
	$context.Credentials = $SPOcredentials
	
	$web = $context.Web
	$context.Load($web)
	$context.Load($web.Webs)
	
	try{
		$context.ExecuteQuery()
		
		If (-NOT [System.String]::IsNullOrEmpty($web)) {
			$SharePointWebs_Observable.Add((
				New-Object PSObject -Property @{
					SharePointWebs_ServerRelativeUrl = $web.ServerRelativeUrl
					SharePointWebs_Title = $web.title
					SharePointWebs_Created = $web.created
					SharePointWebs_LastItemModifiedDate = $web.LastItemModifiedDate
				}
			))
		}

		foreach($web in $web.Webs) {	
			If (-NOT [System.String]::IsNullOrEmpty($web)) {
				$SharePointWebs_Observable.Add((
					New-Object PSObject -Property @{
						SharePointWebs_ServerRelativeUrl = $web.ServerRelativeUrl
						SharePointWebs_Title = $web.title
						SharePointWebs_Created = $web.created
						SharePointWebs_LastItemModifiedDate = $web.LastItemModifiedDate
					}
				))
			}
			Get-SPOWebs($web.url)
		}
	}
	catch{
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			$uiHash.StatusTextBox.Foreground = "orange"
			$uiHash.StatusTextBox.Text = "Error connecting to $($url)...Continuing search"
		})	
	}
 
}
