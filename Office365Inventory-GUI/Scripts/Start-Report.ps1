#Report function
$reportName = ""
Function Start-Report {
    Write-Debug ("Data: {0}" -f $uiHash.ReportComboBox.SelectedItem.Text)
	Write-Debug ("Active Tab {0}" -f $uiHash.Tabs.SelectedItem.header)
	$reportpath = "$($Path)/report"
	$date = get-date
	$today = $date.ToString("ddMMyyyy_HHmm")
	
	#Get Correct listview
	Switch ($uiHash.Tabs.SelectedItem.header){
		"Home"{
			$uiHash.StatusTextBox.Foreground = "Red"
			$uiHash.StatusTextBox.Text = "Reports cannot be created at the home tab." 
		}
		"AAD Users"{
			$lisview = $uiHash.AADUsers_Listview
            $reportName = "AAD Users"
		}
		"AAD Deleted Users"{
			$lisview = $uiHash.AADDeletedUsers_Listview
            $reportName = "AAD Deleted Users"
		}
		"AAD External Users"{
			$lisview = $uiHash.AADExternalUsers_Listview
            $reportName = "AAD External Users"
		}
		"AAD Contacts"{
			$lisview = $uiHash.AADContacts_Listview
            $reportName = "AAD Contacts"
		}
		"AAD Groups"{
			$lisview = $uiHash.AADGroups_Listview
            $reportName = "AAD Groups"
		}
		"AAD Licenses"{
			$lisview = $uiHash.AADLicenses_Listview
            $reportName = "AAD Licenses"
		}
		"AAD Domains"{
			$lisview = $uiHash.AADDomains_Listview
            $reportName = "AAD Domains"
		}
		"Exchange Mailboxes"{
			$lisview = $uiHash.ExchangeMailboxes_Listview
            $reportName = "Exchange Mailboxes"
		}
		"Exchange Archives"{
			$lisview = $uiHash.ExchangeArchives_Listview
            $reportName = "Exchange Archives"
		}
		"Exchange Groups"{
			$lisview = $uiHash.ExchangeGroups_Listview
            $reportName = "Exchange Groups"
		}
		"SharePoint Sites"{
			$lisview = $uiHash.SharePointSites_Listview
            $reportName = "SharePoint Sites"
		}
		"SharePoint Webs"{
			$lisview = $uiHash.SharePointWebs_Listview
            $reportName = "SharePoint Webs"
		}
	}
	
	Write-Debug ("Listview count {0}" -f $lisview.ItemsSource.count)
	
	Switch ($uiHash.ReportComboBox.SelectedItem.Text) 
    {
		"CSV Report" {
			If ($lisview.ItemsSource.count -gt 0) {
				$uiHash.StatusTextBox.Foreground = "Black"
				$savedreport = Join-Path $reportpath "$reportName-$($today).csv"
				$lisview.Items | Export-Csv $savedreport -NoTypeInformation
				$uiHash.StatusTextBox.Text = "Report saved to $savedreport"
			} Else {
				$uiHash.StatusTextBox.Foreground = "Red"
				$uiHash.StatusTextBox.Text = "No report to create!"         
			}
		}
		"Excel Report" {
			If ($lisview.ItemsSource.count -gt 0) {
				$uiHash.StatusTextBox.Foreground = "Black"
				$savedreport = Join-Path $reportpath "$reportName-$($today).csv"
				$lisview.Items | Export-Csv $savedreport -NoTypeInformation
				$uiHash.StatusTextBox.Text = "Report saved to $savedreport"
			} Else {
				$uiHash.StatusTextBox.Foreground = "Red"
				$uiHash.StatusTextBox.Text = "No report to create!"         
			}
		}
		"HTML Report" {
			If ($lisview.ItemsSource.count -gt 0) { 
				$uiHash.StatusTextBox.Foreground = "Black"
				$savedreport = Join-Path $reportpath "HTMLReport_$($today).html"
				
				$HTMLReport = $lisview.Items | ConvertTo-Html `
					-As Table `
					-Fragment `
					-PreContent '<h1>HTMLReport</h1>' | 
							Out-file $savedreport
				$uiHash.StatusTextBox.Text = "Report saved to $savedreport"
			} Else {
				$uiHash.StatusTextBox.Foreground = "Red"
				$uiHash.StatusTextBox.Text = "No report to create!"         
			}    			
		}
		"Full HTML Report" {
			$uiHash.StatusTextBox.Foreground = "Black"
			$savedreport = Join-Path $reportpath "FullHTMLReport_$($today).html"
			
			#HTML Variables
			$articleArray = @()
			$buttonArray = @()			
			
			#Azure Active Directory Users
			$AADUsers_Article +=  "
				<div id=`"div-1`" style=`"display:none;`">
				<h1>Azure Active Directory Users</h1>
			"				
			
			if ($uiHash.AADUsers_Listview.ItemsSource.count -gt 0){
				$AADUsers_Article += $uiHash.AADUsers_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$AADUsers_Article +=  "<p>There are no entries for this query</p>"
			}	
			$AADUsers_Article +=  "</div>"
			$AADUsers_Button = "<li><button id='button_1' onclick=`"showDiv('1')`">Azure Active Directory Users</button></li>"
			
			#Azure Active Directory Deleted Users
			$AADDeletedUsers_Article +=  "
				<div id=`"div-2`" style=`"display:none;`">
				<h1>Azure Active Directory Deleted Users</h1>
			"				
			
			if ($uiHash.AADDeletedUsers_Listview.ItemsSource.count -gt 0){
				$AADDeletedUsers_Article += $uiHash.AADDeletedUsers_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$AADDeletedUsers_Article +=  "<p>There are no entries for this query</p>"
			}	
			$AADDeletedUsers_Article +=  "</div>"
			$AADDeletedUsers_Button = "<li><button id='button_2' onclick=`"showDiv('2')`">Azure Active Directory Deleted Users</button></li>"
			
			
			#Azure Active Directory External Users
			$AADExternalUsers_Article +=  "
				<div id=`"div-3`" style=`"display:none;`">
				<h1>Azure Active Directory External Users</h1>
			"				
			
			if ($uiHash.AADExternalUsers_Listview.ItemsSource.count -gt 0){
				$AADExternalUsers_Article += $uiHash.AADExternalUsers_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$AADExternalUsers_Article +=  "<p>There are no entries for this query</p>"
			}	
			$AADExternalUsers_Article +=  "</div>"
			$AADExternalUsers_Button = "<li><button id='button_3' onclick=`"showDiv('3')`">Azure Active Directory External Users</button></li>"
			
			#Azure Active Directory Contacts
			$AADContacts_Article +=  "
				<div id=`"div-4`" style=`"display:none;`">
				<h1>Azure Active Directory Contacts</h1>
			"				
			
			if ($uiHash.AADContacts_Listview.ItemsSource.count -gt 0){
				$AADContacts_Article += $uiHash.AADContacts_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$AADContacts_Article +=  "<p>There are no entries for this query</p>"
			}	
			$AADContacts_Article +=  "</div>"
			$AADContacts_Button = "<li><button id='button_4' onclick=`"showDiv('4')`">Azure Active Directory Contacts</button></li>"
			
			#Azure Active Directory Groups
			$AADGroups_Article +=  "
				<div id=`"div-5`" style=`"display:none;`">
				<h1>Azure Active Directory Groups</h1>
			"				
			
			if ($uiHash.AADGroups_Listview.ItemsSource.count -gt 0){
				$AADGroups_Article += $uiHash.AADGroups_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$AADGroups_Article +=  "<p>There are no entries for this query</p>"
			}	
			$AADGroups_Article +=  "</div>"
			$AADGroups_Button = "<li><button id='button_5' onclick=`"showDiv('5')`">Azure Active Directory Groups</button></li>"
			
			#Azure Active Directory Licenses
			$AADLicenses_Article +=  "
				<div id=`"div-6`" style=`"display:none;`">
				<h1>Azure Active Directory Licenses</h1>
			"				
			
			if ($uiHash.AADLicenses_Listview.ItemsSource.count -gt 0){
				$AADLicenses_Article += $uiHash.AADLicenses_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$AADLicenses_Article +=  "<p>There are no entries for this query</p>"
			}	
			$AADLicenses_Article +=  "</div>"
			$AADLicenses_Button = "<li><button id='button_6' onclick=`"showDiv('6')`">Azure Active Directory Licenses</button></li>"
			
			#Azure Active Directory Domains
			$AADDomains_Article +=  "
				<div id=`"div-7`" style=`"display:none;`">
				<h1>Azure Active Directory Domains</h1>
			"				
			
			if ($uiHash.AADDomains_Listview.ItemsSource.count -gt 0){
				$AADDomains_Article += $uiHash.AADDomains_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$AADDomains_Article +=  "<p>There are no entries for this query</p>"
			}	
			$AADDomains_Article +=  "</div>"
			$AADDomains_Button = "<li><button id='button_7' onclick=`"showDiv('7')`">Azure Active Directory Domains</button></li>"
			
			#Exchange Mailboxes
			$ExchangeMailboxes_Article +=  "
				<div id=`"div-8`" style=`"display:none;`">
				<h1>Exchange Mailboxes</h1>
			"				
			
			if ($uiHash.ExchangeMailboxes_Listview.ItemsSource.count -gt 0){
				$ExchangeMailboxes_Article += $uiHash.ExchangeMailboxes_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$ExchangeMailboxes_Article +=  "<p>There are no entries for this query</p>"
			}	
			$ExchangeMailboxes_Article +=  "</div>"
			$ExchangeMailboxes_Button = "<li><button id='button_8' onclick=`"showDiv('8')`">Exchange Mailboxes</button></li>"
			
			#Exchange Archives
			$ExchangeMailboxes_Article +=  "
				<div id=`"div-9`" style=`"display:none;`">
				<h1>Exchange Archives</h1>
			"				
			
			if ($uiHash.ExchangeArchives_Listview.ItemsSource.count -gt 0){
				$ExchangeArchives_Article += $uiHash.ExchangeArchives_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$ExchangeArchives_Article +=  "<p>There are no entries for this query</p>"
			}	
			$ExchangeArchives_Article +=  "</div>"
			$ExchangeArchives_Button = "<li><button id='button_9' onclick=`"showDiv('9')`">Exchange Archives</button></li>"
			
			#Exchange Groups
			$ExchangeGroups_Article +=  "
				<div id=`"div-10`" style=`"display:none;`">
				<h1>Exchange Groups</h1>
			"				
			
			if ($uiHash.ExchangeGroups_Listview.ItemsSource.count -gt 0){
				$ExchangeGroups_Article += $uiHash.ExchangeGroups_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$ExchangeGroups_Article +=  "<p>There are no entries for this query</p>"
			}	
			$ExchangeGroups_Article +=  "</div>"
			$ExchangeGroups_Button = "<li><button id='button_10' onclick=`"showDiv('10')`">Exchange Groups</button></li>"
			
			#SharePoint Sites
			$SharePointSites_Article +=  "
				<div id=`"div-11`" style=`"display:none;`">
				<h1>SharePoint Sites</h1>
			"				
			
			if ($uiHash.SharePointSites_Listview.ItemsSource.count -gt 0){
				$SharePointSites_Article += $uiHash.SharePointSites_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$SharePointSites_Article +=  "<p>There are no entries for this query</p>"
			}	
			$SharePointSites_Article +=  "</div>"
			$SharePointSites_Button = "<li><button id='button_11' onclick=`"showDiv('11')`">SharePoint Sites</button></li>"
			
			#SharePoint Webs
			$SharePointWebs_Article +=  "
				<div id=`"div-12`" style=`"display:none;`">
				<h1>SharePoint Webs</h1>
			"				
			
			if ($uiHash.SharePointWebs_Listview.ItemsSource.count -gt 0){
				$SharePointWebs_Article += $uiHash.SharePointWebs_Listview.items |
				ConvertTo-Html `
					-As Table `
					-Fragment | 
				Out-String
				
			}
			else{
				$SharePointWebs_Article +=  "<p>There are no entries for this query</p>"
			}	
			$SharePointWebs_Article +=  "</div>"
			$SharePointWebs_Button = "<li><button id='button_12' onclick=`"showDiv('12')`">SharePoint Webs</button></li>"
			
			#fill article erray
			$articleArray += $AADUsers_Article
			$articleArray += $AADDeletedUsers_Article
			$articleArray += $AADExternalUsers_Article
			$articleArray += $AADContacts_Article
			$articleArray += $AADGroups_Article
			$articleArray += $AADLicenses_Article
			$articleArray += $AADDomains_Article
			$articleArray += $ExchangeMailboxes_Article
			$articleArray += $ExchangeArchives_Article
			$articleArray += $ExchangeGroups_Article
			$articleArray += $SharePointSites_Article
			$articleArray += $SharePointWebs_Article
			
			#fill button array
			$buttonArray += $AADUsers_button
			$buttonArray += $AADDeletedUsers_button
			$buttonArray += $AADExternalUsers_button
			$buttonArray += $AADContacts_button
			$buttonArray += $AADGroups_button
			$buttonArray += $AADLicenses_button
			$buttonArray += $AADDomains_button
			$buttonArray += $ExchangeMailboxes_button
			$buttonArray += $ExchangeArchives_button
			$buttonArray += $ExchangeGroups_button
			$buttonArray += $SharePointSites_button
			$buttonArray += $SharePointWebs_button

			create-Fullhtml -path $savedreport						
						
			$uiHash.StatusTextBox.Text = "Report saved to $savedreport"
		} Else {
			$uiHash.StatusTextBox.Foreground = "Red"
			$uiHash.StatusTextBox.Text = "No report to create!"         
		}    			
	}
}

Function create-Fullhtml(){
    #parameters
    param([string]$path)

	#Head
	$head = "
		<html xmlns=`"http://www.w3.org/1999/xhtml`">
			<head>
				<style>
					@charset `"UTF-8`";
		  
					div.container {
						width: 100%;
						border: 1px solid gray;
					}
					
					header, footer {
						padding: 1em;
						color: white;
						background-color: black;
						clear: left;
						text-align: center;
					}
					
					nav {
						float: left;
						max-width: 160px;
						margin: 0;
						padding: 1em;
					}

					nav button{
						margin-bottom:10px;
						width:100%;
						text-align:center;
					}	

					nav ul {
						list-style-type: none;
						padding: 0;
					}		

					article {
						margin-left: 190px;
						min-width:600px;
						min-height: 600px;
						border-left: 1px solid gray;
						padding: 1em;
					}
					
					th{
						border:1px Solid Black;
						border-Collapse:collapse; 
						background-color:lightblue;
					}
					
					th{
						border:1px Solid Black;
						border-Collapse:collapse; 
					}

					.visibleClass {
						display: block !important;
					}		
					
				</style>
			</head>	
	"
	
	#Header
	$header = "
		<h1>Office 365 Inventory Tool</h1>
	"
	
	#Navigation
	$Navigation = "<ul>"
	
	foreach($item in $buttonArray){
		$Navigation += $item
	}
	$Navigation += "<li><button id='button_showAll' onclick=`"showAll()`">Show all</button></li>"
	$Navigation += "</ul>"

	
	#Article
	$article = ""
	foreach($item in $articleArray){
		$article += $item
	}
	
	#Footer
	$Footer = "
		Copyright &copy; SharePointFire.com
	"
	
	$JavaScript = "
		<script>
			function showDiv(data) {
				var div = `"div-`" + data
				var x = document.getElementById(div)
				x.className = `"visibleClass`";
				hideElement(20, data);
			}
			
			function hideElement(total, active) {
				for (i = 1; i <= total; i++) {
					var div = `"div-`" + i
					var y = document.getElementById(div)
					if (i != active)
						y.className = `"other`";
				}
			}
			
			function showAll() {
				for (i = 1; i <= 20; i++) {
					var div = `"div-`" + i
					var y = document.getElementById(div)
					y.className = `"visibleClass`";
				}
			}
		</script>
	"
	
	#Full HTML
	$HTML = "
		$($Head)		
		<body class=`"Inventory`">
			<div class=`"container`">
				<header>
					$($Header)
				</header>
						
				<nav>
					$($Navigation)
				</nav>

				<article>
					$($article)
				</article>
						
				<footer>
					$($footer)
				</footer>
			</div>
			
		$($JavaScript)	
		</body>
		</html>
	" 
	add-content $HTML -path $Path
} 