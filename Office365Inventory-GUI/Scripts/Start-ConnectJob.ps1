function Start-ConnectJob(){
	param(
		$credential
	)

	if($credential){
		$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
			start-action -message "Connecting to Office 365...Please Wait"
		})
				
		try{
			#connect to Office 365		
			Connect-MsolService -credential $credential	
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				$uiHash.ConnectedAAD_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				error-action -message $_.exception.Message
				$uiHash.ConnectedAAD_Image.Source = "$pwd\Images\Check_Error.ico"
			})
			$errored = $true
		}
			
		try{
			#connect to Exchange Online
			$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Credential -Authentication Basic –AllowRedirection
			
			Import-PSSession $Session
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				$uiHash.ConnectedExchange_Image.Source = "$pwd\Images\Check_Okay.ico"
			})
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				error-action -message $_.exception.Message
				$uiHash.ConnectedExchange_Image.Source = "$pwd\Images\Check_Error.ico"
			})
			$errored = $true
		}
		
		try{
			#first get SharePoint URL by tenant domain
			$FQDN = Get-MsolDomain | where{$_.name -like "*.onmicrosoft.com"}
			$name = ($fqdn.name).split(".")[0]
			
			#import SharePoint Online Module
			Import-Module Microsoft.Online.SharePoint.PowerShell -WarningAction:SilentlyContinue

			#connect to SharePoint Online
			Connect-SPOService -url "https://$($Name)-admin.sharepoint.com" -credential $Credential
			
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				$uiHash.ConnectedSharePoint_Image.Source = "$pwd\Images\Check_Okay.ico"
			})		
		}
		catch{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				error-action -message $_.exception.Message
				$uiHash.ConnectedSharePoint_Image.Source = "$pwd\Images\Check_Error.ico"
			})
			$errored = $true
		}
		
		if(!$errored){
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				$uiHash.ConnectImage.Source = "$pwd\Images\Connected.ico"
				end-action -message "Connected Succesfully"
			})
		}
		else{
			$uiHash.Window.Dispatcher.Invoke("Normal",[action]{
				$uiHash.ConnectImage.Source = "$pwd\Images\Connected.ico"
				error-action -message "Error connecting to Office 365"
			})
		}
	}
}