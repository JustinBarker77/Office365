#Format and display errors
Function Get-Error {
    Process {
        ForEach ($err in $error) {
            Switch ($err) {
                {$err -is [System.Management.Automation.ErrorRecord]} {
                        $hash = @{
                        Category = $err.categoryinfo.Category
                        Activity = $err.categoryinfo.Activity
                        Reason = $err.categoryinfo.Reason
                        Type = $err.GetType().ToString()
                        Exception = ($err.exception -split ": ")[1]
                        QualifiedError = $err.FullyQualifiedErrorId
                        CharacterNumber = $err.InvocationInfo.OffsetInLine
                        LineNumber = $err.InvocationInfo.ScriptLineNumber
                        Line = $err.InvocationInfo.Line
                        TargetObject = $err.TargetObject
                        }
                    }               
                Default {
                    $hash = @{
                        Category = $err.errorrecord.categoryinfo.category
                        Activity = $err.errorrecord.categoryinfo.Activity
                        Reason = $err.errorrecord.categoryinfo.Reason
                        Type = $err.GetType().ToString()
                        Exception = ($err.errorrecord.exception -split ": ")[1]
                        QualifiedError = $err.errorrecord.FullyQualifiedErrorId
                        CharacterNumber = $err.errorrecord.InvocationInfo.OffsetInLine
                        LineNumber = $err.errorrecord.InvocationInfo.ScriptLineNumber
                        Line = $err.errorrecord.InvocationInfo.Line                    
                        TargetObject = $err.errorrecord.TargetObject
                    }               
                }                        
            }
        $object = New-Object PSObject -Property $hash
        $object.PSTypeNames.Insert(0,'ErrorInformation')
        $object
        }
    }
}

#GUI changes when starting action
Function Start-Action{
	param(
		$message
	)
	
	$uiHash.StatusTextBox.Foreground = "Black"
	$uiHash.StatusTextBox.Text = $message	
	$uiHash.RunButton.IsEnabled = $False
	$uiHash.StartImage.Source = "$pwd\Images\Start_Locked.ico"
	$uiHash.RunAllButton.IsEnabled = $False
	$uiHash.StartAllImage.Source = "$pwd\Images\StartAll_Locked.ico"
}

#GUI changes when action has finished
Function End-Action{
	param(
		$message
	)
	
	$uiHash.StatusTextBox.Foreground = "Green"
	$uiHash.StatusTextBox.Text = $message
	$uiHash.RunButton.IsEnabled = $True
	$uiHash.StartImage.Source = "$pwd\Images\Start.ico"
	$uiHash.RunAllButton.IsEnabled = $True
	$uiHash.StartAllImage.Source = "$pwd\Images\StartAll.ico"
}

#GUI changes when action generated error
Function Error-Action{
	param(
		$message
	)
	
	$uiHash.StatusTextBox.Foreground = "red"
	$uiHash.StatusTextBox.Text = $message
	$uiHash.RunButton.IsEnabled = $True
	$uiHash.StartImage.Source = "$pwd\Images\Start.ico"
	$uiHash.RunAllButton.IsEnabled = $True
	$uiHash.StartAllImage.Source = "$pwd\Images\StartAll.ico"
}
