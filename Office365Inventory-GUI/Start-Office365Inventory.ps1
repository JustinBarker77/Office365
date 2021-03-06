<#
.SYNOPSIS
        This GUI will let you get all kinds of information out of Office 365
.DESCRIPTION
        Get all kinds of information from Office 365
.LINK
    Maarten Peeters Blog
        - http://www.sharepointfire.com
.NOTES
        Author:     M. Peeters
        Date:       05/06/2017
        PS Ver.:    5.0
        Script Ver: 1.0

        Change log:
            v0.1 Created GUI
			v0.2 Added SharePoint, Exchange and AAD bits
			v0.3 Edited Helpfiles and about
			v0.4 Added runspaces functionality
			v1.0 Cleaned up script and provided comments
#>

#region Synchronized Collections
$uiHash = [hashtable]::Synchronized(@{})
#endregion

#region Startup Checks and configurations
#Validate user is an Administrator
Write-Verbose "Checking Administrator credentials"
If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(`
    [Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "You are not running this as an Administrator!`nRe-running script and will prompt for administrator credentials."
    Start-Process -Verb "Runas" -File PowerShell.exe -Argument "-STA -noprofile -file `"$($myinvocation.mycommand.definition)`""
    Break
}

#Ensure that we are running the GUI from the correct location
Set-Location $(Split-Path $MyInvocation.MyCommand.Path)
$Global:Path = $(Split-Path $MyInvocation.MyCommand.Path)
Write-Debug "Current location: $Path"

#Determine if this instance of PowerShell can run WPF 
Write-Verbose "Checking the apartment state"
If ($host.Runspace.ApartmentState -ne "STA") {
    Write-Warning "This script must be run in PowerShell started using -STA switch!`nScript will attempt to open PowerShell in STA and run re-run script."
    Start-Process -File PowerShell.exe -Argument "-STA -noprofile -WindowStyle hidden -file `"$($myinvocation.mycommand.definition)`""
    Break
}

#Load Required Assemblies
Add-Type –assemblyName PresentationFramework
Add-Type –assemblyName PresentationCore
Add-Type –assemblyName WindowsBase
Add-Type –assemblyName Microsoft.VisualBasic
Add-Type –assemblyName System.Windows.Forms

#DotSource Help script
. ".\HelpFiles\HelpOverview.ps1"

#DotSource About script
. ".\HelpFiles\About.ps1"

#DotSource shared functions
. ".\Scripts\SharedFunctions.ps1"

#DotSource shared functions
. ".\Scripts\Start-Report.ps1"
#endregion

#region GUI
[xml]$xaml = @"
<Window
    xmlns='http://schemas.microsoft.com/winfx/2006/xaml/presentation'
    xmlns:x='http://schemas.microsoft.com/winfx/2006/xaml'
    x:Name='MainWindow' Title='PowerShell Office 365 Inventory' WindowStartupLocation = 'CenterScreen' 
    Width = '880' Height = '575' ShowInTaskbar = 'True'>
    <Window.Background>
        <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
            <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
            <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
        </LinearGradientBrush>
    </Window.Background> 
    <Window.Resources>        
        <DataTemplate x:Key="HeaderTemplate">
            <DockPanel>
                <TextBlock FontSize="10" Foreground="Green" FontWeight="Bold" >
                    <TextBlock.Text>
                        <Binding/>
                    </TextBlock.Text>
                </TextBlock>
            </DockPanel>
        </DataTemplate>            
    </Window.Resources>    
    <Grid x:Name = 'Grid' ShowGridLines = 'false'>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = '*'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
            <RowDefinition Height = 'Auto'/>
        </Grid.RowDefinitions>    
        <Menu Width = 'Auto' HorizontalAlignment = 'Stretch' Grid.Row = '0'>
        <Menu.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>
        </Menu.Background>
            <MenuItem x:Name = 'FileMenu' Header = '_File'>
				<MenuItem x:Name = 'ConnectMenu' Header = '_Connect' ToolTip = 'Initiate connect operation' InputGestureText ='F4'> </MenuItem>
                <MenuItem x:Name = 'RunMenu' Header = '_Run' ToolTip = 'Initiate Run operation' InputGestureText ='F5'> </MenuItem>
				<MenuItem x:Name = 'RunAllMenu' Header = '_Run All' ToolTip = 'Initiate Run All operation' InputGestureText ='F6'> </MenuItem>
                <MenuItem x:Name = 'GenerateReportMenu' Header = 'Generate Report' ToolTip = 'Generate Report' InputGestureText ='F8'/>
                <Separator />            
                <MenuItem x:Name = 'ExitMenu' Header = 'E_xit' ToolTip = 'Exits the tool.' InputGestureText ='Ctrl+E'/>
            </MenuItem>  		
            <MenuItem x:Name = 'HelpMenu' Header = '_Help'>
                <MenuItem x:Name = 'AboutMenu' Header = '_About' ToolTip = 'Show the current version and other information.'> </MenuItem>
                <MenuItem x:Name = 'HelpFileMenu' Header = 'Office 365 Inventory tool _Help' 
                ToolTip = 'Displays a help file to use this GUI.' InputGestureText ='F1'> </MenuItem>
				<Separator/>
                <MenuItem x:Name = 'ViewErrorMenu' Header = 'View ErrorLog' ToolTip = 'Get error log.'/> 
				<MenuItem x:Name = 'ClearErrorMenu' Header = 'Clear ErrorLog' ToolTip = 'Clears error log.'> </MenuItem>				
            </MenuItem>            
        </Menu>
        <ToolBarTray Grid.Row = '1' Grid.Column = '0'>
        <ToolBarTray.Background>
            <LinearGradientBrush StartPoint='0,0' EndPoint='0,1'>
                <LinearGradientBrush.GradientStops> <GradientStop Color='#C4CBD8' Offset='0' /> <GradientStop Color='#E6EAF5' Offset='0.2' /> 
                <GradientStop Color='#CFD7E2' Offset='0.9' /> <GradientStop Color='#C4CBD8' Offset='1' /> </LinearGradientBrush.GradientStops>
            </LinearGradientBrush>        
        </ToolBarTray.Background>
            <ToolBar Background = 'Transparent' Band = '1' BandIndex = '1'>
				<Button x:Name = 'ConnectButton' Width = 'Auto' ToolTip = 'Connect.'>
                    <Image x:Name = 'ConnectImage' Source = '$Pwd\Images\Connect.ico' Height = '16' Width = '16'/>
                </Button>
                <Separator Background = 'Black'/>
                <Button x:Name = 'RunButton' Width = 'Auto' ToolTip = 'Performs action.'>
                    <Image x:Name = 'StartImage' Source = '$Pwd\Images\Start.ico' Height = '16' Width = '16'/>
                </Button>
				<Button x:Name = 'RunAllButton' Width = 'Auto' ToolTip = 'Performs all actions.'>
                    <Image x:Name = 'StartAllImage' Source = '$Pwd\Images\StartAll.ico' Height = '16' Width = '16'/>
                </Button>
                <Separator Background = 'Black'/>              
                <Button x:Name = 'GenerateReportButton' Width = 'Auto' ToolTip = 'Generates a report based on user selection.'>
                    <Image Source = '$pwd\Images\Gen_Report.ico' Height = '16' Width = '16'/>
                </Button>            
                <ComboBox x:Name = 'ReportComboBox' Width = 'Auto' IsReadOnly = 'True' SelectedIndex = '0'>
                    <TextBlock> CSV Report </TextBlock>
					<TextBlock> Full Excel Report </TextBlock>
                    <TextBlock> HTML Report </TextBlock>
					<TextBlock> Full HTML Report </TextBlock>
                </ComboBox>              
                <Separator Background = 'Black'/>
            </ToolBar>           
        </ToolBarTray>
		<TabControl x:Name='Tabs' Grid.Row = '2' Grid.Column = '0'>
			<TabItem Header="Home">
				<Grid Grid.Row = '2' Grid.Column = '0'>
					<Grid.RowDefinitions>
						<RowDefinition Height = 'auto'/>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions>
					<Label Grid.Row = '0' Grid.Column = '0'>Please verify the help file added for any information regarding this tool. You can verify in the below tabel if an action has been run succesfully.</Label>
					<Grid Grid.Row = '1' Grid.Column = '0' ShowGridLines = 'false'>
						<Grid.ColumnDefinitions>
							<ColumnDefinition Width="25"/>
							<ColumnDefinition Width="auto"/>
							<ColumnDefinition Width="25"/>
							<ColumnDefinition Width="25"/>
							<ColumnDefinition Width="auto"/>
							<ColumnDefinition Width="25"/>
							<ColumnDefinition Width="25"/>
							<ColumnDefinition Width="auto"/>
							<ColumnDefinition Width="25"/>
							<ColumnDefinition Width="25"/>
							<ColumnDefinition Width="auto"/>
							<ColumnDefinition Width="25"/>
							<ColumnDefinition Width="25"/>
							<ColumnDefinition Width="auto"/>
						</Grid.ColumnDefinitions>
						<Grid.RowDefinitions>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
							<RowDefinition Height = 'auto'/>
						</Grid.RowDefinitions>
						<Image x:Name = 'ConnectedAAD_Image' Grid.Column = '0' Grid.Row = '0' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '1' Grid.Row = '0'>Connected AAD</Label>
						<Image x:Name = 'ConnectedExchange_Image' Grid.Column = '0' Grid.Row = '1' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '1' Grid.Row = '1'>Connected Exchange</Label>
						<Image x:Name = 'ConnectedSharePoint_Image' Grid.Column = '0' Grid.Row = '2' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '1' Grid.Row = '2'>Connected SharePoint</Label>
						<Image x:Name = 'AADUsers_Image' Grid.Column = '3' Grid.Row = '0' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '4' Grid.Row = '0'>AAD Users</Label>
						<Image x:Name = 'AADDeletedUsers_Image' Grid.Column = '3' Grid.Row = '1' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '4' Grid.Row = '1'>AAD Deleted Users</Label>
						<Image x:Name = 'AADExternalUsers_Image' Grid.Column = '3' Grid.Row = '2' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '4' Grid.Row = '2'>AAD External Users</Label>
						<Image x:Name = 'AADContacts_Image' Grid.Column = '3' Grid.Row = '3' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '4' Grid.Row = '3'>AAD Contacts</Label>
						<Image x:Name = 'AADGroups_Image' Grid.Column = '3' Grid.Row = '4' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '4' Grid.Row = '4'>AAD Groups</Label>
						<Image x:Name = 'AADLicenses_Image' Grid.Column = '3' Grid.Row = '5' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '4' Grid.Row = '5'>AAD Licenses</Label>
						<Image x:Name = 'AADDomains_Image' Grid.Column = '3' Grid.Row = '6' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '4' Grid.Row = '6'>AAD Domains</Label>
						<Image x:Name = 'ExchangeMailboxes_Image' Grid.Column = '6' Grid.Row = '0' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '7' Grid.Row = '0'>Exchange Mailboxes</Label>
						<Image x:Name = 'ExchangeArchives_Image' Grid.Column = '6' Grid.Row = '1' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '7' Grid.Row = '1'>Exchange Archives</Label>
						<Image x:Name = 'ExchangeGroups_Image' Grid.Column = '6' Grid.Row = '2' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '7' Grid.Row = '2'>Exchange Groups</Label>
						<Image x:Name = 'SharePointSites_Image' Grid.Column = '9' Grid.Row = '0' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '10' Grid.Row = '0'>SharePoint Sites</Label>
						<Image x:Name = 'SharePointWebs_Image' Grid.Column = '9' Grid.Row = '1' Source = '$pwd\Images\Check_Waiting.ico' Height = '16' Width = '16'/>
						<Label Grid.Column = '10' Grid.Row = '1'>SharePoint Webs</Label>
					</Grid>
				</Grid>
			</TabItem>
			<TabItem Header="AAD Users">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Azure Active Directory Users" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'AADUsers_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Azure Active Directory Users.'>
							<ListView.View>
								<GridView x:Name = 'AADUsers_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'AADUser_DisplayName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_DisplayName}' Header='Display Name'/> 
									<GridViewColumn x:Name = 'AADUser_FirstName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_FirstName}' Header='FirstName'/>
									<GridViewColumn x:Name = 'AADUser_LastName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_LastName}' Header='LastName'/>
									<GridViewColumn x:Name = 'AADUser_UserPrincipalName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_UserPrincipalName}' Header='UPN'/>
									<GridViewColumn x:Name = 'AADUser_Title' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_Title}' Header='Title'/>
									<GridViewColumn x:Name = 'AADUser_Department' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_Department}' Header='Department'/> 
									<GridViewColumn x:Name = 'AADUser_Office' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_Office}' Header='Office'/>
									<GridViewColumn x:Name = 'AADUser_PhoneNumber' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_PhoneNumber}' Header='Phone number'/>
									<GridViewColumn x:Name = 'AADUser_MobilePhone' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_MobilePhone}' Header='Mobile number'/>
									<GridViewColumn x:Name = 'AADUser_CloudAnchor' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_CloudAnchor}' Header='ImmutableId'/>
									<GridViewColumn x:Name = 'AADUser_IsLicensed' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_IsLicensed}' Header='Is licensed'/>
									<GridViewColumn x:Name = 'AADUser_Licenses' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADUser_Licenses}' Header='Licenses'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="AAD Deleted Users">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Azure Active Directory Deleted Users" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'AADDeletedUsers_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Azure Active Directory Users.'>
							<ListView.View>
								<GridView x:Name = 'AADDeletedUsers_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'AADDeletedUser_SignInName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADDeletedUser_SignInName}' Header='Sign in Name'/> 
									<GridViewColumn x:Name = 'AADDeletedUser_UserPrincipalName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADDeletedUser_UserPrincipalName}' Header='User Principal Name'/>
									<GridViewColumn x:Name = 'AADDeletedUser_DisplayName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADDeletedUser_DisplayName}' Header='Display Name'/>
									<GridViewColumn x:Name = 'AADDeletedUser_SoftDeletionTimestamp' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADDeletedUser_SoftDeletionTimestamp}' Header='When Deleted'/>
									<GridViewColumn x:Name = 'AADDeletedUser_IsLicensed' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADDeletedUser_IsLicensed}' Header='Is licensed'/>
									<GridViewColumn x:Name = 'AADDeletedUser_Licenses' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADDeletedUser_Licenses}' Header='Licenses'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="AAD External Users">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Azure Active Directory External Users" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'AADExternalUsers_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Azure Active Directory Users.'>
							<ListView.View>
								<GridView x:Name = 'AADExternalUsers_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'AADExternalUser_SignInName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADExternalUser_SignInName}' Header='Sign in Name'/> 
									<GridViewColumn x:Name = 'AADExternalUser_UserPrincipalName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADExternalUser_UserPrincipalName}' Header='User Principal Name'/>
									<GridViewColumn x:Name = 'AADExternalUser_DisplayName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADExternalUser_DisplayName}' Header='Display Name'/>
									<GridViewColumn x:Name = 'AADExternalUser_WhenCreated' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADExternalUser_WhenCreated}' Header='When Created'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="AAD Contacts">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Azure Active Directory Contacts" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'AADContacts_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Azure Active Directory Contacts.'>
							<ListView.View>
								<GridView x:Name = 'AADContacts_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'AADContacts_DisplayName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADContacts_DisplayName}' Header='Display Name'/> 
									<GridViewColumn x:Name = 'AADContacts_EmailAddress' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADContacts_EmailAddress}' Header='Email Address'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="AAD Groups">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Azure Active Directory Groups" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'AADGroups_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Azure Active Directory Groups.'>
							<ListView.View>
								<GridView x:Name = 'AADGroups_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'AADGroup_GroupType' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADGroup_GroupType}' Header='Group Type'/>
									<GridViewColumn x:Name = 'AADGroup_DisplayName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADGroup_DisplayName}' Header='Display Name'/> 
									<GridViewColumn x:Name = 'AADGroup_EmailAddress' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADGroup_EmailAddress}' Header='Email Address'/>
									<GridViewColumn x:Name = 'AADGroup_ValidationStatus' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADGroup_ValidationStatus}' Header='Status'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="AAD Licenses">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Azure Active Directory Licenses" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'AADLicenses_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Azure Active Directory Licenses.'>
							<ListView.View>
								<GridView x:Name = 'AADLicenses_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'AADLicenses_AccountSkuID' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADLicenses_AccountSkuID}' Header='ID'/>
									<GridViewColumn x:Name = 'AADLicenses_ActiveUnits' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADLicenses_ActiveUnits}' Header='Total'/> 
									<GridViewColumn x:Name = 'AADLicenses_ConsumedUnits' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADLicenses_ConsumedUnits}' Header='Used'/>
									<GridViewColumn x:Name = 'AADLicenses_LockedOutUnits' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADLicenses_LockedOutUnits}' Header='Locked'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="AAD Domains">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Azure Active Directory Domains" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'AADDomains_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Azure Active Directory Domains.'>
							<ListView.View>
								<GridView x:Name = 'AADDomains_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'AADDomains_Name' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADDomains_Name}' Header='Name'/>
									<GridViewColumn x:Name = 'AADDomains_Status' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADDomains_Status}' Header='Status'/> 
									<GridViewColumn x:Name = 'AADDomains_Authentication' Width = 'Auto' DisplayMemberBinding = '{Binding Path = AADDomains_Authentication}' Header='Authentication'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="Exchange Mailboxes">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Exchange Mailboxes" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '12' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'ExchangeMailboxes_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Exchange Mailboxes.'>
							<ListView.View>
								<GridView x:Name = 'ExchangeMailboxes_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'ExchangeMailboxes_DisplayName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_DisplayName}' Header='Display Name'/>
									<GridViewColumn x:Name = 'ExchangeMailboxes_Alias' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_Alias}' Header='Alias'/> 
									<GridViewColumn x:Name = 'ExchangeMailboxes_PrimarySMTPAddress' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_PrimarySMTPAddress}' Header='Primary Mail Address'/>
									<GridViewColumn x:Name = 'ExchangeMailboxes_ItemCount' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_ItemCount}' Header='Item Count'/> 
									<GridViewColumn x:Name = 'ExchangeMailboxes_TotalItemSize' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_TotalItemSize}' Header='Total Item Size'/> 
									<GridViewColumn x:Name = 'ExchangeMailboxes_ArchiveStatus' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_ArchiveStatus}' Header='Archive Status'/>
									<GridViewColumn x:Name = 'ExchangeMailboxes_UsageLocation' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_UsageLocation}' Header='Usage Location'/> 
									<GridViewColumn x:Name = 'ExchangeMailboxes_WhenMailboxCreated' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_WhenMailboxCreated}' Header='Created'/>
									<GridViewColumn x:Name = 'ExchangeMailboxes_LastLogonTime' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_LastLogonTime}' Header='Last logon'/>
									<GridViewColumn x:Name = 'ExchangeMailboxes_RecipientTypeDetails' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_RecipientTypeDetails}' Header='RecipientTypeDetails'/>
									<GridViewColumn x:Name = 'ExchangeMailboxes_LegacyExchangeDN' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_LegacyExchangeDN}' Header='Legacy DN'/>
									<GridViewColumn x:Name = 'ExchangeMailboxes_ProxyAddresses' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeMailboxes_ProxyAddresses}' Header='Email Addresses'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="Exchange Archives">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Exchange Archives" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'ExchangeArchives_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Exchange Archives.'>
							<ListView.View>
								<GridView x:Name = 'ExchangeArchives_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'ExchangeArchives_DisplayName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeArchives_DisplayName}' Header='Display Name'/>
									<GridViewColumn x:Name = 'ExchangeArchives_Alias' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeArchives_Alias}' Header='Alias'/> 
									<GridViewColumn x:Name = 'ExchangeArchives_PrimarySMTPAddress' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeArchives_PrimarySMTPAddress}' Header='Primary Mail Address'/>
									<GridViewColumn x:Name = 'ExchangeArchives_ItemCount' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeArchives_ItemCount}' Header='Item Count'/> 
									<GridViewColumn x:Name = 'ExchangeArchives_TotalItemSize' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeArchives_TotalItemSize}' Header='Total Item Size'/> 
									<GridViewColumn x:Name = 'ExchangeArchives_ArchiveStatus' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeArchives_ArchiveStatus}' Header='Archive Status'/>
									<GridViewColumn x:Name = 'ExchangeArchives_UsageLocation' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeArchives_UsageLocation}' Header='Usage Location'/> 
									<GridViewColumn x:Name = 'ExchangeArchives_WhenMailboxCreated' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeArchives_WhenMailboxCreated}' Header='Created'/>
									<GridViewColumn x:Name = 'ExchangeArchives_LastLogonTime' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeArchives_LastLogonTime}' Header='Last logon'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="Exchange Groups">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "Exchange Groups" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'ExchangeGroups_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all Exchange Groups.'>
							<ListView.View>
								<GridView x:Name = 'ExchangeGroups_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'ExchangeGroups_DisplayName' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeGroups_DisplayName}' Header='Display Name'/>
									<GridViewColumn x:Name = 'ExchangeGroups_RecipientTypeDetails' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeGroups_RecipientTypeDetails}' Header='Type'/> 
									<GridViewColumn x:Name = 'ExchangeGroups_Owner' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeGroups_Owner}' Header='Owner'/> 
									<GridViewColumn x:Name = 'ExchangeGroups_WindowsEmailAddress' Width = 'Auto' DisplayMemberBinding = '{Binding Path = ExchangeGroups_WindowsEmailAddress}' Header='Windows EmailAddress'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="SharePoint Sites">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "SharePoint Sites" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'SharePointSites_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all SharePoint Sites.'>
							<ListView.View>
								<GridView x:Name = 'SharePointSites_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'SharePointSites_Url' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointSites_Url}' Header='Url'/>
									<GridViewColumn x:Name = 'SharePointSites_Title' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointSites_Title}' Header='Title'/>
									<GridViewColumn x:Name = 'SharePointSites_WebsCount' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointSites_WebsCount}' Header='Webs'/>
									<GridViewColumn x:Name = 'SharePointSites_StorageUsageCurrent' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointSites_StorageUsageCurrent}' Header='Storage Usage'/>
									<GridViewColumn x:Name = 'SharePointSites_Status' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointSites_Status}' Header='Status'/> 
									<GridViewColumn x:Name = 'SharePointSites_LocaleId' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointSites_LocaleId}' Header='LocaleId'/> 
									<GridViewColumn x:Name = 'SharePointSites_Template' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointSites_Template}' Header='Template'/>
									<GridViewColumn x:Name = 'SharePointSites_Owner' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointSites_Owner}' Header='Owner'/>
									<GridViewColumn x:Name = 'SharePointSites_LastContentModifiedDate' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointSites_LastContentModifiedDate}' Header='Last modified'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>                                    
				</Grid> 
			</TabItem>
			<TabItem Header="SharePoint Webs">
				<Grid Grid.Row = '2' Grid.Column = '0' ShowGridLines = 'false'>  
					<Grid.Resources>
						<Style x:Key="AlternatingRowStyle" TargetType="{x:Type Control}" >
							<Setter Property="Background" Value="LightGray"/>
							<Setter Property="Foreground" Value="Black"/>
							<Style.Triggers>
								<Trigger Property="ItemsControl.AlternationIndex" Value="1">                            
									<Setter Property="Background" Value="White"/>
									<Setter Property="Foreground" Value="Black"/>                                
								</Trigger>                            
							</Style.Triggers>
						</Style>                    
					</Grid.Resources>                  
					<Grid.ColumnDefinitions>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
						<ColumnDefinition Width="*"/>
					</Grid.ColumnDefinitions>
					<Grid.RowDefinitions>
						<RowDefinition Height = '*'/>
					</Grid.RowDefinitions> 
					<GroupBox Header = "SharePoint Webs" Grid.Column = '0' Grid.Row = '2' Grid.ColumnSpan = '11' Grid.RowSpan = '3'>
						<Grid Width = 'Auto' Height = 'Auto' ShowGridLines = 'false'>
						<ListView x:Name = 'SharePointWebs_Listview' AllowDrop = 'True' AlternationCount="2" ItemContainerStyle="{StaticResource AlternatingRowStyle}"
						ToolTip = 'Lists all SharePoint Webs.'>
							<ListView.View>
								<GridView x:Name = 'SharePointWebs_GridView' AllowsColumnReorder = 'True' ColumnHeaderTemplate="{StaticResource HeaderTemplate}">
									<GridViewColumn x:Name = 'SharePointWebs_ServerRelativeUrl' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointWebs_ServerRelativeUrl}' Header='Site Collection'/>
									<GridViewColumn x:Name = 'SharePointWebs_Title' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointWebs_Title}' Header='Title'/>
									<GridViewColumn x:Name = 'SharePointWebs_Created' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointWebs_Created}' Header='Created'/>
									<GridViewColumn x:Name = 'SharePointWebs_LastItemModifiedDate' Width = 'Auto' DisplayMemberBinding = '{Binding Path = SharePointWebs_LastItemModifiedDate}' Header='Last item modified'/>
								</GridView>
							</ListView.View>        
						</ListView>                
						</Grid>
					</GroupBox>
				</Grid> 
			</TabItem>
		</TabControl>
        <TextBox x:Name = 'StatusTextBox' Grid.Row = '4' ToolTip = 'Displays current status of operation'> Waiting for Action... </TextBox>                           
    </Grid>   
</Window>
"@ 
#endregion

#region Load XAML into PowerShell
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
$uiHash.Window=[Windows.Markup.XamlReader]::Load( $reader )
#endregion

#region create runspace
$newRunspace =[runspacefactory]::CreateRunspace()
$newRunspace.ApartmentState = "STA"
$newRunspace.ThreadOptions = "ReuseThread"
$newRunspace.Open()
$newRunspace.SessionStateProxy.SetVariable("uiHash",$uiHash)
$sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
#endregion
 
#region Connect to all controls
[xml]$XAML = $xaml
        $xaml.SelectNodes("//*[@*[contains(translate(name(.),'n','N'),'Name')]]") | %{
        #Find all of the form types and add them as members to the synchash
        $uiHash.Add($_.Name,$uiHash.Window.FindName($_.Name) )
    }
#endregion

#region Event Handlers

#region Window Load Events
$uiHash.Window.Add_SourceInitialized({  
    #Define hashtable of settings
    $Script:SortHash = @{}
    
	#region AAD Users Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:AADUsers_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.AADUsers_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.AADUsers_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.AADUsers_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.AADUsers_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $AADUsers_ColumnSortHandler)
	#endregion
	
	#region AAD Deleted Users Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:AADDeletedUsers_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.AADDeletedUsers_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.AADDeletedUsers_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.AADDeletedUsers_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.AADDeletedUsers_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $AADDeletedUsers_ColumnSortHandler)
	#endregion
	
	#region AAD External Users Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:AADExternalUsers_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.AADExternalUsers_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.AADExternalUsers_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.AADExternalUsers_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.AADExternalUsers_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $AADExternalUsers_ColumnSortHandler)
	#endregion
	
	#region AAD Contacts Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:AADContacts_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.AADContacts_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.AADContacts_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.AADContacts_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.AADContacts_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $AADContacts_ColumnSortHandler)
	#endregion
	
	#region AADGroup Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:AADGroups_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.AADGroups_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.AADGroups_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.AADGroups_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.AADGroups_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $AADGroups_ColumnSortHandler)
	#endregion
	
	#region AADLicenses Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:AADLicenses_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.AADLicenses_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.AADLicenses_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.AADLicenses_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.AADLicenses_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $AADLicenses_ColumnSortHandler)
	#endregion
	
	#region AADDomains Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:AADDomains_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.AADDomains_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.AADDomains_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.AADDomains_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.AADDomains_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $AADDomains_ColumnSortHandler)
	#endregion

	#region ExchangeMailboxes Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:ExchangeMailboxes_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.ExchangeMailboxes_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.ExchangeMailboxes_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.ExchangeMailboxes_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.ExchangeMailboxes_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ExchangeMailboxes_ColumnSortHandler)
	#endregion

	#region ExchangeArchives Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:ExchangeArchives_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.ExchangeMailboxes_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.ExchangeMailboxes_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.ExchangeMailboxes_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.ExchangeArchives_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ExchangeArchives_ColumnSortHandler)
	#endregion
	
	#region ExchangeGroups Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:ExchangeGroups_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.ExchangeGroups_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.ExchangeGroups_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.ExchangeGroups_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.ExchangeGroups_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $ExchangeGroups_ColumnSortHandler)
	#endregion
	
	#region SharePointSites Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:SharePointSites_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.SharePointSites_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.SharePointSites_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.SharePointSites_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.SharePointSites_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $SharePointSites_ColumnSortHandler)
	#endregion
	
	#region SharePointWebs Sort
    #Sort event handler
    [System.Windows.RoutedEventHandler]$Global:SharePointWebs_ColumnSortHandler = {
        If ($_.OriginalSource -is [System.Windows.Controls.GridViewColumnHeader]) {
            Write-Verbose ("{0}" -f $_.Originalsource.getType().FullName)
            If ($_.OriginalSource -AND $_.OriginalSource.Role -ne 'Padding') {
                $Column = $_.Originalsource.Column.DisplayMemberBinding.Path.Path
                Write-Debug ("Sort: {0}" -f $Column)
                If ($SortHash[$Column] -eq 'Ascending') {
                    Write-Debug "Descending"
                    $SortHash[$Column]  = 'Descending'
                } Else {
                    Write-Debug "Ascending"
                    $SortHash[$Column]  = 'Ascending'
                }
                Write-Verbose ("Direction: {0}" -f $SortHash[$Column])
                $lastColumnsort = $Column
                Write-Verbose "Clearing sort descriptions"
                $uiHash.SharePointWebs_Listview.Items.SortDescriptions.clear()
                Write-Verbose ("Sorting {0} by {1}" -f $Column, $SortHash[$Column])
                $uiHash.SharePointWebs_Listview.Items.SortDescriptions.Add((New-Object System.ComponentModel.SortDescription $Column, $SortHash[$Column]))
                Write-Verbose "Refreshing View"
                $uiHash.SharePointWebs_Listview.Items.Refresh()   
            }             
        }
    }
    $uiHash.SharePointWebs_Listview.AddHandler([System.Windows.Controls.GridViewColumnHeader]::ClickEvent, $SharePointWebs_ColumnSortHandler)
	#endregion
}) 
#endregion   

#region Window Close Events
#Window Close Events
$uiHash.Window.Add_Closed({
	$newRunspace.close()
	$newRunspace.dispose()
	
    [gc]::Collect()
    [gc]::WaitForPendingFinalizers()    
}) 

#Exit Menu
$uiHash.ExitMenu.Add_Click({
	$newRunspace.close()
	$newRunspace.dispose()
    $uiHash.Window.Close()
})
#endregion

#region Error events
#View Error Event
$uiHash.ViewErrorMenu.Add_Click({
    Get-Error | Out-GridView
})

#Clear Error log
$uiHash.ClearErrorMenu.Add_Click({
    Write-Verbose "Clearing error log"
    $Error.Clear()
})
#endregion

#region Run events
#Run Menu
$uiHash.RunMenu.Add_Click({
	$PowerShell = [PowerShell]::Create().AddScript({
		param(
			$path,
			$header,
			$uiHash
		)
		Set-Location $path
		#DotSource shared functions
		. "$($path)\Scripts\SharedFunctions.ps1"

		#DotSource shared functions
		. "$($path)\Scripts\Start-RunJob.ps1"
		
		Start-RunJob -header $header 
	}).AddArgument($path).AddArgument($uiHash.Tabs.SelectedItem.header).AddArgument($uiHash)
	
	$PowerShell.Runspace = $newRunspace
	$data = $PowerShell.BeginInvoke()
})

#Run All Menu
$uiHash.RunAllMenu.Add_Click({
	$PowerShell = [PowerShell]::Create().AddScript({
		param(
			$path,
			$uiHash
		)
		
		Set-Location $path
		
		#DotSource shared functions
		. "$($path)\Scripts\SharedFunctions.ps1"

		#DotSource shared functions
		. "$($path)\Scripts\Start-RunJob.ps1"
		
		Start-RunJob -header "AAD Users"
		Start-RunJob -header "AAD Deleted Users"
		Start-RunJob -header "AAD External Users"
		Start-RunJob -header "AAD Contacts"
		Start-RunJob -header "AAD Groups"
		Start-RunJob -header "AAD Licenses"
		Start-RunJob -header "AAD Domains"
		Start-RunJob -header "Exchange Mailboxes"
		Start-RunJob -header "Exchange Archives"
		Start-RunJob -header "Exchange Groups"
		Start-RunJob -header "SharePoint Sites"
		Start-RunJob -header "SharePoint Webs"
	}).AddArgument($path).AddArgument($uiHash)
	
	$PowerShell.Runspace = $newRunspace
	$data = $PowerShell.BeginInvoke()	
})

#RunButton Event    
$uiHash.RunButton.add_Click({
	$PowerShell = [PowerShell]::Create().AddScript({
		param(
			$path,
			$header,
			$uiHash
		)
		Set-Location $path
		#DotSource shared functions
		. "$($path)\Scripts\SharedFunctions.ps1"

		#DotSource shared functions
		. "$($path)\Scripts\Start-RunJob.ps1"
		
		Start-RunJob -header $header 
	}).AddArgument($path).AddArgument($uiHash.Tabs.SelectedItem.header).AddArgument($uiHash)
	
	$PowerShell.Runspace = $newRunspace
	$data = $PowerShell.BeginInvoke()
})

#RunAllButton Event   
$uiHash.RunAllButton.add_Click({
	$PowerShell = [PowerShell]::Create().AddScript({
		param(
			$path,
			$uiHash
		)
		
		Set-Location $path
		
		#DotSource shared functions
		. "$($path)\Scripts\SharedFunctions.ps1"

		#DotSource shared functions
		. "$($path)\Scripts\Start-RunJob.ps1"
		
		Start-RunJob -header "AAD Users"
		Start-RunJob -header "AAD Deleted Users"
		Start-RunJob -header "AAD External Users"
		Start-RunJob -header "AAD Contacts"
		Start-RunJob -header "AAD Groups"
		Start-RunJob -header "AAD Licenses"
		Start-RunJob -header "AAD Domains"
		Start-RunJob -header "Exchange Mailboxes"
		Start-RunJob -header "Exchange Archives"
		Start-RunJob -header "Exchange Groups"
		Start-RunJob -header "SharePoint Sites"
		Start-RunJob -header "SharePoint Webs"
	}).AddArgument($path).AddArgument($uiHash)
	
	$PowerShell.Runspace = $newRunspace
	$data = $PowerShell.BeginInvoke()	
})


#endregion

#region Report events
#Report Generation
$uiHash.GenerateReportButton.Add_Click({
    Start-Report
})
#endregion

#region connect events
#ConnectButton Events    
$uiHash.ConnectButton.add_Click({
	$credential = get-credential

	$PowerShell = [PowerShell]::Create().AddScript({
		param(
			$path,
			$uiHash,
			$credential
		)
		
		Set-Location $path
		
		#DotSource shared functions
		. "$($path)\Scripts\SharedFunctions.ps1"
		
		#DotSource shared functions
		. "$($path)\Scripts\Start-ConnectJob.ps1"
		
		Start-ConnectJob -credential $credential
	}).AddArgument($path).AddArgument($uiHash).AddArgument($credential)
	
	$PowerShell.Runspace = $newRunspace
	$data = $PowerShell.BeginInvoke()	
})

#ConnectMenu Events   
$uiHash.ConnectMenu.add_Click({
	$credential = get-credential

	$PowerShell = [PowerShell]::Create().AddScript({
		param(
			$path,
			$uiHash,
			$credential
		)
		
		Set-Location $path
		
		#DotSource shared functions
		. "$($path)\Scripts\SharedFunctions.ps1"
		
		#DotSource shared functions
		. "$($path)\Scripts\Start-ConnectJob.ps1"
		
		Start-ConnectJob -credential $credential
	}).AddArgument($path).AddArgument($uiHash).AddArgument($credential)
	
	$PowerShell.Runspace = $newRunspace
	$data = $PowerShell.BeginInvoke()	
})
#endregion

#region help menu items
#AboutMenu Event
$uiHash.AboutMenu.Add_Click({
    Open-About
})

#HelpFileMenu Event
$uiHash.HelpFileMenu.Add_Click({
    Open-Help
})
#endregion

#region Key Up Events
#Key Up Event
$uiHash.Window.Add_KeyUp({
    $Global:Test = $_
    Write-Debug ("Key Pressed: {0}" -f $_.Key)
    Switch ($_.Key) {
        "F1" {Open-Help}
		"F4" {
			$PowerShell = [PowerShell]::Create().AddScript({
				param(
					$path,
					$uiHash
				)
				
				Set-Location $path
				
				#DotSource shared functions
				. "$($path)\Scripts\SharedFunctions.ps1"
				
				#DotSource shared functions
				. "$($path)\Scripts\Start-ConnectJob.ps1"
				
				Start-ConnectJob
			}).AddArgument($path).AddArgument($uiHash)
			
			$PowerShell.Runspace = $newRunspace
			$data = $PowerShell.BeginInvoke()			
		}
        "F5" {
			$PowerShell = [PowerShell]::Create().AddScript({
				param(
					$path,
					$header,
					$uiHash
				)
				Set-Location $path
				#DotSource shared functions
				. "$($path)\Scripts\SharedFunctions.ps1"

				#DotSource shared functions
				. "$($path)\Scripts\Start-RunJob.ps1"
				
				Start-RunJob -header $header 
			}).AddArgument($path).AddArgument($uiHash.Tabs.SelectedItem.header).AddArgument($uiHash)
			
			$PowerShell.Runspace = $newRunspace
			$data = $PowerShell.BeginInvoke()
		}
		"F6" {
			$PowerShell = [PowerShell]::Create().AddScript({
				param(
					$path,
					$uiHash
				)
		
				Set-Location $path
				
				#DotSource shared functions
				. "$($path)\Scripts\SharedFunctions.ps1"

				#DotSource shared functions
				. "$($path)\Scripts\Start-RunJob.ps1"
				
				Start-RunJob -header "AAD Users"
				Start-RunJob -header "AAD Deleted Users"
				Start-RunJob -header "AAD External Users"
				Start-RunJob -header "AAD Contacts"
				Start-RunJob -header "AAD Groups"
				Start-RunJob -header "AAD Licenses"
				Start-RunJob -header "AAD Domains"
				Start-RunJob -header "Exchange Mailboxes"
				Start-RunJob -header "Exchange Archives"
				Start-RunJob -header "Exchange Groups"
				Start-RunJob -header "SharePoint Sites"
				Start-RunJob -header "SharePoint Webs"
			}).AddArgument($path).AddArgument($uiHash)
			$PowerShell.Runspace = $newRunspace
			$data = $PowerShell.BeginInvoke()	
		}
        "F8" {Start-Report}
        Default {$Null}
    }

})
#endregion
#endregion        

#Start the GUI
$uiHash.Window.ShowDialog() | Out-Null
