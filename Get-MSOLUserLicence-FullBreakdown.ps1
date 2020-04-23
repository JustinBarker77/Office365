#Requires -Modules MSOnline
#Requires -Version 5

<#
	.SYNOPSIS
		Name: Get-MSOLUserLicence-FullBreakdown.ps1
		The purpose of this script is is to export licensing details to excel

	.DESCRIPTION
		This script will log in to Office 365 and then create a license report by SKU, with each component level status for each user, where 1 or more is assigned. This then conditionally formats the output to colours and autofilter.

	.NOTES
		Version 1.15
		Updated: 20190602	V1.7	Parameters, Comment based help, creates folder and deletes folder for csv's, require statements
		Updated: 20190614	V1.8	Added more SKU's and Components
       	Updated: 20190627	V1.9	Added more Components
		Updated: 20190830   V1.10   Added more components. Updated / renamed refreshed licences
		Updated: 20190916	V1.11	Added more components and SKU's
		Updated: 20191015	V1.12	Tidied up old comments
		Updated: 20200204   V1.13   Added more SKU's and Components
		Updated: 20200408   V1.14   Added Teams Exploratory SKU
		Updated: 20200422	V1.15   Formats to Segoe UI 9pt. Removed unnecessary True output. 
		Release Date: 20190530
		Release notes from original:
			1.0 - Initital Release                                                                                                            	
			1.1 - Added Switch for additional licence components                                                                              	
			1.2 - Added PowerApps Plan 2 Trial  for additional licence components                                                             	
			1.3 - Added Freeze Panes to Excel output                                                                      	
			1.4 - Added AX7 User Trial, Project Online Professional, Visio Online Plan 2, Office 365 E1, Whiteboard SKUs                      	
			1.5 - Added Microsoft Search, Premium Encryption and Teams Commercial Trial, RMS Ad Hoc SKUs                                      	
			1.6 - Added Microsoft 365 E3 and F1 SKU & performs actions on cell by cell basis for colouring					                    
		Authors: Mark Lofthouse, Justin Barker & Robin Dadswell
		References: 
			https://gallery.technet.microsoft.com/scriptcenter/Export-a-Licence-b200ca2a
			https://stackoverflow.com/questions/31183106/can-powershell-generate-a-plain-excel-file-with-multiple-sheets
#>
[CmdletBinding()]
param (
	[Parameter(
		Mandatory,
		HelpMessage="Name of the Company you are running this against. This will form part of the output file name")
		]
		[string]$CompanyName,
	[Parameter(
		Mandatory,
		HelpMessage="The location you would like the final excel file to reside"
		)][ValidateScript({
			if (!(Test-Path -Path $_)) {
				throw "The folder $_ does not exist"
			}
			else {
				return $true
			}
		})]
	[System.IO.DirectoryInfo]$OutputPath,
	[Parameter(
		HelpMessage = "Credentials to connect to Office 365 if not already connected"
	)]
	[PSCredential]$Office365Credentials
)
#Following Function Switches Complicated Component Names with Friendly Names
function componentlicenseswitch {
	param
	(
		[parameter (Mandatory=$true, Position = 1)][string]$component
	)
	switch -wildcard ($($component)) {
		#Office 365 E5
		"RECORDS_MANAGEMENT"		{$thisLicence = "Microsoft Records Management"}
		"INFO_GOVERNANCE"			{$thisLicence = "Microsoft Information Governance"}
		"DATA_INVESTIGATIONS"		{$thisLicence = "Microsoft Data Investigations"}
		"CUSTOMER_KEY"			    {$thisLicence = "Microsoft Customer Key"}
		"COMMUNICATIONS_DLP"		{$thisLicence = "Microsoft Communications DLP"}
		"COMMUNICATIONS_COMPLIANCE"	{$thisLicence = "Microsoft Communications Compliance"}
		"M365_ADVANCED_AUDITING"	{$thisLicence = "Microsoft 365 Advanced Auditing"}
		"MYANALYTICS_P2"       {$thisLicence = "Insights by MyAnalytics" }
		"PAM_ENTERPRISE"       {$thisLicence = "O365 Priviledged Access Management" }
		"BPOS_S_TODO_3"        {$thisLicence = "To-Do (Plan 3)" }
		"FORMS_PLAN_E5"        {$thisLicence = "Microsoft Forms E5" }
		"STREAM_O365_E5"       {$thisLicence = "Stream E5" }
		"THREAT_INTELLIGENCE"  {$thisLicence = "O365 ATP Plan 2" }
		"Deskless"             {$thisLicence = "Microsoft StaffHub" }
		"FLOW_O365_P3"         {$thisLicence = "Flow for Office 365" }
		"POWERAPPS_O365_P3"    {$thisLicence = "PowerApps for Office 365" }
		"TEAMS1"               {$thisLicence = "Microsoft Teams" }
		"MCO_TEAMS_IW"         {$thisLicence = "Microsoft Teams Trial"}
		"ADALLOM_S_O365"       {$thisLicence = "Office 365 Cloud App Security" }
		"EQUIVIO_ANALYTICS"    {$thisLicence = "Office 365 Advanced eDiscovery" }
		"LOCKBOX_ENTERPRISE"   {$thisLicence = "Customer Lockbox" }
		"EXCHANGE_ANALYTICS"   {$thisLicence = "Microsoft MyAnalytics (Full)" }
		"SWAY"                 {$thisLicence = "Sway" }
		"ATP_ENTERPRISE"       {$thisLicence = "O365 ATP Plan 1 (not licenced individually)" }
		"MCOEV"                {$thisLicence = "M365 Phone System" }
		"MCOMEETADV"           {$thisLicence = "M365 Audio Conferencing" }
		"BI_AZURE_P2"          {$thisLicence = "Power BI Pro" }
		"INTUNE_O365"          {$thisLicence = "MDM for Office 365 (not licenced individually)" }
		"PROJECTWORKMANAGEMENT"{$thisLicence = "Microsoft Planner" }
		"RMS_S_ENTERPRISE"     {$thisLicence = "Azure Rights Management" }
		"YAMMER_ENTERPRISE"    {$thisLicence = "Yammer Enterprise" }
		"OFFICESUBSCRIPTION"   {$thisLicence = "Office 365 ProPlus" }
		"MCOSTANDARD"          {$thisLicence = "Skype for Business Online (Plan 2)" }
		"EXCHANGE_S_ENTERPRISE"{$thisLicence = "Exchange Online (Plan 2)" }
		"SHAREPOINTENTERPRISE" {$thisLicence = "SharePoint Online (Plan 2)" }

		"SHAREPOINTWAC"        {$thisLicence = "Office Online" }			
		"EXCHANGE_S_FOUNDATION"{$thisLicence = "Core Exchange for non-Exch SKUs (e.g. setting profile pic)" }
		"ATA"                  {$thisLicence = "Azure Advanced Threat Protection" }
		"ADALLOM_S_STANDALONE" {$thisLicence = "Microsoft Cloud App Security" }
		"RMS_S_PREMIUM2"       {$thisLicence = "Azure Information Protection Premium P2" }
		"RMS_S_PREMIUM"        {$thisLicence = "Azure Information Protection Premium P1" } 
		"INTUNE_A"             {$thisLicence = "Microsoft Intune" }
		"AAD_PREMIUM_P2"       {$thisLicence = "Azure Active Directory Premium P2" }
		"MFA_PREMIUM"          {$thisLicence = "Microsoft Azure Multi-Factor Authentication" }
		"AAD_PREMIUM"          {$thisLicence = "Azure Active Directory Premium P1" }
		"BPOS_S_TODO_2"        {$thisLicence = "To-Do (Plan 2)" }
		"FORMS_PLAN_E3"        {$thisLicence = "Microsoft Forms (Plan E3)" }
		"STREAM_O365_E3"       {$thisLicence = "Microsoft Stream for O365 E3 SKU" }
		"FLOW_O365_P2"         {$thisLicence = "Flow for Office 365" }
		"POWERAPPS_O365_P2"    {$thisLicence = "PowerApps for Office 365" }
		"SPZA"                 {$thisLicence = "App Connect" }
		"DYN365_CDS_VIRAL"     {$thisLicence = "Common Data Service" }
		"FLOW_P2_VIRAL"        {$thisLicence = "Flow Free" }
		"MICROSOFT_BUSINESS_CENTER" {$thisLicence = "Microsoft Business Center" }
		"BI_AZURE_P0"          {$thisLicence = "Power BI (Free)" }
		"ADALLOM_S_DISCOVERY"  {$thisLicence = "Cloud App Security Discovery" }
		"POWERAPPS_O365_P1"    {$thisLicence = "PowerApps for Office 365"}
		"FLOW_O365_P1"         {$thisLicence = "Flow for Office 365"}
		"SHAREPOINTDESKLESS"   {$thisLicence = "SharePoint Online Kiosk"}
		"FLOW_DYN_APPS"        {$thisLicence = "Flow for Dynamics 365"}
		"POWERAPPS_DYN_APPS"   {$thisLicence = "PowerApps for Dynamics 365"}
		"PROJECT_ESSENTIALS"   {$thisLicence = "Project Online Essentials"}
		"NBENTERPRISE"         {$thisLicence = "Microsoft Social Engagement - Service Discontinuation"}
		"DYN365_ENTERPRISE_SALES" {$thisLicence = "Dynamics 365 for Sales"}
		"Dynamics_365_for_Talent_Team_members" {$thisLicence = "Dynamics 365 for Talent Team members"}
		"Dynamics_365_for_Retail_Team_members" {$thisLicence = "Dynamics 365 for Retail Team members"}
		"DYN365_Enterprise_Talent_Onboard_TeamMember" {$thisLicence = "Dynamics 365 for Talent - Onboard Experience"}
		"DYN365_Enterprise_Talent_Attract_TeamMember" {$thisLicence = "Dynamics 365 for Talent - Attract Experience Team Member"}
		"Dynamics_365_for_Operations_Team_members" {$thisLicence = "Dynamics_365_for_Operations_Team_members"}
		"DYN365_ENTERPRISE_TEAM_MEMBERS"{$thisLicence = "Dynamics 365 for Team Members"}
		"POWERAPPS_P2_VIRAL"      		{$thisLicence = "PowerApps Plan 2 Trial"}
		"FLOW_P2_VIRAL_REAL"      		{$thisLicence = "Flow P2 Viral"}
		"MIP_S_CLP1"      				{$thisLicence = "Information Protection for Office 365 - Standard"}
		"MIP_S_CLP2"      				{$thisLicence = "Information Protection for Office 365 - Premium"}
		"ERP_TRIAL_INSTANCE"      		{$thisLicence = "AX7 Instance"}
		"PROJECT_PROFESSIONAL"      	{$thisLicence = "Project P3"}
		"FLOW_FOR_PROJECT"      		{$thisLicence = "Data Integration for Project with Flow"}
		"DYN365_CDS_PROJECT"      		{$thisLicence = "Common Data Service for Project"}
		"SHAREPOINT_PROJECT"      		{$thisLicence = "Project Online Service"}
		"PROJECT_CLIENT_SUBSCRIPTION"	{$thisLicence = "Project Online Desktop Client"}
		"ONEDRIVE_BASIC"      			{$thisLicence = "OneDrive Basic"}
		"VISIOONLINE"      				{$thisLicence = "Visio Online"}
		"VISIO_CLIENT_SUBSCRIPTION" 	{$thisLicence = "Visio Pro for Office 365"}
		"WHITEBOARD_PLAN1"      	{$thisLicence = "Whiteboard Plan 1"}
		"OFFICEMOBILE_SUBSCRIPTION" {$thisLicence = "Office Mobile Apps for Office 365"}
		"BPOS_S_TODO_1"      		{$thisLicence = "To-Do Plan 1"}
		"FORMS_PLAN_E1"      		{$thisLicence = "Microsoft Forms (Plan E1)"}
		"STREAM_O365_E1"      		{$thisLicence = "Microsoft Stream for O365 E1"}
		"SHAREPOINTSTANDARD"      	{$thisLicence = "SharePoint Online (Plan 1)"}
		"EXCHANGE_S_STANDARD"      	{$thisLicence = "Exchange Online (Plan 1)"}
		"WHITEBOARD_PLAN2"      	{$thisLicence = "Whiteboard Plan 2"}
		"WHITEBOARD_PLAN3"      	{$thisLicence = "Whiteboard Plan 3"}
		"MICROSOFT_SEARCH"      	{$thisLicence = "Microsoft Search"}
		"PREMIUM_ENCRYPTION"      	{$thisLicence = "Premium Encryption"}
		"RMS_S_ADHOC"      			{$thisLicence = "Rights Management Adhoc"}
        "WIN10_PRO_ENT_SUB"      	{$thisLicence = "Win 10 Enterprise E3"}
        "WHITEBOARD_FIRSTLINE1"     {$thisLicence = "Whiteboard for Firstline"}
        "BPOS_S_TODO_FIRSTLINE"     {$thisLicence = "To-Do Firstline"}
        "WIN10_ENT_LOC_F1"          {$thisLicence = "Win 10 Enterprise E3 (Local Only)"}
        "MCOIMP"                    {$thisLicence = "Skype for Business (Plan 1)"}
        "POWERAPPS_O365_S1"         {$thisLicence = "PowerApps for Office 365 Firstline"}
        "STREAM_O365_K"             {$thisLicence = "Stream for Office 365 Firstline"}
        "POWERAPPS_O365_S1"      	{$thisLicence = "PowerApps for Office 365 Firstline"}
        "FORMS_PLAN_K"      		{$thisLicence = "Microsoft Forms (Plan F1)"}
        "FLOW_O365_S1"      		{$thisLicence = "Flow for Office 365 (F1)"}
        "EXCHANGE_S_DESKLESS"      	{$thisLicence = "Exchange Online Firstline"}
        "WINDEFATP"                	{$thisLicence = "Windows Defender ATP"}
		"DYN365_ENTERPRISE_P1_IW"	{$thisLicence = "Dyn 365 P1 Trial Info Workers"}
		"FLOW_DYN_P2"				{$thisLicence = "Flow for Dynamics 365"}
		"POWERAPPS_DYN_P2"			{$thisLicence = "PowerApps for Dynamics 365"}
		"DYN365_ENTERPRISE_P1"		{$thisLicence = "Dynamics Enterprise P1"}
		"D365_CSI_EMBED_CE" 		{$thisLicence = "Dynamics 365 Customer Service Insights for CE Plan"}
		"EXCHANGE_S_ARCHIVE_ADDON"	{$thisLicence = "Exchange Online Archiving Add-on"}
		"DYN365_CDS_P1"				{$thisLicence = "Common Data Service"}
		"FLOW_P1"					{$thisLicence = "Microsoft Flow Plan 1"}
		"FLOW_P2"					{$thisLicence = "Microsoft Flow Plan 2"}
		"POWERAPPS_P2"				{$thisLicence = "PowerApps Plan 2"}
		"DYN365_CDS_P2"				{$thisLicence = "Common Data Service"}
		"INFORMATION_BARRIERS"		{$thisLicence = "Information Barriers"}
		"KAIZALA_STANDALONE"		{$thisLicence = "Microsoft Kaizala Pro"}
		"KAIZALA_O365_P3"			{$thisLicence = "Kaizala for Office 365"}
		"FLOW_CCI_BOTS"				{$thisLicence = "Flow for CCI Bots"}
		"CCIBOTS_PRIVPREV_VIRAL"	{$thisLicence = "Dynamics 365 AI for Customer Service Virtual Agents Viral"}
		"DYN365_CDS_CCI_BOTS"		{$thisLicence = "Common Data Service for CCI Bots"}
		"DYN365_AI_SERVICE_INSIGHS"	{$thisLicence = "Dynamics 365 Customer Service Insights"}
		"POWERAPPS_DYN_TEAM"		{$thisLicence = "PowerApps for Dynamics 365"}
		"FLOW_DYN_TEAM"				{$thisLicence = "Flow for Dynamics 365"}
		"DYN365_TEAM_MEMBERS"		{$thisLicence = "Dynamics 365 Team Members"}
		"Forms_Pro_CE"				{$thisLicence = "Forms Pro for Customer Engagement Plan"}
        "DYN365_BUSINESS_Marketing"	{$thisLicence = "Dynamics 365 Marketing"}
        "DYN365_RETAIL_TRIAL"		{$thisLicence = "Dynamics 365 Retail Trial"}
        "FORMS_PRO"			        {$thisLicence = "Forms Pro"}
        "FLOW_FORMS_PRO"			{$thisLicence = "Flow for Forms Pro"}
        "DYN365_CDS_FORMS_PRO"		{$thisLicence = "Common Data Service"}
        "KAIZALA_O365_P1"		    {$thisLicence = "Microsoft Kaizala Pro (P1)"}
        "KAIZALA_O365_P2"		    {$thisLicence = "Microsoft Kaizala Pro"}
		"DYN365_CDS_DYN_APPS"		{$thisLicence = "Common Data Service"}
		"Forms_Pro_Operations" 		{$thisLicence = "Microsoft Forms Pro for Operations"}
		"Dynamics_365_for_Retail"	{$thisLicence = "Dynamics 365 for Retail"}
		"DYN365_TALENT_ENTERPRISE"	{$thisLicence = "Dynamics 365 for Talent"}
		"DYN365_CDS_DYN_P2"			{$thisLicence = "Common Data Service"}
		"ONEDRIVESTANDARD"			{$thisLicence = "OneDrive for Business (Plan 1)"}
		"Dynamics_365_Talent_Onboard"			{$thisLicence = "Dynamics 365 for Talent: Onboard"}
        "Dynamics_365_Onboarding_Free_PLAN"	    {$thisLicence = "Dynamics 365 for Talent: Onboard"}
        "Dynamics_365_Hiring_Free_PLAN"		    {$thisLicence = "Dynamics 365 for Talent: Attract"}
		"Dynamics_365_for_HCM_Trial"		    {$thisLicence = "Dynamics_365_for_HCM_Trial"}
        "TEAMS_FREE_SERVICE"			{$thisLicence = "Teams Free Service (Not assigned per user)"}
        "MCOFREE"			        {$thisLicence = "MCO Free for Microsoft Teams (free)"}
        "TEAMS_FREE"			    {$thisLicence = "Microsoft Teams (free)"}
        "ML_CLASSIFICATION"		    {$thisLicence = "Microsoft ML-Based Classification"}
        "INSIDER_RISK_MANAGEMENT"   {$thisLicence = "Microsoft Insider Risk Management"}
        "SAFEDOCS"			        {$thisLicence = "Office 365 SafeDocs"}
		default {$thisLicence = $component }
	}
	Write-Output $thisLicence
}
#Following Function Switches Complicated Top Level SKU Names with Friendly Names
function RootLicenceswitch {
	param (
		[parameter (Mandatory=$true, Position = 1)][string]$licensesku
	)
	switch -wildcard ($($licensesku)) {
		"O365_BUSINESS_ESSENTIALS"		    {$RootLicence = "Office 365 Business Essentials"}
		"O365_BUSINESS_PREMIUM"			    {$RootLicence = "Office 365 Business Premium"}
		"DESKLESSPACK"					    {$RootLicence = "Office 365 (Plan F1)"}
		"DESKLESSWOFFPACK"				    {$RootLicence = "Office 365 (Plan K2)"}
		"LITEPACK"						    {$RootLicence = "Office 365 (Plan P1)"}
		"EXCHANGESTANDARD"				    {$RootLicence = "Office 365 Exchange Online Only"}
		"STANDARDPACK"					    {$RootLicence = "Office 365 (Plan E1)"}
		"STANDARDWOFFPACK"				    {$RootLicence = "Office 365 (Plan E2)"}
		"ENTERPRISEPACK"					{$RootLicence = "Office 365 (Plan E3)"}
		"ENTERPRISEPACKLRG"				    {$RootLicence = "Office 365 (Plan E3)"}
		"ENTERPRISEWITHSCAL"				{$RootLicence = "Office 365 (Plan E4)"}
		"ENTERPRISEPREMIUM_NOPSTNCONF"	    {$RootLicence = "Office 365 (Plan E5) (without Audio Conferencing)"}
		"ENTERPRISEPREMIUM"				    {$RootLicence = "Office 365 (Plan E5)"}
		"STANDARDPACK_STUDENT"			    {$RootLicence = "Office 365 (Plan A1) for Students"}
		"STANDARDWOFFPACKPACK_STUDENT"	    {$RootLicence = "Office 365 (Plan A2) for Students"}
		"ENTERPRISEPACK_STUDENT"			{$RootLicence = "Office 365 (Plan A3) for Students"}
		"ENTERPRISEWITHSCAL_STUDENT"		{$RootLicence = "Office 365 (Plan A4) for Students"}
		"STANDARDPACK_FACULTY"			    {$RootLicence = "Office 365 (Plan A1) for Faculty"}
		"STANDARDWOFFPACKPACK_FACULTY"	    {$RootLicence = "Office 365 (Plan A2) for Faculty"}
		"ENTERPRISEPACK_FACULTY"			{$RootLicence = "Office 365 (Plan A3) for Faculty"}
		"ENTERPRISEWITHSCAL_FACULTY"		{$RootLicence = "Office 365 (Plan A4) for Faculty"}
		"ENTERPRISEPACK_B_PILOT"			{$RootLicence = "Office 365 (Enterprise Preview)"}
		"STANDARD_B_PILOT"				    {$RootLicence = "Office 365 (Small Business Preview)"}
		"VISIOCLIENT"					    {$RootLicence = "Visio Pro Online"}
		"POWER_BI_ADDON"					{$RootLicence = "Office 365 Power BI Addon"}
		"POWER_BI_INDIVIDUAL_USE"		    {$RootLicence = "Power BI Individual User"}
		"POWER_BI_STANDALONE"			    {$RootLicence = "Power BI Stand Alone"}
		"POWER_BI_STANDARD"				    {$RootLicence = "Power BI Free"}
		"PROJECTESSENTIALS"				    {$RootLicence = "Project Lite"}
		"PROJECTCLIENT"					    {$RootLicence = "Project Professional"}
		"PROJECTONLINE_PLAN_1"			    {$RootLicence = "Project Online"}
		"PROJECTONLINE_PLAN_2"			    {$RootLicence = "Project Online and PRO"}
		"ProjectPremium"					{$RootLicence = "Project Online Premium"}
		"ECAL_SERVICES"					    {$RootLicence = "ECAL"}
		"EMS"							    {$RootLicence = "EMS (Plan E3)"}
		"EMSPREMIUM"                        {$RootLicence = "EMS (Plan E5)"}
		"RIGHTSMANAGEMENT_ADHOC"			{$RootLicence = "Windows Azure RMS"}
		"MCOMEETADV"						{$RootLicence = "PSTN conferencing"}
		"SHAREPOINTSTORAGE"				    {$RootLicence = "SharePoint storage"}
		"PLANNERSTANDALONE"				    {$RootLicence = "Planner Standalone"}
		"CRMIUR"							{$RootLicence = "CMRIUR"}
		"BI_AZURE_P1"					    {$RootLicence = "Power BI Reporting and Analytics"}
		"INTUNE_A"						    {$RootLicence = "Windows Intune Plan A"}
		"PROJECTWORKMANAGEMENT"			    {$RootLicence = "Office 365 Planner Preview"}
		"ATP_ENTERPRISE"					{$RootLicence = "Ex Online ATP Plan 1"}
		"EQUIVIO_ANALYTICS"				    {$RootLicence = "Office 365 Advanced Compliance"}
		"AAD_BASIC"						    {$RootLicence = "Azure Active Directory Basic"}
		"RMS_S_ENTERPRISE"				    {$RootLicence = "Azure Active Directory Rights Management"}
		"AAD_PREMIUM"					    {$RootLicence = "Azure Active Directory Premium"}
		"MFA_PREMIUM"					    {$RootLicence = "Azure Multi-Factor Authentication"}
		"STANDARDPACK_GOV"				    {$RootLicence = "Microsoft Office 365 (Plan G1) for Government"}
		"STANDARDWOFFPACK_GOV"			    {$RootLicence = "Microsoft Office 365 (Plan G2) for Government"}
		"ENTERPRISEPACK_GOV"				{$RootLicence = "Microsoft Office 365 (Plan G3) for Government"}
		"ENTERPRISEWITHSCAL_GOV"			{$RootLicence = "Microsoft Office 365 (Plan G4) for Government"}
		"DESKLESSPACK_GOV"				    {$RootLicence = "Microsoft Office 365 (Plan K1) for Government"}
		"DESKLESSWOFFPACK_GOV"			    {$RootLicence = "Microsoft Office 365 (Plan K2) for Government"}
		"EXCHANGESTANDARD_GOV"			    {$RootLicence = "Microsoft Office 365 Exchange Online (Plan 1) only for Government"}
		"EXCHANGEENTERPRISE_GOV"			{$RootLicence = "Microsoft Office 365 Exchange Online (Plan 2) only for Government"}
		"SHAREPOINTDESKLESS_GOV"			{$RootLicence = "SharePoint Online Kiosk"}
		"EXCHANGE_S_DESKLESS_GOV"		    {$RootLicence = "Exchange Kiosk"}
		"RMS_S_ENTERPRISE_GOV"			    {$RootLicence = "Windows Azure AD RMS"}
		"OFFICESUBSCRIPTION_GOV"			{$RootLicence = "Office ProPlus"}
		"MCOSTANDARD_GOV"				    {$RootLicence = "Lync Plan 2G"}
		"SHAREPOINTWAC_GOV"				    {$RootLicence = "Office Online for Government"}
		"SHAREPOINTENTERPRISE_GOV"		    {$RootLicence = "SharePoint Plan 2G"}
		"EXCHANGE_S_ENTERPRISE_GOV"		    {$RootLicence = "Exchange Plan 2G"}
		"EXCHANGE_S_ARCHIVE_ADDON_GOV"	    {$RootLicence = "Exchange Online Archiving"}
		"EXCHANGE_S_DESKLESS"			    {$RootLicence = "Exchange Online Kiosk"}
		"SHAREPOINTDESKLESS"				{$RootLicence = "SharePoint Online Kiosk"}
		"SHAREPOINTWAC"					    {$RootLicence=  "Office Online"}
		"YAMMER_ENTERPRISE"				    {$RootLicence = "Yammer for the Starship Enterprise"}
		"EXCHANGE_L_STANDARD"			    {$RootLicence = "Exchange Online (Plan 1)"}
		"MCOLITE"						    {$RootLicence = "Lync Online (Plan 1)"}
		"SHAREPOINTLITE"					{$RootLicence = "SharePoint Online (Plan 1)"}
		"OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ"{$RootLicence = "Office ProPlus"}
		"EXCHANGE_S_STANDARD_MIDMARKET"	    {$RootLicence = "Exchange Online (Plan 1)"}
		"MCOSTANDARD_MIDMARKET"			    {$RootLicence = "Lync Online (Plan 1)"}
		"SHAREPOINTENTERPRISE_MIDMARKET"	{$RootLicence = "SharePoint Online (Plan 1)"}
		"OFFICESUBSCRIPTION"				{$RootLicence = "Office ProPlus"}
		"YAMMER_MIDSIZE"					{$RootLicence = "Yammer"}
		"DYN365_ENTERPRISE_PLAN1"		    {$RootLicence = "Dyn 365 Customer Engage Ent Ed"}
		"MCOSTANDARD"					    {$RootLicence = "SFBO Plan 2"}
		"PROJECT_MADEIRA_PREVIEW_IW_SKU"	{$RootLicence = "Dynamics 365 for Financials for IWs"}
		"STANDARDWOFFPACK_IW_STUDENT"	    {$RootLicence = "Office 365 Education for Students"}
		"STANDARDWOFFPACK_IW_FACULTY"	    {$RootLicence = "Office 365 Education for Faculty"}
		"EOP_ENTERPRISE_FACULTY"			{$RootLicence = "Exchange Online Protection for Faculty"}
		"EXCHANGESTANDARD_STUDENT"		    {$RootLicence = "Exchange Online (Plan 1) for Students"}
		"OFFICESUBSCRIPTION_STUDENT"		{$RootLicence = "Office ProPlus Student Benefit"}
		"STANDARDWOFFPACK_FACULTY"		    {$RootLicence = "Office 365 Education E1 for Faculty"}
		"STANDARDWOFFPACK_STUDENT"		    {$RootLicence = "Microsoft Office 365 (Plan A2) for Students"}
		"DYN365_FINANCIALS_BUSINESS_SKU"	{$RootLicence = "Dyn 365 Financials Business Edition"}
		"DYN365_FINANCIALS_TEAM_MEMBERS_SKU"{$RootLicence = "Dyn 365 Team Members Business Edition"}
		"FLOW_FREE"						    {$RootLicence = "Microsoft Flow Free"}
		"POWER_BI_PRO"					    {$RootLicence = "Power BI Pro"}
		"O365_BUSINESS"					    {$RootLicence = "Office 365 Business"}
		"DYN365_ENTERPRISE_SALES"		    {$RootLicence = "Dyn 365 Enterprise Sales"}
		"RIGHTSMANAGEMENT"				    {$RootLicence = "Rights Management"}
		"PROJECTPROFESSIONAL"			    {$RootLicence = "Project Professional"}
		"VISIOONLINE_PLAN1"				    {$RootLicence = "Visio Online Plan 1"}
		"EXCHANGEENTERPRISE"				{$RootLicence = "Exchange Online Plan 2"}
		"DYN365_ENTERPRISE_P1_IW"		    {$RootLicence = "Dyn 365 P1 Trial Info Workers"}
		"DYN365_ENTERPRISE_TEAM_MEMBERS"	{$RootLicence = "Dyn 365 Team Members Ent Ed"}
		"CRMSTANDARD"					    {$RootLicence = "Microsoft Dynamics CRM Online Professional"}
		"EXCHANGEARCHIVE_ADDON"			    {$RootLicence = "O-Archive for Exchange Online"}
		"EXCHANGEDESKLESS"				    {$RootLicence = "Exchange Online Kiosk"}
		"SPZA_IW"						    {$RootLicence = "App Connect"}
		"WINDOWS_STORE"					    {$RootLicence = "Windows Store for Business"}
		"MCOEV"							    {$RootLicence = "Microsoft Phone System"}
		"VIDEO_INTEROP"					    {$RootLicence = "Polycom Skype Meeting Video Interop for Skype for Business"}
		"SPE_E5"							{$RootLicence = "Microsoft 365 E5"}
		"SPE_E3"							{$RootLicence = "Microsoft 365 E3"}
		"ATA"							    {$RootLicence = "Advanced Threat Analytics"}
		"MCOPSTN2"						    {$RootLicence = "Domestic and International Calling Plan"}
		"FLOW_P1"						    {$RootLicence = "Microsoft Flow Plan 1"}
		"FLOW_P2"						    {$RootLicence = "Microsoft Flow Plan 2"}
		"CRMSTORAGE"						{$RootLicence = "Microsoft Dynamics CRM Online Additional Storage"}
		"SMB_APPS"						    {$RootLicence = "Microsoft Business Apps"}
		"MICROSOFT_BUSINESS_CENTER"		    {$RootLicence = "Microsoft Business Center"}
		"DYN365_TEAM_MEMBERS"			    {$RootLicence = "Dynamics 365 Team Members"}
		"STREAM"							{$RootLicence = "Microsoft Stream Trial"}
		"MS_TEAMS_IW"                       {$RootLicence = "Microsoft Teams Trial"}
		"ADALLOM_STANDALONE"                {$RootLicence = "Microsoft Cloud App Security"}
		"POWERAPPS_VIRAL"                   {$RootLicence = "PowerApps Plan 2 Trial"}
		"AX7_USER_TRIAL"                    {$RootLicence = "Dynamics AX7 Trial"}
		"TEAMS_COMMERCIAL_TRIAL"            {$RootLicence = "Teams Commercial Cloud"}
        "SPE_F1"                            {$RootLicence = "Microsoft 365 F1"}
        "FORMS_PRO"                         {$RootLicence = "Forms Pro Trial"}
        "WIN_DEF_ATP"                   	{$RootLicence = "Windows 10 Defender ATP"}
        "ENTERPRISEPACKWITHOUTPROPLUS"      {$RootLicence = "Office 365 E3 No Pro Plus"}
		"Win10_VDA_E3"                      {$RootLicence = "Windows 10 E3"}
		"IDENTITY_THREAT_PROTECTION"		{$RootLicence = "Microsoft 365 E5 Security"}
        "INFORMATION_PROTECTION_COMPLIANCE" {$RootLicence = "Microsoft 365 E5 Compliance"}
		"EMS_FACULTY"						{$RootLicence = "EMS (Plan E3) Faculty"}
		"POWER_BI_PRO_CE"					{$RootLicence = "Power BI Pro"}
		"POWER_BI_PRO_FACULTY"				{$RootLicence = "Power BI Pro Faculty"}
		"POWERFLOW_P2"						{$RootLicence = "Microsoft PowerApps Plan 2"}
		"POWERAPPS_INDIVIDUAL_USER"			{$RootLicence = "PowerApps and Logic Flows"}
		"DYN365_AI_SERVICE_INSIGHTS"		{$RootLicence = "Dyn 365 CSI Trial"}
		"Dynamics_365_for_Operations"		{$RootLicence = "Dyn 365 Unified Operations Plan"}
		"Dynamics_365_Onboarding_SKU"		{$RootLicence = "Dyn 365 for Talent Onboard"}
		"CCIBOTS_PRIVPREV_VIRAL"			{$RootLicence = "Dyn 365 AI for CSVAV"}
        "DYN365_BUSINESS_MARKETING"			{$RootLicence = "Dyn 365 Marketing"}
        "DYN365_RETAIL_TRIAL"			    {$RootLicence = "Dyn 365 Retail Trial"}
        "SKU_Dynamics_365_for_HCM_Trial"	{$RootLicence = "Dyn 365 Talent"}
        "AAD_PREMIUM_P2"			        {$RootLicence = "Azure AD Premium P2"}
        "MCOPSTN1"			                {$RootLicence = "Domestic Calling Plan"}
		"TEAMS_FREE"			            {$RootLicence = "Microsoft Teams (Free)"}
		"TEAMS_EXPLORATORY"					{$RootLicence = "Microsoft Teams Exploratory"}
		default                             {$RootLicence = $licensesku }
	}
	Write-Output $RootLicence
}
$date = get-date -Format yyyyMMdd
$OutputPath = Get-Item $OutputPath
if ($OutputPath.FullName -notmatch '\\$') {
	$excelfilepath = $OutputPath.FullName + "\"
}
else {
	$excelfilepath = $OutputPath.FullName
}
$XLOutput= $excelfilepath + "$CompanyName - $date.xlsx" ## Output file name
$CSVPath = $excelfilepath + $(-join ((65..90) + (97..122) | Get-Random -Count 14 | ForEach-Object {[char]$_}))
$CSVPath = (New-Item -Type Directory -Path $csvpath).FullName
Write-Host "Checking Connection to Office 365"
$test365 = Get-MsolCompanyInformation -ErrorAction silentlycontinue
if ($null -eq $test365) {
	do {
		if ($Office365Credentials) {
		Connect-MsolService -Credential $Office365Credentials
		}
		else {
			Connect-MsolService
		}
		$test365 = Get-MsolCompanyInformation -ErrorAction silentlycontinue
	} while ($null -eq $test365)
}
Write-Host "Connected to Office 365" 
# Get a list of all licences that exist within the tenant 
$licensetype = Get-MsolAccountSku | Where-Object {$_.ConsumedUnits -ge 1}
# Replace the above with the below if only a single SKU is required
#$licensetype = Get-MsolAccountSku | Where-Object {$_.AccountSkuID -like "*Power*"}
# Get all licences for a summary view
get-msolaccountsku | Where-Object {$_.TargetClass -eq "User"} | select-object @{Name = 'AccountLicenseSKU(Friendly)';  Expression = {$(RootLicenceswitch($_.SkuPartNumber))}}, ActiveUnits, ConsumedUnits | export-csv $CSVPath\AllLicences.csv -NoTypeInformation -Delimiter `t
#get all users with licence
Write-Host "Retrieving all licensed users - this may take a while."
$alllicensedusers = Get-MsolUser -All | Where-Object {$_.isLicensed -eq $true}
# Loop through all licence types found in the tenant 
foreach ($license in $licensetype) {    
    # Build and write the Header for the CSV file 
    $headerstring = "DisplayName`tUserPrincipalName`tAccountSku" 
    foreach ($row in $($license.ServiceStatus)) {
		# Build header string
		$thisLicence = componentlicenseswitch([string]($row.ServicePlan.servicename))
        $headerstring = ($headerstring + "`t" + $thisLicence) 
    } 
    Write-Host ("Gathering users with the following subscription: " + $license.accountskuid) 
    # Gather users for this particular AccountSku from pre-existing array of users
    $users = $alllicensedusers | Where-Object {$_.licenses.accountskuid -contains $license.accountskuid} 
	$RootLicence = RootLicenceswitch($($license.SkuPartNumber))
	#$logfile = $CompanyName + "-" +$RootLicence + ".csv"
	$logfile = $CSVpath + "\" +$RootLicence + ".csv"
	Out-File -FilePath $LogFile -InputObject $headerstring -Encoding UTF8 -append
    # Loop through all users and write them to the CSV file 
    foreach ($user in $users) {
        Write-Verbose ("Processing " + $user.displayname) 
        $thislicense = $user.licenses | Where-Object {$_.accountskuid -eq $license.accountskuid} 
        $datastring = ($user.displayname + "`t" + $user.userprincipalname + "`t" + $rootLicence) 
        foreach ($row in $($thislicense.servicestatus)) {
            # Build data string 
			$datastring = ($datastring + "`t" + $($row.provisioningstatus)) 
        }
        Out-File -FilePath $LogFile -InputObject $datastring -Encoding UTF8 -append 
	}
}             
Write-Host ("Merging CSV Files")
Function Merge-CSVFiles {
	$csvFiles = Get-ChildItem ("$CSVPath\*") -Include *.csv
	$Excel = New-Object -ComObject excel.application
	$Excel.visible = $false
	$Excel.sheetsInNewWorkbook = $csvFiles.Count
	$workbooks = $excel.Workbooks.Add()
	$CSVSheet = 1
	Foreach ($CSV in $Csvfiles) {
		$worksheets = $workbooks.worksheets
		$CSVFullPath = $CSV.FullName
		$SheetName = ($CSV.name -split "\.")[0]
		$worksheet = $worksheets.Item($CSVSheet)
		$worksheet.Name = $SheetName
		$TxtConnector = ("TEXT;" + $CSVFullPath)
		$CellRef = $worksheet.Range("A1")
		$Connector = $worksheet.QueryTables.add($TxtConnector,$CellRef)
		$worksheet.QueryTables.item($Connector.name).TextFileOtherDelimiter = "`t"
		$worksheet.QueryTables.item($Connector.name).TextFileParseType  = 1
		$worksheet.QueryTables.item($Connector.name).Refresh()
		$worksheet.QueryTables.item($Connector.name).delete()
		#autofilter
        [void]$worksheet.UsedRange.Autofilter()
        #autofit
        $worksheet.UsedRange.EntireColumn.AutoFit()		
        $CSVSheet++
	}
	$worksheets = $workbooks.worksheets
	$xlTextString = [Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlTextString
	$xlContains = [Microsoft.Office.Interop.Excel.XlContainsOperator]::xlContains
	foreach ($worksheet in $worksheets){
		Write-Host "Freezing panes on "$Worksheet.Name
		$worksheet.Select()
		$worksheet.application.activewindow.splitcolumn = 1
		$worksheet.application.activewindow.splitrow = 1
		$worksheet.application.activewindow.freezepanes = $true
		$rows = $worksheet.UsedRange.Rows.count
		$columns = $worksheet.UsedRange.Columns.count
		$Selection = $worksheet.Range($worksheet.Cells(1,1), $worksheet.Cells($rows,$columns))
		$Selection.Font.Name = "Segoe UI"
		$Selection.Font.Size = 9
		if ($Worksheet.Name -ne "AllLicences") {
			Write-Host "Setting Conditional Formatting on "$Worksheet.Name
			$Selection= $worksheet.Range($worksheet.Cells(2,4), $worksheet.Cells($rows,$columns))
			$Selection.FormatConditions.Add($xlTextString, "", $xlContains, 'Success','Success',0,0) | Out-Null
			$Selection.FormatConditions.Item(1).Interior.ColorIndex = 35
			$Selection.FormatConditions.Item(1).Font.ColorIndex = 51
			$Selection.FormatConditions.Add($xlTextString,"", $xlContains, 'Pending','Pending',0,0) | Out-Null
			$Selection.FormatConditions.Item(2).Interior.ColorIndex = 36
			$Selection.FormatConditions.Item(2).Font.ColorIndex = 12
			$Selection.FormatConditions.Add($xlTextString,"", $xlContains, 'Disabled','Disabled',0,0) | Out-Null
			$Selection.FormatConditions.Item(3).Interior.ColorIndex = 38
			$Selection.FormatConditions.Item(3).Font.ColorIndex = 30
		}
	}
	$workbooks.Worksheets.Item("AllLicences").Select()
	$workbooks.SaveAs($XLOutput,51)
	$workbooks.Saved = $true
	$workbooks.Close()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
	$excel.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
	[System.GC]::Collect()
	[System.GC]::WaitForPendingFinalizers()
}
Merge-CSVFiles -CSVPath $CSVPath -XLOutput $XLOutput | Out-Null
Write-Host ("Tidying up - deleting CSV Files")
Remove-Item $CSVPath -Recurse -Confirm:$false -Force
Write-Host ("CSV Files Deleted")
Write-Host ("Script Completed.  Results available in $XLOutput")
