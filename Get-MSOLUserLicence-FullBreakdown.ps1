#Requires -Modules MSOnline
#Requires -Version 5

<#
	.SYNOPSIS
		Name: Get-MSOLUserLicence-FullBreakdown.ps1
		The purpose of this script is is to export licensing details to excel

	.DESCRIPTION
		This script will log in to Office 365 and then create a license report by SKU, with each component level status for each user, where 1 or more is assigned. This then conditionally formats the output to colours and autofilter.

	.NOTES
		Version 1.32
		Updated: 20201013	V1.32	Redid group based licensing to improve performance.
		Updated: 20201013	V1.31	Added User Enabled column
		Updated: 20200929 	V1.30	Added RMS_Basic
        Updated: 20200929	V1.29	Added components for E5 Compliance
        Updated: 20200929	V1.28	Added code for group assigned and direct assigned licensing
		Updated: 20200820	V1.27	Added additional Office 365 E1 components
        Updated: 20200812	V1.26	Added Links to Licensing Sheets on All Licenses Page and move All Licenses Page to be first worksheet
		Updated: 20200730	V1.25	Added AIP P2 and Project for Office (E3 + E5)
		Updated: 20200720	V1.24	Added Virtual User component
		Updated: 20200718	V1.23	Added AAD Basic friendly component name
		Updated: 20200706   V1.22   Updated SKU error and added additional friendly names
		Updated: 20200626 	V1.21	Updated F1 to F3 as per Microsoft's update
		Updated: 20200625	V1.20	Added Telephony Virtual User
		Updated: 20200603	V1.19	Added Switch for no name translation		
		Updated: 20200603	V1.18	Added Telephony SKU's
		Updated: 20200501	V1.17	Script readability changes
		Updated: 20200430	V1.16	Made script more readable for Product type within component breakdown
		Updated: 20200422	V1.15   Formats to Segoe UI 9pt. Removed unnecessary True output. 
		Updated: 20200408   V1.14   Added Teams Exploratory SKU
		Updated: 20200204   V1.13   Added more SKU's and Components
		Updated: 20191015	V1.12	Tidied up old comments
		Updated: 20190916	V1.11	Added more components and SKU's
		Updated: 20190830   V1.10   Added more components. Updated / renamed refreshed licences
       	Updated: 20190627	V1.09	Added more Components
		Updated: 20190614	V1.08	Added more SKU's and Components
        Updated: 20190602	V1.07	Parameters, Comment based help, creates folder and deletes folder for csv's, require statements

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
	[PSCredential]$Office365Credentials,
	[Parameter(
		HelpMessage="This stops translation into Friendly Names of SKU's and Components"
	)][switch]$NoNameTranslation
)
#Following Function Switches Complicated Component Names with Friendly Names
function componentlicenseswitch {
	param
	(
		[parameter (Mandatory=$true, Position = 1)][string]$component
	)
	switch -wildcard ($($component)) {
		#AAD
		"AAD_BASIC"				{$thisLicence = "Azure Acitve Directory Basic"}
		"AAD_PREMIUM"           {$thisLicence = "Azure Active Directory Premium P1"}
		"AAD_PREMIUM_P2"        {$thisLicence = "Azure Active Directory Premium P2"}

        #Dynamics
		"DYN365_ENTERPRISE_SALES" {$thisLicence = "Dynamics 365 for Sales"}
		"Dynamics_365_for_Talent_Team_members" {$thisLicence = "Dynamics 365 for Talent Team members"}
		"Dynamics_365_for_Retail_Team_members" {$thisLicence = "Dynamics 365 for Retail Team members"}
		"DYN365_Enterprise_Talent_Onboard_TeamMember" {$thisLicence = "Dynamics 365 for Talent - Onboard Experience"}
		"DYN365_Enterprise_Talent_Attract_TeamMember" {$thisLicence = "Dynamics 365 for Talent - Attract Experience Team Member"}
		"Dynamics_365_for_Operations_Team_members" {$thisLicence = "Dynamics_365_for_Operations_Team_members"}
		"DYN365_ENTERPRISE_TEAM_MEMBERS"{$thisLicence = "Dynamics 365 for Team Members"}
		"CCIBOTS_PRIVPREV_VIRAL"	{$thisLicence = "Dynamics 365 AI for Customer Service Virtual Agents Viral"}
		"DYN365_CDS_CCI_BOTS"		{$thisLicence = "Common Data Service for CCI Bots"}
		"DYN365_AI_SERVICE_INSIGHS"	{$thisLicence = "Dynamics 365 Customer Service Insights"}
		"POWERAPPS_DYN_TEAM"		{$thisLicence = "PowerApps for Dynamics 365"}
		"FLOW_DYN_TEAM"				{$thisLicence = "Flow for Dynamics 365"}
		"DYN365_TEAM_MEMBERS"		{$thisLicence = "Dynamics 365 Team Members"}
		"DYN365_BUSINESS_Marketing"	{$thisLicence = "Dynamics 365 Marketing"}
		"DYN365_RETAIL_TRIAL"		{$thisLicence = "Dynamics 365 Retail Trial"}
		"Dynamics_365_for_Retail"	{$thisLicence = "Dynamics 365 for Retail"}
		"DYN365_TALENT_ENTERPRISE"	{$thisLicence = "Dynamics 365 for Talent"}
		"ERP_TRIAL_INSTANCE"      		{$thisLicence = "AX7 Instance"}
		"Dynamics_365_Talent_Onboard"			{$thisLicence = "Dynamics 365 for Talent: Onboard"}
		"Dynamics_365_Onboarding_Free_PLAN"	    {$thisLicence = "Dynamics 365 for Talent: Onboard"}
		"Dynamics_365_Hiring_Free_PLAN"		    {$thisLicence = "Dynamics 365 for Talent: Attract"}
		"Dynamics_365_for_HCM_Trial"		    {$thisLicence = "Dynamics_365_for_HCM_Trial"}
		"DYN365_ENTERPRISE_P1"		{$thisLicence = "Dynamics Enterprise P1"}
		"D365_CSI_EMBED_CE" 		{$thisLicence = "Dynamics 365 Customer Service Insights for CE Plan"}
		"DYN365_ENTERPRISE_P1_IW"	{$thisLicence = "Dyn 365 P1 Trial Info Workers"}

		#Dynamics Common Data Service
		"DYN365_CDS_O365_P1"			{$thisLicence = "Common Data Service"}
		"DYN365_CDS_O365_P2"			{$thisLicence = "Common Data Service"}
		"DYN365_CDS_O365_P3"			{$thisLicence = "Common Data Service"}
		"DYN365_CDS_O365_F1"			{$thisLicence = "Common Data Service"}
		"DYN365_CDS_P1"				{$thisLicence = "Common Data Service"}
		"DYN365_CDS_P2"				{$thisLicence = "Common Data Service"}
		"DYN365_CDS_FORMS_PRO"		{$thisLicence = "Common Data Service"}   
		"DYN365_CDS_DYN_APPS"		{$thisLicence = "Common Data Service"}
		"DYN365_CDS_DYN_P2"			{$thisLicence = "Common Data Service"}
		"DYN365_CDS_VIRAL"     {$thisLicence = "Common Data Service"}
        "CDS_O365_P1"     {$thisLicence = "Common Data Service for Teams"}
        "CDS_O365_P2"     {$thisLicence = "Common Data Service for Teams"}
        "CDS_O365_P3"     {$thisLicence = "Common Data Service for Teams"}

		#Exchange
		"EXCHANGE_S_ENTERPRISE"{$thisLicence = "Exchange Online (Plan 2)"}
		"EXCHANGE_S_FOUNDATION"{$thisLicence = "Core Exchange for non-Exch SKUs (e.g. setting profile pic)"}
		"EXCHANGE_S_DESKLESS"      	{$thisLicence = "Exchange Online Firstline"}
		"EXCHANGE_S_STANDARD"      	{$thisLicence = "Exchange Online (Plan 1)"}
		"EXCHANGE_S_ARCHIVE_ADDON"	{$thisLicence = "Exchange Online Archiving Add-on"}

		#Flow
		"FLOW_P1"		{$thisLicence = "Microsoft Flow Plan 1"}
		"FLOW_P2"	        {$thisLicence = "Microsoft Flow Plan 2"}
		"FLOW_O365_P1"          {$thisLicence = "Flow for Office 365"}
		"FLOW_O365_P2"          {$thisLicence = "Flow for Office 365"}
		"FLOW_O365_P3"          {$thisLicence = "Flow for Office 365"}
		"FLOW_DYN_APPS"         {$thisLicence = "Flow for Dynamics 365"}
		"FLOW_P2_VIRAL"         {$thisLicence = "Flow Free"}
		"FLOW_P2_VIRAL_REAL"    {$thisLicence = "Flow P2 Viral"}
		"FLOW_CCI_BOTS"		{$thisLicence = "Flow for CCI Bots"}
		"Forms_Pro_CE"		{$thisLicence = "Forms Pro for Customer Engagement Plan"}
		"FORMS_PRO"		{$thisLicence = "Forms Pro"}
		"FLOW_FORMS_PRO"	{$thisLicence = "Flow for Forms Pro"}
		"FLOW_O365_S1"      	{$thisLicence = "Flow for Office 365 (F1)"}
		"FLOW_DYN_P2"		{$thisLicence = "Flow for Dynamics 365"}

		#Forms
		"FORMS_PLAN_E1"         {$thisLicence = "Microsoft Forms (Plan E1)"}
		"FORMS_PLAN_E3"         {$thisLicence = "Microsoft Forms (Plan E3)"}
		"FORMS_PLAN_E5"         {$thisLicence = "Microsoft Forms E5"}
		"Forms_Pro_Operations" 		{$thisLicence = "Microsoft Forms Pro for Operations"}
		"Forms_Pro_CE"				{$thisLicence = "Forms Pro for Customer Engagement Plan"}
		"FORMS_PRO"			        {$thisLicence = "Forms Pro"}
		"FORMS_PLAN_K"      		{$thisLicence = "Microsoft Forms (Plan F1)"}

		#Kaizala
		"KAIZALA_STANDALONE"		{$thisLicence = "Microsoft Kaizala Pro"}
		"KAIZALA_O365_P1"		    {$thisLicence = "Microsoft Kaizala Pro (P1)"}
		"KAIZALA_O365_P2"		    {$thisLicence = "Microsoft Kaizala Pro"}
		"KAIZALA_O365_P3"			{$thisLicence = "Kaizala for Office 365"}

		#Misc Services
		"MYANALYTICS_P2"       {$thisLicence = "Insights by MyAnalytics"}
		"EXCHANGE_ANALYTICS"   {$thisLicence = "Microsoft MyAnalytics (Full)"}
		"Deskless"             {$thisLicence = "Microsoft StaffHub"}
		"SWAY"                 {$thisLicence = "Sway"}
		"PROJECTWORKMANAGEMENT"{$thisLicence = "Microsoft Planner"}
		"YAMMER_ENTERPRISE"    {$thisLicence = "Yammer Enterprise"}
		"SPZA"                 {$thisLicence = "App Connect"}
		"MICROSOFT_BUSINESS_CENTER" {$thisLicence = "Microsoft Business Center"}
		"NBENTERPRISE"         {$thisLicence = "Microsoft Social Engagement - Service Discontinuation"}
		"MICROSOFT_SEARCH"      	{$thisLicence = "Microsoft Search"}
		"MICROSOFTBOOKINGS"      	{$thisLicence = "Microsoft Bookings"}
		"EXCEL_PREMIUM"      	{$thisLicence = "Microsoft Excel Advanced Analytics"}
		
        #Office
		"SHAREPOINTWAC"        {$thisLicence = "Office Online"}	
		"OFFICESUBSCRIPTION"   {$thisLicence = "Office 365 ProPlus"}
		"OFFICEMOBILE_SUBSCRIPTION" {$thisLicence = "Office Mobile Apps for Office 365"}

		#OneDrive
		"ONEDRIVESTANDARD"			{$thisLicence = "OneDrive for Business (Plan 1)"}
		"ONEDRIVE_BASIC"      			{$thisLicence = "OneDrive Basic"}

		#PowerBI
		"BI_AZURE_P0"           {$thisLicence = "Power BI (Free)"}
		"BI_AZURE_P2"           {$thisLicence = "Power BI Pro"}

		#Phone System
		"MCOEV"                {$thisLicence = "M365 Phone System"}
		"MCOMEETADV"           {$thisLicence = "M365 Audio Conferencing"}
		"MCOEV_VIRTUALUSER"	   {$thisLicence = "Microsoft 365 Phone System Virtual User"}

		#PowerApps
		"POWERAPPS_O365_S1"         {$thisLicence = "PowerApps for Office 365 Firstline"}
		"POWERAPPS_O365_P1"    {$thisLicence = "PowerApps for Office 365"}
		"POWERAPPS_O365_P2"    {$thisLicence = "PowerApps for Office 365"}
		"POWERAPPS_O365_P3"    {$thisLicence = "PowerApps for Office 365"}
		"POWERAPPS_DYN_APPS"   {$thisLicence = "PowerApps for Dynamics 365"}
		"POWERAPPS_P2_VIRAL"   {$thisLicence = "PowerApps Plan 2 Trial"}
		"POWERAPPS_P2"				{$thisLicence = "PowerApps Plan 2"}
		"POWERAPPS_DYN_P2"			{$thisLicence = "PowerApps for Dynamics 365"}
        "POWER_VIRTUAL_AGENTS_O365_P1"	{$thisLicence = "Power Virtual Agents for Office 365"}
        "POWER_VIRTUAL_AGENTS_O365_P2"	{$thisLicence = "Power Virtual Agents for Office 365"}
        "POWER_VIRTUAL_AGENTS_O365_P3"	{$thisLicence = "Power Virtual Agents for Office 365"}

		#Project
		"PROJECT_PROFESSIONAL"      	{$thisLicence = "Project P3"}
		"FLOW_FOR_PROJECT"      		{$thisLicence = "Data Integration for Project with Flow"}
		"DYN365_CDS_PROJECT"      		{$thisLicence = "Common Data Service for Project"}
		"SHAREPOINT_PROJECT"      		{$thisLicence = "Project Online Service"}
		"PROJECT_CLIENT_SUBSCRIPTION"	{$thisLicence = "Project Online Desktop Client"}
		"PROJECT_ESSENTIALS"            {$thisLicence = "Project Online Essentials"}
		"PROJECT_O365_P1"				{$thisLicence = "Project for Office (Plan E1)"}
        "PROJECT_O365_P2"				{$thisLicence = "Project for Office (Plan E3)"}
		"PROJECT_O365_P3"				{$thisLicence = "Project for Office (Plan E5)"}

		#Security & Compliance
		"RECORDS_MANAGEMENT"		{$thisLicence = "Microsoft Records Management"}
		"INFO_GOVERNANCE"			{$thisLicence = "Microsoft Information Governance"}
		"DATA_INVESTIGATIONS"		{$thisLicence = "Microsoft Data Investigations"}
		"CUSTOMER_KEY"			    {$thisLicence = "Microsoft Customer Key"}
		"COMMUNICATIONS_DLP"		{$thisLicence = "Microsoft Communications DLP"}
		"COMMUNICATIONS_COMPLIANCE"	{$thisLicence = "Microsoft Communications Compliance"}
		"M365_ADVANCED_AUDITING" {$thisLicence = "Microsoft 365 Advanced Auditing"}
		"ATP_ENTERPRISE"        {$thisLicence = "O365 ATP Plan 1 (not licenced individually)"}
		"THREAT_INTELLIGENCE"   {$thisLicence = "O365 ATP Plan 2"}
		"ADALLOM_S_O365"        {$thisLicence = "Office 365 Cloud App Security"}
		"EQUIVIO_ANALYTICS"     {$thisLicence = "Office 365 Advanced eDiscovery"}
		"LOCKBOX_ENTERPRISE"    {$thisLicence = "Customer Lockbox"}
		"ATA"                   {$thisLicence = "Azure Advanced Threat Protection"}
		"ADALLOM_S_STANDALONE"  {$thisLicence = "Microsoft Cloud App Security"}
		"RMS_S_BASIC"			{$thisLicence = "Azure Rights Management Service (non-assignable)"}
		"RMS_S_ENTERPRISE"      {$thisLicence = "Azure Rights Management"}
		"RMS_S_PREMIUM"         {$thisLicence = "Azure Information Protection Premium P1"} 
		"RMS_S_PREMIUM2"        {$thisLicence = "Azure Information Protection Premium P2"}
		"RMS_S_ADHOC"      			{$thisLicence = "Rights Management Adhoc"}
		"INTUNE_A"              {$thisLicence = "Microsoft Intune"}
        "INTUNE_A_VL"           {$thisLicence = "Microsoft Intune"}
		"MFA_PREMIUM"           {$thisLicence = "Microsoft Azure Multi-Factor Authentication"}
		"INTUNE_O365"           {$thisLicence = "MDM for Office 365 (not licenced individually)"}
		"PAM_ENTERPRISE"        {$thisLicence = "O365 Priviledged Access Management"}
		"ADALLOM_S_DISCOVERY"   {$thisLicence = "Cloud App Security Discovery"}
		"MIP_S_CLP1"             {$thisLicence = "Information Protection for Office 365 - Standard"}
		"MIP_S_CLP2"            {$thisLicence = "Information Protection for Office 365 - Premium"}
		"PREMIUM_ENCRYPTION"      	{$thisLicence = "Premium Encryption"}
		"INFORMATION_BARRIERS"		{$thisLicence = "Information Barriers"}
		"WINDEFATP"                	{$thisLicence = "Windows Defender ATP"}
		"MTP"                	{$thisLicence = "Microsoft Threat Protection"}
        "Content_Explorer"					{$thisLicence = "Content Explorer (Assigned at Org Level)"}
        "MICROSOFTENDPOINTDLP"				{$thisLicence = "Microsoft Endpoint DLP"}
        "INSIDER_RISK"   					{$thisLicence = "Microsoft Insider Risk Management"}
        "INSIDER_RISK_MANAGEMENT"			{$thisLicence = "RETIRED - Microsoft Insider Risk Management"}
		"ML_CLASSIFICATION"   				{$thisLicence = "Microsoft ML_based Classification"}

		#SharePoint
		"SHAREPOINTDESKLESS"   {$thisLicence = "SharePoint Online Kiosk"}
		"SHAREPOINTSTANDARD"   {$thisLicence = "SharePoint Online (Plan 1)"}
		"SHAREPOINTENTERPRISE" {$thisLicence = "SharePoint Online (Plan 2)"}

		#Skype
		"MCOIMP"                    {$thisLicence = "Skype for Business (Plan 1)"}
		"MCOSTANDARD"          {$thisLicence = "Skype for Business Online (Plan 2)"}

		#Stream
		"STREAM_O365_K"        {$thisLicence = "Stream for Office 365 Firstline"}
		"STREAM_O365_E1"       {$thisLicence = "Microsoft Stream for O365 E1"}
		"STREAM_O365_E3"       {$thisLicence = "Microsoft Stream for O365 E3 SKU"}
		"STREAM_O365_E5"       {$thisLicence = "Stream E5"}

		#Teams
		"TEAMS1"               {$thisLicence = "Microsoft Teams"}
		"MCO_TEAMS_IW"         {$thisLicence = "Microsoft Teams Trial"}
		"TEAMS_FREE_SERVICE"			{$thisLicence = "Teams Free Service (Not assigned per user)"}
		"MCOFREE"			        {$thisLicence = "MCO Free for Microsoft Teams (free)"}
		"TEAMS_FREE"			    {$thisLicence = "Microsoft Teams (free)"}

        #Telephony
        "MCOPSTN1"                  {$thisLicence = "Domestic Calling Plan (1200 min)"}
        "MCOPSTN2"                  {$thisLicence = "Domestic and International Calling Plan"}
		"MCOPSTN5"                  {$thisLicence = "Domestic Calling Plan (120 min)"}
		"PHONESYSTEM_VIRTUALUSER"	{$thisLicence = "M365 Phone System - Virtual User"}

		#To-Do
		"BPOS_S_TODO_FIRSTLINE"     {$thisLicence = "To-Do Firstline"}
		"BPOS_S_TODO_1"      	    {$thisLicence = "To-Do Plan 1"}
		"BPOS_S_TODO_2"             {$thisLicence = "To-Do (Plan 2)"}
		"BPOS_S_TODO_3"             {$thisLicence = "To-Do (Plan 3)"}

		#Visio
		"VISIOONLINE"      		{$thisLicence = "Visio Online"}
		"VISIO_CLIENT_SUBSCRIPTION" 	{$thisLicence = "Visio Pro for Office 365"}

		#Whiteboard
		"WHITEBOARD_FIRSTLINE1"         {$thisLicence = "Whiteboard for Firstline"}
		"WHITEBOARD_PLAN1"      	{$thisLicence = "Whiteboard Plan 1"}
		"WHITEBOARD_PLAN2"      	{$thisLicence = "Whiteboard Plan 2"}
		"WHITEBOARD_PLAN3"      	{$thisLicence = "Whiteboard Plan 3"}

		#Windows 10
		"WIN10_PRO_ENT_SUB"      	{$thisLicence = "Win 10 Enterprise E3"}
		"WIN10_ENT_LOC_F1"              {$thisLicence = "Win 10 Enterprise E3 (Local Only)"}                
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
        #Azure AD
		"AAD_BASIC"						    {$RootLicence = "Azure Active Directory Basic"}
		"AAD_PREMIUM"					    {$RootLicence = "Azure Active Directory Premium"}
		"AAD_PREMIUM_P2"			        {$RootLicence = "Azure AD Premium P2"}
        #Dynamics
		"DYN365_ENTERPRISE_PLAN1"		    {$RootLicence = "Dyn 365 Customer Engage Ent Ed"}
		"PROJECT_MADEIRA_PREVIEW_IW_SKU"	{$RootLicence = "Dynamics 365 for Financials for IWs"}
        "DYN365_AI_SERVICE_INSIGHTS"		{$RootLicence = "Dyn 365 CSI Trial"}
		"Dynamics_365_for_Operations"		{$RootLicence = "Dyn 365 Unified Operations Plan"}
		"Dynamics_365_Onboarding_SKU"		{$RootLicence = "Dyn 365 for Talent Onboard"}
		"CCIBOTS_PRIVPREV_VIRAL"			{$RootLicence = "Dyn 365 AI for CSVAV"}
        "DYN365_BUSINESS_MARKETING"			{$RootLicence = "Dyn 365 Marketing"}
        "DYN365_RETAIL_TRIAL"			    {$RootLicence = "Dyn 365 Retail Trial"}
        "SKU_Dynamics_365_for_HCM_Trial"	{$RootLicence = "Dyn 365 Talent"}
        "DYN365_FINANCIALS_BUSINESS_SKU"	{$RootLicence = "Dyn 365 Financials Business Edition"}
		"DYN365_FINANCIALS_TEAM_MEMBERS_SKU"{$RootLicence = "Dyn 365 Team Members Business Edition"}
        "AX7_USER_TRIAL"                    {$RootLicence = "Dynamics AX7 Trial"}
        "DYN365_ENTERPRISE_P1_IW"		    {$RootLicence = "Dyn 365 P1 Trial Info Workers"}
        "DYN365_ENTERPRISE_TEAM_MEMBERS"	{$RootLicence = "Dyn 365 Team Members Ent Ed"}
        "DYN365_TEAM_MEMBERS"			    {$RootLicence = "Dynamics 365 Team Members"}
        "CRMSTORAGE"						{$RootLicence = "Microsoft Dynamics CRM Online Additional Storage"}
        "CRMSTANDARD"					    {$RootLicence = "Microsoft Dynamics CRM Online Professional"}
        "DYN365_ENTERPRISE_SALES"		    {$RootLicence = "Dyn 365 Enterprise Sales"}
        #Exchange
        "EXCHANGESTANDARD_GOV"			    {$RootLicence = "Microsoft Office 365 Exchange Online (Plan 1) only for Government"}
        "EXCHANGEENTERPRISE_GOV"			{$RootLicence = "Microsoft Office 365 Exchange Online (Plan 2) only for Government"}
        "EXCHANGE_S_DESKLESS_GOV"		    {$RootLicence = "Exchange Kiosk"}
        "ECAL_SERVICES"					    {$RootLicence = "ECAL"}
        "EXCHANGE_S_ENTERPRISE_GOV"		    {$RootLicence = "Exchange Plan 2G"}
        "EXCHANGE_S_ARCHIVE_ADDON_GOV"	    {$RootLicence = "Exchange Online Archiving"}
        "EXCHANGE_S_DESKLESS"			    {$RootLicence = "Exchange Online Kiosk"}
        "EXCHANGE_L_STANDARD"			    {$RootLicence = "Exchange Online (Plan 1)"}
        "EXCHANGE_S_STANDARD_MIDMARKET"	    {$RootLicence = "Exchange Online (Plan 1)"}
        "EXCHANGESTANDARD"				    {$RootLicence = "Exchange Online (Plan 1)"}
        "EXCHANGEENTERPRISE"				{$RootLicence = "Exchange Online Plan 2"}
        "EOP_ENTERPRISE_FACULTY"			{$RootLicence = "Exchange Online Protection for Faculty"}
        "EXCHANGESTANDARD_STUDENT"		    {$RootLicence = "Exchange Online (Plan 1) for Students"}
        "EXCHANGEARCHIVE_ADDON"			    {$RootLicence = "O-Archive for Exchange Online"}
        "EXCHANGEDESKLESS"				    {$RootLicence = "Exchange Online Kiosk"}
        #Flow
        "FLOW_FREE"						    {$RootLicence = "Microsoft Flow Free"}
        "FLOW_P1"						    {$RootLicence = "Microsoft Flow Plan 1"}
        "FLOW_P2"						    {$RootLicence = "Microsoft Flow Plan 2"}
        #Forms
        "FORMS_PRO"                         {$RootLicence = "Forms Pro Trial"}
        #Microsoft 365 Subscription
        "STANDARDPACK_GOV"				    {$RootLicence = "Microsoft Office 365 (Plan G1) for Government"}
		"STANDARDWOFFPACK_GOV"			    {$RootLicence = "Microsoft Office 365 (Plan G2) for Government"}
		"ENTERPRISEPACK_GOV"				{$RootLicence = "Microsoft Office 365 (Plan G3) for Government"}
		"ENTERPRISEWITHSCAL_GOV"			{$RootLicence = "Microsoft Office 365 (Plan G4) for Government"}
		"DESKLESSPACK_GOV"				    {$RootLicence = "Microsoft Office 365 (Plan K1) for Government"}
        "DESKLESSWOFFPACK_GOV"			    {$RootLicence = "Microsoft Office 365 (Plan K2) for Government"}
		"SPE_E3"							{$RootLicence = "Microsoft 365 E3"}
        "SPE_E5"							{$RootLicence = "Microsoft 365 E5"}
        "SPE_F1"                            {$RootLicence = "Microsoft 365 D1"}
        "STANDARDWOFFPACK_STUDENT"		    {$RootLicence = "Microsoft Office 365 (Plan A2) for Students"}
        #Misc Services
		"PLANNERSTANDALONE"				    {$RootLicence = "Planner Standalone"}
		"CRMIUR"							{$RootLicence = "CMRIUR"}
		"PROJECTWORKMANAGEMENT"			    {$RootLicence = "Office 365 Planner Preview"}
        "STREAM"							{$RootLicence = "Microsoft Stream Trial"}
        "SPZA_IW"						    {$RootLicence = "App Connect"}
        #Office 365 Subscription
        "O365_BUSINESS"					    {$RootLicence = "Office 365 Business"}
        "O365_BUSINESS_ESSENTIALS"		    {$RootLicence = "Office 365 Business Essentials"}
        "O365_BUSINESS_PREMIUM"			    {$RootLicence = "Office 365 Business Premium"}
		"DESKLESSPACK"					    {$RootLicence = "Office 365 (Plan F3)"}
		"DESKLESSWOFFPACK"				    {$RootLicence = "Office 365 (Plan K2)"}
		"LITEPACK"						    {$RootLicence = "Office 365 (Plan P1)"}
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
        "STANDARDWOFFPACK_IW_STUDENT"	    {$RootLicence = "Office 365 Education for Students"}
        "STANDARDWOFFPACK_IW_FACULTY"	    {$RootLicence = "Office 365 Education for Faculty"}
        "STANDARDWOFFPACK_FACULTY"		    {$RootLicence = "Office 365 Education E1 for Faculty"}
        "ENTERPRISEPACKWITHOUTPROPLUS"      {$RootLicence = "Office 365 E3 No Pro Plus"}
        #Office Suite
		"OFFICESUBSCRIPTION_GOV"			{$RootLicence = "Office ProPlus"}
		"SHAREPOINTWAC_GOV"				    {$RootLicence = "Office Online for Government"}
		"SHAREPOINTWAC"					    {$RootLicence=  "Office Online"}
        "OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ"{$RootLicence = "Office ProPlus"}
        "OFFICESUBSCRIPTION"				{$RootLicence = "Office ProPlus"}
        "OFFICESUBSCRIPTION_STUDENT"		{$RootLicence = "Office ProPlus Student Benefit"}
        #PowerApps
		"POWERFLOW_P1"						{$RootLicence = "Microsoft PowerApps Plan 1"}
		"POWERFLOW_P2"						{$RootLicence = "Microsoft PowerApps Plan 2"}
        "POWERAPPS_INDIVIDUAL_USER"			{$RootLicence = "PowerApps and Logic Flows"}
        "POWERAPPS_VIRAL"                   {$RootLicence = "PowerApps Plan 2 Trial"}
        #Power BI
        "POWER_BI_ADDON"					{$RootLicence = "Office 365 Power BI Addon"}
		"POWER_BI_INDIVIDUAL_USE"		    {$RootLicence = "Power BI Individual User"}
		"POWER_BI_STANDALONE"			    {$RootLicence = "Power BI Stand Alone"}
        "POWER_BI_STANDARD"				    {$RootLicence = "Power BI Free"}
        "BI_AZURE_P1"					    {$RootLicence = "Power BI Reporting and Analytics"}
        "POWER_BI_PRO"					    {$RootLicence = "Power BI Pro"}
		"POWER_BI_PRO_CE"					{$RootLicence = "Power BI Pro"}
        "POWER_BI_PRO_FACULTY"				{$RootLicence = "Power BI Pro Faculty"}
        #Project
        "PROJECTESSENTIALS"				    {$RootLicence = "Project Lite"}
		"PROJECTCLIENT"					    {$RootLicence = "Project Professional"}
		"PROJECTONLINE_PLAN_1"			    {$RootLicence = "Project Online"}
		"PROJECTONLINE_PLAN_2"			    {$RootLicence = "Project Online and PRO"}
		"ProjectPremium"					{$RootLicence = "Project Online Premium"}
        "PROJECTPROFESSIONAL"			    {$RootLicence = "Project Professional"}
        #Security and Compliance
        "EMS"							    {$RootLicence = "EMS (Plan E3)"}
        "EMSPREMIUM"                        {$RootLicence = "EMS (Plan E5)"}
        "RIGHTSMANAGEMENT_ADHOC"			{$RootLicence = "Windows Azure RMS"}
        "INTUNE_A"						    {$RootLicence = "Microsoft Intune"}
        "INTUNE_A_VL"                       {$RootLicence = "Microsoft Intune"}
        "ATP_ENTERPRISE"					{$RootLicence = "Office 365 ATP Plan 1"}
        "THREAT_INTELLIGENCE"               {$RootLicence = "Office 365 ATP Plan 2"}
        "EQUIVIO_ANALYTICS"				    {$RootLicence = "Office 365 Advanced Compliance"}
        "RMS_S_ENTERPRISE"				    {$RootLicence = "Azure Active Directory Rights Management"}
        "MFA_PREMIUM"					    {$RootLicence = "Azure Multi-Factor Authentication"}
        "RMS_S_ENTERPRISE_GOV"			    {$RootLicence = "Windows Azure AD RMS"}
        "IDENTITY_THREAT_PROTECTION"		{$RootLicence = "Microsoft 365 E5 Security"}
        "INFORMATION_PROTECTION_COMPLIANCE" {$RootLicence = "Microsoft 365 E5 Compliance"}
        "EMS_FACULTY"						{$RootLicence = "EMS (Plan E3) Faculty"}
        "ADALLOM_STANDALONE"                {$RootLicence = "Microsoft Cloud App Security"}
        "ATA"							    {$RootLicence = "Advanced Threat Analytics"}
        "WIN_DEF_ATP"                   	{$RootLicence = "Windows 10 Defender ATP"}
		"RIGHTSMANAGEMENT"				    {$RootLicence = "Rights Management"}
		"INFOPROTECTION_P2"					{$RootLicence = "AIP Premium P2"}

        #Skype
        "MCOSTANDARD_GOV"				    {$RootLicence = "Lync Plan 2G"}
        "MCOLITE"						    {$RootLicence = "Lync Online (Plan 1)"}
        "MCOSTANDARD_MIDMARKET"			    {$RootLicence = "Lync Online (Plan 1)"}
        "MCOSTANDARD"					    {$RootLicence = "SFBO Plan 2"}
        "VIDEO_INTEROP"					    {$RootLicence = "Polycom Video Interop"}
        #SharePoint
        "SHAREPOINTSTORAGE"				    {$RootLicence = "SharePoint storage"}
        "SHAREPOINTDESKLESS_GOV"			{$RootLicence = "SharePoint Online Kiosk"}
        "SHAREPOINTENTERPRISE_GOV"		    {$RootLicence = "SharePoint Plan 2G"}
        "SHAREPOINTDESKLESS"				{$RootLicence = "SharePoint Online Kiosk"}
        "SHAREPOINTLITE"					{$RootLicence = "SharePoint Online (Plan 1)"}
        "SHAREPOINTENTERPRISE_MIDMARKET"	{$RootLicence = "SharePoint Online (Plan 1)"}
        #Teams
		"TEAMS_FREE"			            {$RootLicence = "Microsoft Teams (Free)"}
        "TEAMS_EXPLORATORY"					{$RootLicence = "Microsoft Teams Exploratory"}
        "TEAMS_COMMERCIAL_TRIAL"            {$RootLicence = "Teams Commercial Cloud"}
        "MS_TEAMS_IW"                       {$RootLicence = "Microsoft Teams Trial"}
		"MEETING_ROOM"                      {$RootLicence = "Meeting Room"}
		"MCOCAP"							{$RootLicence = "Common Area Phone"}
        #Telephony
        "MCOMEETADV"						{$RootLicence = "PSTN conferencing"}
        "MCOPSTN1"			                {$RootLicence = "Dom Calling Plan (1200 mins)"}
        "MCOPSTN2"						    {$RootLicence = "Dom and Intl Calling Plan"}
        "MCOEV"							    {$RootLicence = "Microsoft Phone System"}
		"MCOPSTN_5"                         {$RootLicence = "Dom Calling Plan (120mins)"}
		"PHONESYSTEM_VIRTUALUSER"			{$RootLicence = "Phone System Virtual User"}
        #Visio
        "VISIOCLIENT"					    {$RootLicence = "Visio Pro Online"}
        "VISIOONLINE_PLAN1"				    {$RootLicence = "Visio Online Plan 1"}
        #Windows 10
        "Win10_VDA_E3"                      {$RootLicence = "Windows 10 E3"}
        "WINDOWS_STORE"					    {$RootLicence = "Windows Store for Business"}
		"SMB_APPS"						    {$RootLicence = "Microsoft Business Apps"}
        "MICROSOFT_BUSINESS_CENTER"		    {$RootLicence = "Microsoft Business Center"}
        #Yammer
		"YAMMER_ENTERPRISE"				    {$RootLicence = "Yammer Enterprise"}
        "YAMMER_MIDSIZE"					{$RootLicence = "Yammer"}
		default                             {$RootLicence = $licensesku }
	}
	Write-Output $RootLicence
}
#Helper function for tidier select of Groups for Group Based Licensing
Function Invoke-GroupGuidConversion { 
	[CmdletBinding()]
	param (
		[Parameter(Mandatory)]
		[String[]]
		$GroupGuid,
		[Parameter(Mandatory)]
		[Object[]]
		$LicenseGroups
	)
	$output = @()
	foreach ($guid in $GroupGuid) {
		$temp = [PSCustomObject]@{
			DisplayName = ($LicenseGroups | Where-Object {$_.ObjectID -eq $guid}).Displayname
		}	
		$output += $temp
		Remove-Variable temp
	}
	Write-Output $output
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
if ($NoNameTranslation) {
	get-msolaccountsku | Where-Object {$_.TargetClass -eq "User"} | select-object @{Name = 'AccountLicenseSKU';  Expression = {$($_.SkuPartNumber)}}, ActiveUnits, ConsumedUnits | Sort-Object 'AccountLicenseSKU' | export-csv $CSVPath\AllLicences.csv -NoTypeInformation -Delimiter `t
}
else {
	get-msolaccountsku | Where-Object {$_.TargetClass -eq "User"} | select-object @{Name = 'AccountLicenseSKU(Friendly)';  Expression = {$(RootLicenceswitch($_.SkuPartNumber))}}, ActiveUnits, ConsumedUnits | Sort-Object 'AccountLicenseSKU(Friendly)' | export-csv $CSVPath\AllLicences.csv -NoTypeInformation -Delimiter `t
}
#get all users with licence
Write-Host "Retrieving all licensed users - this may take a while."
$alllicensedusers = Get-MsolUser -All | Where-Object {$_.isLicensed -eq $true}
Write-Host "Retrieving all groups and filtering based on if they apply licenses - this may take a while."
$allLicensedGroups = Get-MsolGroup -All | Where-Object {$_.licenses -ne $null}
# Loop through all licence types found in the tenant 
foreach ($license in $licensetype) {    
    # Build and write the Header for the CSV file 
    $headerstring = "DisplayName`tUserPrincipalName`tAccountEnabled`tAccountSku`tDirectAssigned`tGroupsAssigning" 
    foreach ($row in $($license.ServiceStatus)) {
		# Build header string
		if ($NoNameTranslation) {
			$thisLicence = [string]$row.ServicePlan.servicename
		}
		else {
			$thisLicence = componentlicenseswitch([string]($row.ServicePlan.servicename))
		}
        $headerstring = ($headerstring + "`t" + $thisLicence) 
    } 
    Write-Host ("Gathering users with the following subscription: " + $license.accountskuid) 
    # Gather users for this particular AccountSku from pre-existing array of users
    $users = $alllicensedusers | Where-Object {$_.licenses.accountskuid -contains $license.accountskuid} 
	if ($NoNameTranslation) {
		$RootLicence = ($($license.SkuPartNumber))
	}
	else {
		$RootLicence = RootLicenceswitch($($license.SkuPartNumber))
	}
	#$logfile = $CompanyName + "-" +$RootLicence + ".csv"
	$logfile = $CSVpath + "\" +$RootLicence + ".csv"
	Out-File -FilePath $LogFile -InputObject $headerstring -Encoding UTF8 -append
    # Loop through all users and write them to the CSV file 
    foreach ($user in $users) {
        Write-Verbose ("Processing " + $user.displayname) 
		$thislicense = $user.licenses | Where-Object {$_.accountskuid -eq $license.accountskuid} 
		if ($user.BlockCredential -eq $true) {
			$enabled = $false
		} else { 
			$enabled = $true
		}
		$datastring = ($user.displayname + "`t" + $user.userprincipalname + "`t" + $enabled + "`t" + $rootLicence)
		if ($thislicense.GroupsAssigningLicense.Count -eq 0) {
			$datastring = $datastring + "`t" + $true + "`t" + $false
		}
		else {
			if ($thislicense.GroupsAssigningLicense -contains $user.ObjectID) {
				$groups = $thislicense.groupsassigninglicense.guid | Where-Object {$_ -notlike $user.objectid}
				if ($null -eq $groups) {
					$groups = $false
				} else {
					$groups = (Invoke-GroupGuidConversion -GroupGuid $groups -LicenseGroups $allLicensedGroups).DisplayName -Join ";"
				}
				$datastring = $datastring + "`t" + $true + "`t" + $groups
			} else {
				if ($null -eq $groups) {
					$groups = $false
				} else {
				$groups = (Invoke-GroupGuidConversion -GroupGuid $groups -LicenseGroups $allLicensedGroups).DisplayName -Join ";"
				}
				$datastring = $datastring + "`t" + $false + "`t" + $groups
			}
		}
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
		$Selection = $worksheet.Range($worksheet.Cells(2,5), $worksheet.Cells($rows,6))
		[void]$Selection.Cells.Replace(";","`n",[Microsoft.Office.Interop.Excel.XlLookAt]::xlPart)
		$Selection = $worksheet.Range($worksheet.Cells(1,1), $worksheet.Cells($rows,$columns))
		$Selection.Font.Name = "Segoe UI"
		$Selection.Font.Size = 9
		if ($Worksheet.Name -ne "AllLicences") {
			Write-Host "Setting Conditional Formatting on "$Worksheet.Name
			$Selection= $worksheet.Range($worksheet.Cells(2,6), $worksheet.Cells($rows,$columns))
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
		else {
			foreach ($Item in (Import-Csv $CSVPath\AllLicences.csv -Delimiter "`t")) {	
				if ($NoNameTranslation) {
					$SearchString = $Item.'AccountLicenseSKU'
					$Selection = $worksheet.Range("A2").EntireColumn
					$Search = $Selection.find($SearchString,[Type]::Missing,[Type]::Missing,1) 
					$ResultCell = "A$($Search.Row)"
					$worksheet.Hyperlinks.Add($worksheet.Range($ResultCell),"","`'$($SearchString)`'!A1","$($SearchString)",$worksheet.Range($ResultCell).text)
				}
				else {
					$SearchString = $Item.'AccountLicenseSKU(Friendly)'
					$Selection = $worksheet.Range("A2").EntireColumn
					$Search = $Selection.find($SearchString,[Type]::Missing,[Type]::Missing,1) 
					$ResultCell = "A$($Search.Row)"
					$worksheet.Hyperlinks.Add($worksheet.Range($ResultCell),"","`'$($SearchString)`'!A1","$($SearchString)",$worksheet.Range($ResultCell).text)
				}
			}
			$worksheet.Move($worksheets.Item(1))
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