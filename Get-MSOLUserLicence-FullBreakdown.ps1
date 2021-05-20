#Requires -Modules MSOnline
#Requires -Version 5

<#
    .SYNOPSIS
        Name: Get-MSOLUserLicence-FullBreakdown.ps1
        The purpose of this script is is to export licensing details to excel

    .DESCRIPTION
        This script will log in to Microsoft 365 and then create a license report by SKU, with each component level status for each user, where 1 or more is assigned. This then conditionally formats the output to colours and autofilter.

    .NOTES
        Version 1.48
        Updated: 20210520    V1.48    1 tab = 4 spaces
        Updated: 20210520    V1.47    Added more components, renamed some components and added more SKUs
        Updated: 20210514    V1.46    Added more components, renamed some components and updated/added more SKUs
        Updated: 20210506    V1.45    Formatted Script to remove whitespace etc.
        Updated: 20210506    V1.44    Added Windows Update for Business Deployment Service component
        Updated: 20210323    V1.43    Added more Components
        Updated: 20210323    V1.42    Added more SKUs (F3, Conf PPM, E5 without Conf)
        Updated: 20210302    V1.41    Fixed missing New-Object's
        Updated: 20210223    V1.40    performance improvements for Group Based Licensing - no longer gets all groups; only gets the group once the GUID is found as an assigning group
        Updated: 20210222    V1.39    Added some EDU Root Level SKUs
        Updated: 20210222    V1.38    Moved Autofit and Autofilter to fix autofit on GBL column
        Updated: 20210208    V1.37    No longer out-files for everyline and performance improved
        Updated: 20201216    V1.36    Added components for Power Automate User with RPA Plan
        Updated: 20201216    V1.35    Added more SKUs (Multi-Geo, Communications Credits, M365 F1, Power Automate User with RPA Plan & Dynamics 365 Remote Assist)
        Updated: 20201028    V1.34    Added additional licence components (E5 Suite, PowerApps per IW, Win10 VDAE5)
        Updated: 20201021    V1.33    Resolved GBL issues
        Updated: 20201013    V1.32    Redid group based licensing to improve performance.
        Updated: 20201013    V1.31    Added User Enabled column
        Updated: 20200929    V1.30    Added RMS_Basic
        Updated: 20200929    V1.29    Added components for E5 Compliance
        Updated: 20200929    V1.28    Added code for group assigned and direct assigned licensing
        Updated: 20200820    V1.27    Added additional Office 365 E1 components
        Updated: 20200812    V1.26    Added Links to Licensing Sheets on All Licenses Page and move All Licenses Page to be first worksheet
        Updated: 20200730    V1.25    Added AIP P2 and Project for Office (E3 + E5)
        Updated: 20200720    V1.24    Added Virtual User component
        Updated: 20200718    V1.23    Added AAD Basic friendly component name
        Updated: 20200706    V1.22    Updated SKU error and added additional friendly names
        Updated: 20200626    V1.21    Updated F1 to F3 as per Microsoft's update
        Updated: 20200625    V1.20    Added Telephony Virtual User
        Updated: 20200603    V1.19    Added Switch for no name translation
        Updated: 20200603    V1.18    Added Telephony SKU's
        Updated: 20200501    V1.17    Script readability changes
        Updated: 20200430    V1.16    Made script more readable for Product type within component breakdown
        Updated: 20200422    V1.15    Formats to Segoe UI 9pt. Removed unnecessary True output.
        Updated: 20200408    V1.14    Added Teams Exploratory SKU
        Updated: 20200204    V1.13    Added more SKU's and Components
        Updated: 20191015    V1.12    Tidied up old comments
        Updated: 20190916    V1.11    Added more components and SKU's
        Updated: 20190830    V1.10    Added more components. Updated / renamed refreshed licences
        Updated: 20190627    V1.09    Added more Components
        Updated: 20190614    V1.08    Added more SKU's and Components
        Updated: 20190602    V1.07    Parameters, Comment based help, creates folder and deletes folder for csv's, require statements

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
        HelpMessage = 'Name of the Company you are running this against. This will form part of the output file name')
    ]
    [string]$CompanyName,
    [Parameter(
        Mandatory,
        HelpMessage = 'The location you would like the final excel file to reside'
    )][ValidateScript( {
            if (!(Test-Path -Path $_))
            {
                throw "The folder $_ does not exist"
            }
            else
            {
                return $true
            }
        })]
    [System.IO.DirectoryInfo]$OutputPath,
    [Parameter(
        HelpMessage = 'Credentials to connect to Office 365 if not already connected'
    )]
    [PSCredential]$Office365Credentials,
    [Parameter(
        HelpMessage = "This stops translation into Friendly Names of SKU's and Components"
    )][switch]$NoNameTranslation
)

#Enables Information Stream
$initialInformationPreference = $InformationPreference
$InformationPreference = 'Continue'

#Following Function Switches Complicated Component Names with Friendly Names
function componentlicenseswitch
{
    param
    (
        [parameter (Mandatory = $true, Position = 1)][string]$component
    )
    switch -wildcard ($($component))
    {
        #AAD
        'AAD_BASIC' { $thisLicence = 'Azure Acitve Directory Basic' }
        'AAD_PREMIUM' { $thisLicence = 'Azure Active Directory Premium P1' }
        'AAD_PREMIUM_P2' { $thisLicence = 'Azure Active Directory Premium P2' }
        'AAD_BASIC_EDU' { $thisLicence = 'Azure Active Directory Basic for EDU'}

        #Dynamics
        'DYN365_ENTERPRISE_SALES' { $thisLicence = 'Dynamics 365 for Sales' }
        'Dynamics_365_for_Talent_Team_members' { $thisLicence = 'Dynamics 365 for Talent Team members' }
        'Dynamics_365_for_Retail_Team_members' { $thisLicence = 'Dynamics 365 for Retail Team members' }
        'DYN365_Enterprise_Talent_Onboard_TeamMember' { $thisLicence = 'Dynamics 365 for Talent - Onboard Experience' }
        'DYN365_Enterprise_Talent_Attract_TeamMember' { $thisLicence = 'Dynamics 365 for Talent - Attract Experience Team Member' }
        'Dynamics_365_for_Operations_Team_members' { $thisLicence = 'Dynamics_365_for_Operations_Team_members' }
        'DYN365_ENTERPRISE_TEAM_MEMBERS' { $thisLicence = 'Dynamics 365 for Team Members' }
        'CCIBOTS_PRIVPREV_VIRAL'    { $thisLicence = 'Dynamics 365 AI for Customer Service Virtual Agents Viral' }
        'DYN365_CDS_CCI_BOTS' { $thisLicence = 'Common Data Service for CCI Bots' }
        'DYN365_AI_SERVICE_INSIGHS'    { $thisLicence = 'Dynamics 365 Customer Service Insights' }
        'POWERAPPS_DYN_TEAM' { $thisLicence = 'PowerApps for Dynamics 365' }
        'FLOW_DYN_TEAM' { $thisLicence = 'Flow for Dynamics 365' }
        'DYN365_TEAM_MEMBERS' { $thisLicence = 'Dynamics 365 Team Members' }
        'DYN365_BUSINESS_Marketing'    { $thisLicence = 'Dynamics 365 Marketing' }
        'DYN365_RETAIL_TRIAL' { $thisLicence = 'Dynamics 365 Retail Trial' }
        'Dynamics_365_for_Retail'    { $thisLicence = 'Dynamics 365 for Retail' }
        'DYN365_TALENT_ENTERPRISE'    { $thisLicence = 'Dynamics 365 for Talent' }
        'ERP_TRIAL_INSTANCE' { $thisLicence = 'AX7 Instance' }
        'Dynamics_365_Talent_Onboard' { $thisLicence = 'Dynamics 365 for Talent: Onboard' }
        'Dynamics_365_Onboarding_Free_PLAN' { $thisLicence = 'Dynamics 365 for Talent: Onboard' }
        'Dynamics_365_Hiring_Free_PLAN' { $thisLicence = 'Dynamics 365 for Talent: Attract' }
        'Dynamics_365_for_HCM_Trial' { $thisLicence = 'Dynamics_365_for_HCM_Trial' }
        'DYN365_ENTERPRISE_P1' { $thisLicence = 'Dynamics Enterprise P1' }
        'D365_CSI_EMBED_CE' { $thisLicence = 'Dynamics 365 Customer Service Insights for CE Plan' }
        'DYN365_ENTERPRISE_P1_IW' { $thisLicence = 'Dyn 365 P1 Trial Info Workers' }
        'MICROSOFT_REMOTE_ASSIST' { $thisLicence = 'Dynamics 365 Remote Assist' }
        'D365_ProjectOperations' { $thisLicence = 'Dynamics 365 Project Operations' }
        'D365_ProjectOperationsCDS'  { $thisLicence = 'Dynamics 365 Project Operations CDS' }
        'D365_CSI_EMBED_CSEnterprise' { $thisLicence = 'Dynamics 365 Customer Service Insights for CS Enterprise' }
        'DYN365_ENTERPRISE_CUSTOMER_SERVICE' { $thisLicence = 'Dynamics 365 for Customer Service'}

        #Dynamics Common Data Service
        'DYN365_CDS_O365_P1' { $thisLicence = 'Common Data Service' }
        'DYN365_CDS_O365_P2' { $thisLicence = 'Common Data Service' }
        'DYN365_CDS_O365_P3' { $thisLicence = 'Common Data Service' }
        'DYN365_CDS_O365_F1' { $thisLicence = 'Common Data Service' }
        'DYN365_CDS_P1' { $thisLicence = 'Common Data Service' }
        'DYN365_CDS_P2' { $thisLicence = 'Common Data Service' }
        'DYN365_CDS_FORMS_PRO' { $thisLicence = 'Common Data Service' }
        'DYN365_CDS_DYN_APPS' { $thisLicence = 'Common Data Service' }
        'DYN365_CDS_DYN_P2' { $thisLicence = 'Common Data Service' }
        'DYN365_CDS_VIRAL' { $thisLicence = 'Common Data Service' }
        'CDS_O365_P1' { $thisLicence = 'Common Data Service for Teams' }
        'CDS_O365_P2' { $thisLicence = 'Common Data Service for Teams' }
        'CDS_O365_P3' { $thisLicence = 'Common Data Service for Teams' }
        'CDS_O365_F1' { $thisLicence = 'Common Data Service for Teams' }
        'CDS_REMOTE_ASSIST'    { $thisLicence = 'Common Data Service for Remote Assist' }
        'CDS_ATTENDED_RPA'    { $thisLicence = 'Common Data Service Attended RPA' }

        #Exchange
        'EXCHANGE_S_ENTERPRISE' { $thisLicence = 'Exchange Online (Plan 2)' }
        'EXCHANGE_S_FOUNDATION' { $thisLicence = 'Core Exchange for non-Exch SKUs (e.g. setting profile pic)' }
        'EXCHANGE_S_DESKLESS' { $thisLicence = 'Exchange Online Firstline' }
        'EXCHANGE_S_STANDARD' { $thisLicence = 'Exchange Online (Plan 1)' }
        'EXCHANGE_S_ARCHIVE_ADDON'    { $thisLicence = 'Exchange Online Archiving Add-on' }
        'EXCHANGEONLINE_MULTIGEO'    { $thisLicence = 'Exchange Online Multi-Geo' }

        #Flow
        'FLOW_P1' { $thisLicence = 'Microsoft Flow Plan 1' }
        'FLOW_P2' { $thisLicence = 'Microsoft Flow Plan 2' }
        'FLOW_O365_P1' { $thisLicence = 'Flow for Office 365' }
        'FLOW_O365_P2' { $thisLicence = 'Flow for Office 365' }
        'FLOW_O365_P3' { $thisLicence = 'Flow for Office 365' }
        'FLOW_DYN_APPS' { $thisLicence = 'Flow for Dynamics 365' }
        'FLOW_P2_VIRAL' { $thisLicence = 'Flow Free' }
        'FLOW_P2_VIRAL_REAL' { $thisLicence = 'Flow P2 Viral' }
        'FLOW_CCI_BOTS' { $thisLicence = 'Flow for CCI Bots' }
        'Forms_Pro_CE' { $thisLicence = 'Forms Pro for Customer Engagement Plan' }
        'FORMS_PRO' { $thisLicence = 'Forms Pro' }
        'FLOW_FORMS_PRO'    { $thisLicence = 'Flow for Forms Pro' }
        'FLOW_O365_S1' { $thisLicence = 'Flow for Office 365 (F1)' }
        'FLOW_DYN_P2' { $thisLicence = 'Flow for Dynamics 365' }
        'POWER_AUTOMATE_ATTENDED_RPA'    { $thisLicence = 'Power Automate per user with attended RPA plan' }

        #Forms
        'FORMS_PLAN_E1' { $thisLicence = 'Microsoft Forms (Plan E1)' }
        'FORMS_PLAN_E3' { $thisLicence = 'Microsoft Forms (Plan E3)' }
        'FORMS_PLAN_E5' { $thisLicence = 'Microsoft Forms E5' }
        'Forms_Pro_Operations' { $thisLicence = 'Microsoft Forms Pro for Operations' }
        'Forms_Pro_CE' { $thisLicence = 'Forms Pro for Customer Engagement Plan' }
        'FORMS_PRO' { $thisLicence = 'Forms Pro' }
        'FORMS_PLAN_K' { $thisLicence = 'Microsoft Forms (Plan F1)' }
        'OFFICE_FORMS_PLAN_3' { $thisLicence = 'Microsoft Forms (Plan 3)' }
        'OFFICE_FORMS_PLAN_2' { $thisLicence = 'Microsoft Forms (Plan 2)' }

        #Kaizala
        'KAIZALA_STANDALONE' { $thisLicence = 'Microsoft Kaizala Pro' }
        'KAIZALA_O365_P1' { $thisLicence = 'Microsoft Kaizala Pro (P1)' }
        'KAIZALA_O365_P2' { $thisLicence = 'Microsoft Kaizala Pro' }
        'KAIZALA_O365_P3' { $thisLicence = 'Kaizala for Office 365' }

        #Misc Services
        'MYANALYTICS_P2' { $thisLicence = 'Insights by MyAnalytics' }
        'EXCHANGE_ANALYTICS' { $thisLicence = 'Microsoft MyAnalytics (Full)' }
        'Deskless' { $thisLicence = 'Microsoft StaffHub' }
        'SWAY' { $thisLicence = 'Sway' }
        'PROJECTWORKMANAGEMENT' { $thisLicence = 'Microsoft Planner' }
        'YAMMER_ENTERPRISE' { $thisLicence = 'Yammer Enterprise' }
        'SPZA' { $thisLicence = 'App Connect' }
        'MICROSOFT_BUSINESS_CENTER' { $thisLicence = 'Microsoft Business Center' }
        'NBENTERPRISE' { $thisLicence = 'Microsoft Social Engagement - Service Discontinuation' }
        'MICROSOFT_SEARCH' { $thisLicence = 'Microsoft Search' }
        'MICROSOFTBOOKINGS' { $thisLicence = 'Microsoft Bookings' }
        'EXCEL_PREMIUM' { $thisLicence = 'Microsoft Excel Advanced Analytics' }
        'GRAPH_CONNECTORS_SEARCH_INDEX' { $thisLicence = 'Graph Connectors Search with Index' }
        'UNIVERSAL_PRINT_01' { $thisLicence = 'Universal Print' }
        'SCHOOL_DATA_SYNC_P1' { $thisLicence = 'School Data Sync (Plan 1)' }
        'SCHOOL_DATA_SYNC_P2' { $thisLicence = 'School Data Sync (Plan 2)' }
        'MINECRAFT_EDUCATION_EDITION' { $thisLicence = 'Minecraft Education Edition' }
        'EducationAnalyticsP1' { $thisLicence = 'Education Analytics' }
        'YAMMER_EDU' { $thisLicence = 'Yammer for Academic' }

        #Office
        'SHAREPOINTWAC' { $thisLicence = 'Office Online' }
        'OFFICESUBSCRIPTION' { $thisLicence = 'M365 Apps for Enterprise' }
        'OFFICEMOBILE_SUBSCRIPTION' { $thisLicence = 'Office Mobile Apps for Office 365' }
        'SAFEDOCS' { $thisLicence = 'Office 365 SafeDocs' }
        'SHAREPOINTWAC_EDU' { $thisLicence = 'Office for the web (Education)' }
        'OFFICESUBSCRIPTION_unattended' { $thisLicence = 'M365 Apps for Enterprise (unattended)'}

        #OneDrive
        'ONEDRIVESTANDARD' { $thisLicence = 'OneDrive for Business (Plan 1)' }
        'ONEDRIVE_BASIC' { $thisLicence = 'OneDrive Basic' }

        #PowerBI
        'BI_AZURE_P0' { $thisLicence = 'Power BI (Free)' }
        'BI_AZURE_P2' { $thisLicence = 'Power BI Pro' }

        #Phone System
        'MCOEV' { $thisLicence = 'M365 Phone System' }
        'MCOMEETADV' { $thisLicence = 'M365 Audio Conferencing' }
        'MCOEV_VIRTUALUSER' { $thisLicence = 'Microsoft 365 Phone System Virtual User' }
        'MCOPSTNC' { $thisLicence = 'Communications Credits' }

        #PowerApps
        'POWERAPPS_O365_S1' { $thisLicence = 'PowerApps for Office 365 Firstline' }
        'POWERAPPS_O365_P1' { $thisLicence = 'PowerApps for Office 365' }
        'POWERAPPS_O365_P2' { $thisLicence = 'PowerApps for Office 365' }
        'POWERAPPS_O365_P3' { $thisLicence = 'PowerApps for Office 365' }
        'POWERAPPS_DYN_APPS' { $thisLicence = 'PowerApps for Dynamics 365' }
        'POWERAPPS_P2_VIRAL' { $thisLicence = 'PowerApps Plan 2 Trial' }
        'POWERAPPS_P2' { $thisLicence = 'PowerApps Plan 2' }
        'POWERAPPS_DYN_P2' { $thisLicence = 'PowerApps for Dynamics 365' }
        'POWER_VIRTUAL_AGENTS_O365_P1'    { $thisLicence = 'Power Virtual Agents for Office 365' }
        'POWER_VIRTUAL_AGENTS_O365_P2'    { $thisLicence = 'Power Virtual Agents for Office 365' }
        'POWER_VIRTUAL_AGENTS_O365_P3'    { $thisLicence = 'Power Virtual Agents for Office 365' }
        'POWER_VIRTUAL_AGENTS_O365_F1'    { $thisLicence = 'Power Virtual Agents for Office 365' }
        'POWERAPPS_PER_APP_IWTRIAL' { $thisLicence = 'Power Apps per app baseline access' }
        'Flow_Per_APP_IWTRIAL' { $thisLicence = 'Flow per app baseline access' }
        'CDS_PER_APP_IWTRIAL' { $thisLicence = 'CDS per app baseline access' }

        #Project
        'PROJECT_PROFESSIONAL' { $thisLicence = 'Project P3' }
        'FLOW_FOR_PROJECT' { $thisLicence = 'Data Integration for Project with Flow' }
        'DYN365_CDS_PROJECT' { $thisLicence = 'Common Data Service for Project' }
        'SHAREPOINT_PROJECT' { $thisLicence = 'Project Online Service' }
        'PROJECT_CLIENT_SUBSCRIPTION'    { $thisLicence = 'Project Online Desktop Client' }
        'PROJECT_ESSENTIALS' { $thisLicence = 'Project Online Essentials' }
        'PROJECT_O365_P1' { $thisLicence = 'Project for Office (Plan E1)' }
        'PROJECT_O365_P2' { $thisLicence = 'Project for Office (Plan E3)' }
        'PROJECT_O365_P3' { $thisLicence = 'Project for Office (Plan E5)' }
        'PROJECT_O365_F3' { $thisLicence = 'Project for Office (Plan F3)' }
        'PROJECT_FOR_PROJECT_OPERATIONS' { $thisLicence = 'Project for Project Operations' }

        #Security & Compliance
        'RECORDS_MANAGEMENT' { $thisLicence = 'Microsoft Records Management' }
        'INFO_GOVERNANCE' { $thisLicence = 'Microsoft Information Governance' }
        'DATA_INVESTIGATIONS' { $thisLicence = 'Microsoft Data Investigations' }
        'CUSTOMER_KEY' { $thisLicence = 'Microsoft Customer Key' }
        'COMMUNICATIONS_DLP' { $thisLicence = 'Microsoft Communications DLP' }
        'COMMUNICATIONS_COMPLIANCE'    { $thisLicence = 'Microsoft Communications Compliance' }
        'M365_ADVANCED_AUDITING' { $thisLicence = 'Microsoft 365 Advanced Auditing' }
        'ATP_ENTERPRISE' { $thisLicence = 'O365 ATP Plan 1 (not licenced individually)' }
        'THREAT_INTELLIGENCE' { $thisLicence = 'O365 ATP Plan 2' }
        'ADALLOM_S_O365' { $thisLicence = 'Office 365 Cloud App Security' }
        'EQUIVIO_ANALYTICS' { $thisLicence = 'Office 365 Advanced eDiscovery' }
        'LOCKBOX_ENTERPRISE' { $thisLicence = 'Customer Lockbox' }
        'ATA' { $thisLicence = 'Azure Advanced Threat Protection' }
        'ADALLOM_S_STANDALONE' { $thisLicence = 'Microsoft Cloud App Security' }
        'RMS_S_BASIC' { $thisLicence = 'Azure Rights Management Service (non-assignable)' }
        'RMS_S_ENTERPRISE' { $thisLicence = 'Azure Rights Management' }
        'RMS_S_PREMIUM' { $thisLicence = 'Azure Information Protection Premium P1' }
        'RMS_S_PREMIUM2' { $thisLicence = 'Azure Information Protection Premium P2' }
        'RMS_S_ADHOC' { $thisLicence = 'Rights Management Adhoc' }
        'INTUNE_A' { $thisLicence = 'Microsoft Intune' }
        'INTUNE_A_VL' { $thisLicence = 'Microsoft Intune' }
        'MFA_PREMIUM' { $thisLicence = 'Microsoft Azure Multi-Factor Authentication' }
        'INTUNE_O365' { $thisLicence = 'MDM for Office 365 (not licenced individually)' }
        'PAM_ENTERPRISE' { $thisLicence = 'O365 Priviledged Access Management' }
        'ADALLOM_S_DISCOVERY' { $thisLicence = 'Cloud App Security Discovery' }
        'MIP_S_CLP1' { $thisLicence = 'Information Protection for Office 365 - Standard' }
        'MIP_S_CLP2' { $thisLicence = 'Information Protection for Office 365 - Premium' }
        'MIP_S_EXCHANGE' { $thisLicence = 'Data Classification in Microsoft 365' }
        'PREMIUM_ENCRYPTION' { $thisLicence = 'Premium Encryption' }
        'INFORMATION_BARRIERS' { $thisLicence = 'Information Barriers' }
        'WINDEFATP' { $thisLicence = 'Windows Defender ATP' }
        'MTP' { $thisLicence = 'Microsoft Threat Protection' }
        'Content_Explorer' { $thisLicence = 'Content Explorer (Assigned at Org Level)' }
        'MICROSOFTENDPOINTDLP' { $thisLicence = 'Microsoft Endpoint DLP' }
        'INSIDER_RISK' { $thisLicence = 'Microsoft Insider Risk Management' }
        'INSIDER_RISK_MANAGEMENT' { $thisLicence = 'RETIRED - Microsoft Insider Risk Management' }
        'ML_CLASSIFICATION' { $thisLicence = 'Microsoft ML_based Classification' }
        'MICROSOFT_COMMUNICATION_COMPLIANCE' { $thisLicence = 'RETIRED - Microsoft Communications Compliance' }
        'INTUNE_EDU' { $thisLicence = 'Intune for Education' }

        #SharePoint
        'SHAREPOINTDESKLESS' { $thisLicence = 'SharePoint Online Kiosk' }
        'SHAREPOINTSTANDARD' { $thisLicence = 'SharePoint Online (Plan 1)' }
        'SHAREPOINTENTERPRISE' { $thisLicence = 'SharePoint Online (Plan 2)' }
        'SHAREPOINTONLINE_MULTIGEO'    { $thisLicence = 'Sharepoint Multi-Geo' }
        'SHAREPOINTSTANDARD_EDU'  { $thisLicence = 'SharePoint Plan 1 for EDU' }
        'SHAREPOINTENTERPRISE_EDU' { $thisLicence = 'SharePoint Plan 2 for EDU' }

        #Skype
        'MCOIMP' { $thisLicence = 'Skype for Business (Plan 1)' }
        'MCOSTANDARD' { $thisLicence = 'Skype for Business Online (Plan 2)' }

        #Stream
        'STREAM_O365_K' { $thisLicence = 'Stream for Office 365 Firstline' }
        'STREAM_O365_E1' { $thisLicence = 'Microsoft Stream for O365 E1' }
        'STREAM_O365_E3' { $thisLicence = 'Microsoft Stream for O365 E3 SKU' }
        'STREAM_O365_E5' { $thisLicence = 'Stream E5' }

        #Teams
        'TEAMS1' { $thisLicence = 'Microsoft Teams' }
        'MCO_TEAMS_IW' { $thisLicence = 'Microsoft Teams Trial' }
        'TEAMS_FREE_SERVICE' { $thisLicence = 'Teams Free Service (Not assigned per user)' }
        'MCOFREE' { $thisLicence = 'MCO Free for Microsoft Teams (free)' }
        'TEAMS_FREE' { $thisLicence = 'Microsoft Teams (free)' }

        #Telephony
        'MCOPSTN1' { $thisLicence = 'Domestic Calling Plan (1200 min)' }
        'MCOPSTN2' { $thisLicence = 'Domestic and International Calling Plan' }
        'MCOPSTN5' { $thisLicence = 'Domestic Calling Plan (120 min)' }
        'PHONESYSTEM_VIRTUALUSER'    { $thisLicence = 'M365 Phone System - Virtual User' }
        'MCOMEETACPEA' { $thisLicence = 'M365 Audio Conf Pay-Per-Minute' }

        #To-Do
        'BPOS_S_TODO_FIRSTLINE' { $thisLicence = 'To-Do Firstline' }
        'BPOS_S_TODO_1' { $thisLicence = 'To-Do Plan 1' }
        'BPOS_S_TODO_2' { $thisLicence = 'To-Do (Plan 2)' }
        'BPOS_S_TODO_3' { $thisLicence = 'To-Do (Plan 3)' }

        #Visio
        'VISIOONLINE' { $thisLicence = 'Visio Online' }
        'VISIO_CLIENT_SUBSCRIPTION' { $thisLicence = 'Visio Pro for Office 365' }

        #Whiteboard
        'WHITEBOARD_FIRSTLINE1' { $thisLicence = 'Whiteboard for Firstline' }
        'WHITEBOARD_PLAN1' { $thisLicence = 'Whiteboard Plan 1' }
        'WHITEBOARD_PLAN2' { $thisLicence = 'Whiteboard Plan 2' }
        'WHITEBOARD_PLAN3' { $thisLicence = 'Whiteboard Plan 3' }

        #Windows 10
        'WIN10_PRO_ENT_SUB' { $thisLicence = 'Win 10 Enterprise E3' }
        'WIN10_ENT_LOC_F1' { $thisLicence = 'Win 10 Enterprise E3 (Local Only)' }
        'WINDOWSUPDATEFORBUSINESS_DEPLOYMENTSERVICE' { $thisLicence = 'Windows Update for Business Deployment Service' }
        default { $thisLicence = $component }
    }
    Write-Output $thisLicence
}
#Following Function Switches Complicated Top Level SKU Names with Friendly Names
function RootLicenceswitch
{
    param (
        [parameter (Mandatory = $true, Position = 1)][string]$licensesku
    )
    switch -wildcard ($($licensesku))
    {
        #Azure AD
        'AAD_BASIC' { $RootLicence = 'Azure Active Directory Basic' }
        'AAD_PREMIUM' { $RootLicence = 'Azure Active Directory Premium' }
        'AAD_PREMIUM_P2' { $RootLicence = 'Azure AD Premium P2' }

        #Dynamics
        'DYN365_ENTERPRISE_PLAN1' { $RootLicence = 'Dyn 365 Customer Engage Ent Ed' }
        'DYN365_ENTERPRISE_CUSTOMER_SERVICE' { $RootLicence = 'Dyn 365 Customer Service' }
        'PROJECT_MADEIRA_PREVIEW_IW_SKU' { $RootLicence = 'Dynamics 365 for Financials for IWs' }
        'DYN365_AI_SERVICE_INSIGHTS' { $RootLicence = 'Dyn 365 CSI Trial' }
        'Dynamics_365_for_Operations' { $RootLicence = 'Dyn 365 Unified Operations Plan' }
        'Dynamics_365_Onboarding_SKU' { $RootLicence = 'Dyn 365 for Talent Onboard' }
        'CCIBOTS_PRIVPREV_VIRAL' { $RootLicence = 'Dyn 365 AI for CSVAV' }
        'DYN365_BUSINESS_MARKETING' { $RootLicence = 'Dyn 365 Marketing' }
        'DYN365_RETAIL_TRIAL' { $RootLicence = 'Dyn 365 Retail Trial' }
        'SKU_Dynamics_365_for_HCM_Trial' { $RootLicence = 'Dyn 365 Talent' }
        'DYN365_FINANCIALS_BUSINESS_SKU' { $RootLicence = 'Dyn 365 Financials Business Edition' }
        'DYN365_FINANCIALS_TEAM_MEMBERS_SKU' { $RootLicence = 'Dyn 365 Team Members Business Edition' }
        'AX7_USER_TRIAL' { $RootLicence = 'Dynamics AX7 Trial' }
        'DYN365_ENTERPRISE_P1_IW' { $RootLicence = 'Dyn 365 P1 Trial Info Workers' }
        'DYN365_ENTERPRISE_TEAM_MEMBERS' { $RootLicence = 'Dyn 365 Team Members Ent Ed' }
        'DYN365_TEAM_MEMBERS' { $RootLicence = 'Dynamics 365 Team Members' }
        'CRMSTORAGE' { $RootLicence = 'Microsoft Dynamics CRM Online Additional Storage' }
        'CRMSTANDARD' { $RootLicence = 'Microsoft Dynamics CRM Online Professional' }
        'DYN365_ENTERPRISE_SALES' { $RootLicence = 'Dyn 365 Enterprise Sales' }
        'MICROSOFT_REMOTE_ASSIST' { $RootLicence = 'Dyn 365 Remote Assist' }
        'MICROSOFT_REMOTE_ASSIST_ATTACH' { $RootLicence = 'Dyn 365 Remote Assist Attach' }

        #Exchange
        'EXCHANGESTANDARD_GOV' { $RootLicence = 'Exchange Online (P1) for Government' }
        'EXCHANGEENTERPRISE_GOV' { $RootLicence = 'Exchange Online (P2) for Government' }
        'EXCHANGE_S_DESKLESS_GOV' { $RootLicence = 'Exchange Kiosk' }
        'ECAL_SERVICES' { $RootLicence = 'ECAL' }
        'EXCHANGE_S_ENTERPRISE_GOV' { $RootLicence = 'Exchange Plan 2G' }
        'EXCHANGE_S_ARCHIVE_ADDON_GOV' { $RootLicence = 'Exchange Online Archiving' }
        'EXCHANGE_S_DESKLESS' { $RootLicence = 'Exchange Online Kiosk' }
        'EXCHANGE_L_STANDARD' { $RootLicence = 'Exchange Online (Plan 1)' }
        'EXCHANGE_S_STANDARD_MIDMARKET' { $RootLicence = 'Exchange Online (Plan 1)' }
        'EXCHANGESTANDARD' { $RootLicence = 'Exchange Online (Plan 1)' }
        'EXCHANGEENTERPRISE' { $RootLicence = 'Exchange Online Plan 2' }
        'EXCHANGEENTERPRISE_FACULTY' { $RootLicence = 'Exchange Online P2 Faculty' }
        'EOP_ENTERPRISE_FACULTY' { $RootLicence = 'Exchange Online Protection for Faculty' }
        'EXCHANGESTANDARD_STUDENT' { $RootLicence = 'Exchange Online (P1) Students' }
        'EXCHANGEARCHIVE_ADDON' { $RootLicence = 'O-Archive for Exchange Online' }
        'EXCHANGEDESKLESS' { $RootLicence = 'Exchange Online Kiosk' }

        #Flow
        'FLOW_FREE' { $RootLicence = 'Microsoft Flow Free' }
        'FLOW_P1' { $RootLicence = 'Microsoft Flow Plan 1' }
        'FLOW_P2' { $RootLicence = 'Microsoft Flow Plan 2' }
        'POWER_AUTOMATE_ATTENDED_RPA' { $RootLicence = 'Power Automate User with RPA' }

        #Forms
        'FORMS_PRO' { $RootLicence = 'Forms Pro Trial' }

        #Microsoft 365 Subscription
        'STANDARDPACK_GOV' { $RootLicence = 'O365 (Plan G1) Government' }
        'STANDARDWOFFPACK_GOV' { $RootLicence = 'O365 (Plan G2) Government' }
        'ENTERPRISEPACK_GOV' { $RootLicence = 'O365 (Plan G3) Government' }
        'ENTERPRISEWITHSCAL_GOV' { $RootLicence = 'O365 (Plan G4) Government' }
        'DESKLESSPACK_GOV' { $RootLicence = 'O365 (Plan K1) Government' }
        'DESKLESSWOFFPACK_GOV' { $RootLicence = 'O365 (Plan K2) Government' }
        'SPE_E3' { $RootLicence = 'M365 E3' }
        'SPE_E5' { $RootLicence = 'M365 E5' }
        'SPE_E5_NOPSTNCONF' { $RootLicence = 'M365 E5 without Audio Conf' }
        'SPE_F1' { $RootLicence = 'M365 F3' }
        'STANDARDWOFFPACK_STUDENT' { $RootLicence = 'O365 A1 for Students' }
        'M365_F1_COMM' { $RootLicence = 'Microsoft 365 F1' }
        'M365_E5_SUITE_COMPONENTS' { $RootLicence = 'M365 E5 Suite Features' }
        'M365EDU_A5_FACULTY' { $RootLicence = 'M365 Education A5 Faculty' }
        'M365EDU_A5_STUUSEBNFT' { $RootLicence = 'M365 EDU A5 Student Bens' }
        'SPE_E3_RPA1' { $RootLicence = 'M365 E3 Unattended' }

        #Misc Services
        'PLANNERSTANDALONE' { $RootLicence = 'Planner Standalone' }
        'CRMIUR' { $RootLicence = 'CMRIUR' }
        'PROJECTWORKMANAGEMENT' { $RootLicence = 'Office 365 Planner Preview' }
        'STREAM' { $RootLicence = 'Microsoft Stream Trial' }
        'SPZA_IW' { $RootLicence = 'App Connect' }
        'IT_ACADEMY_AD' { $RootLicence = 'MS Imagine Academy' }
        'EXTRA_CONNECTOR_CAPACITY' { $RootLicence = 'Extra Graph Connector' }

        #Office 365 Subscription
        'O365_BUSINESS' { $RootLicence = 'O365 Business' }
        'O365_BUSINESS_ESSENTIALS' { $RootLicence = 'O365 Business Essentials' }
        'O365_BUSINESS_PREMIUM' { $RootLicence = 'O365 Business Premium' }
        'DESKLESSPACK' { $RootLicence = 'O365 F3' }
        'DESKLESSWOFFPACK' { $RootLicence = 'O365 K2' }
        'LITEPACK' { $RootLicence = 'O365 P1' }
        'STANDARDPACK' { $RootLicence = 'O365 E1' }
        'STANDARDWOFFPACK' { $RootLicence = 'O365 E2' }
        'ENTERPRISEPACK' { $RootLicence = 'O365 E3' }
        'ENTERPRISEPACKLRG' { $RootLicence = 'O365 E3' }
        'ENTERPRISEWITHSCAL' { $RootLicence = 'O365 E4' }
        'ENTERPRISEPREMIUM_NOPSTNCONF' { $RootLicence = 'O365 E5 (without Audio Conferencing' }
        'ENTERPRISEPREMIUM' { $RootLicence = 'O365 E5' }
        'STANDARDPACK_STUDENT' { $RootLicence = 'O365 A1 Students' }
        'STANDARDWOFFPACKPACK_STUDENT' { $RootLicence = 'O365 A2 Students' }
        'ENTERPRISEPACK_STUDENT' { $RootLicence = 'O365 A3 Students' }
        'ENTERPRISEWITHSCAL_STUDENT' { $RootLicence = 'O365 A4 Students' }
        'STANDARDPACK_FACULTY' { $RootLicence = 'O365 A1 Faculty' }
        'STANDARDWOFFPACKPACK_FACULTY' { $RootLicence = 'O365 A2 Faculty' }
        'ENTERPRISEPACK_FACULTY' { $RootLicence = 'O365 A3 Faculty' }
        'ENTERPRISEWITHSCAL_FACULTY' { $RootLicence = 'O365 A4 Faculty' }
        'ENTERPRISEPACK_B_PILOT' { $RootLicence = 'O365 (Enterprise Preview)' }
        'STANDARD_B_PILOT' { $RootLicence = 'O365 (Small Business Preview)' }
        'STANDARDWOFFPACK_IW_STUDENT' { $RootLicence = 'O365 Education for Students' }
        'STANDARDWOFFPACK_IW_FACULTY' { $RootLicence = 'O365 Education for Faculty' }
        'STANDARDWOFFPACK_FACULTY' { $RootLicence = 'O365 A1 for Faculty' }
        'ENTERPRISEPACKWITHOUTPROPLUS' { $RootLicence = 'O365 E3 No Pro Plus' }
        'OFFICE365_MULTIGEO' { $RootLicence = 'Multi-Geo in Office 365' }

        #Office Suite
        'OFFICESUBSCRIPTION_GOV' { $RootLicence = 'Office ProPlus' }
        'SHAREPOINTWAC_GOV' { $RootLicence = 'Office Online for Government' }
        'SHAREPOINTWAC' { $RootLicence = 'Office Online' }
        'OFFICE_PRO_PLUS_SUBSCRIPTION_SMBIZ' { $RootLicence = 'Office ProPlus' }
        'OFFICESUBSCRIPTION' { $RootLicence = 'Office ProPlus' }
        'OFFICESUBSCRIPTION_FACULTY' { $RootLicence = 'Office ProPlus Faculty' }
        'OFFICESUBSCRIPTION_STUDENT' { $RootLicence = 'Office ProPlus Student Benefit' }

        #PowerApps
        'POWERFLOW_P1' { $RootLicence = 'PowerApps Plan 1' }
        'POWERFLOW_P2' { $RootLicence = 'PowerApps Plan 2' }
        'POWERFLOW_P2_FACULTY' { $RootLicence = 'PowerApps Plan 2 Faculty' }
        'POWERAPPS_INDIVIDUAL_USER' { $RootLicence = 'PowerApps and Logic Flows' }
        'POWERAPPS_VIRAL' { $RootLicence = 'PowerApps Plan 2 Trial' }
        'POWERAPPS_PER_APP_IWL' { $RootLicence = 'PowerApps per app Baselinel' }

        #Power BI
        'POWER_BI_ADDON' { $RootLicence = 'Office 365 Power BI Addon' }
        'POWER_BI_INDIVIDUAL_USE' { $RootLicence = 'Power BI Individual User' }
        'POWER_BI_STANDALONE' { $RootLicence = 'Power BI Stand Alone' }
        'POWER_BI_STANDARD' { $RootLicence = 'Power BI Free' }
        'BI_AZURE_P1' { $RootLicence = 'Power BI Reporting and Analytics' }
        'POWER_BI_PRO' { $RootLicence = 'Power BI Pro' }
        'POWER_BI_PRO_CE' { $RootLicence = 'Power BI Pro' }
        'POWER_BI_PRO_FACULTY' { $RootLicence = 'Power BI Pro Faculty' }
        'POWER_BI_PRO_DEPT' { $RootLicence = 'Power BI Pro DEPT' }

        #PhoneSystem
        'MCOPSTNC' { $RootLicence = 'Communications Credits' }

        #Project
        'PROJECTESSENTIALS' { $RootLicence = 'Project Lite' }
        'PROJECTCLIENT' { $RootLicence = 'Project Professional' }
        'PROJECTONLINE_PLAN_1' { $RootLicence = 'Project Online' }
        'PROJECTONLINE_PLAN_2' { $RootLicence = 'Project Online and PRO' }
        'ProjectPremium' { $RootLicence = 'Project Online Premium' }
        'PROJECTPROFESSIONAL' { $RootLicence = 'Project Professional' }

        #Security and Compliance
        'EMS' { $RootLicence = 'EMS (Plan E3)' }
        'EMSPREMIUM' { $RootLicence = 'EMS (Plan E5)' }
        'RIGHTSMANAGEMENT_ADHOC' { $RootLicence = 'Windows Azure RMS' }
        'INTUNE_A' { $RootLicence = 'Microsoft Intune' }
        'INTUNE_A_VL' { $RootLicence = 'Microsoft Intune' }
        'ATP_ENTERPRISE' { $RootLicence = 'Office 365 ATP Plan 1' }
        'THREAT_INTELLIGENCE' { $RootLicence = 'Office 365 ATP Plan 2' }
        'EQUIVIO_ANALYTICS' { $RootLicence = 'Office 365 Advanced Compliance' }
        'RMS_S_ENTERPRISE' { $RootLicence = 'Azure Active Directory Rights Management' }
        'MFA_PREMIUM' { $RootLicence = 'Azure Multi-Factor Authentication' }
        'RMS_S_ENTERPRISE_GOV' { $RootLicence = 'Windows Azure AD RMS' }
        'IDENTITY_THREAT_PROTECTION' { $RootLicence = 'Microsoft 365 E5 Security' }
        'INFORMATION_PROTECTION_COMPLIANCE' { $RootLicence = 'Microsoft 365 E5 Compliance' }
        'EMS_FACULTY' { $RootLicence = 'EMS (Plan E3) Faculty' }
        'ADALLOM_STANDALONE' { $RootLicence = 'Microsoft Cloud App Security' }
        'ATA' { $RootLicence = 'Advanced Threat Analytics' }
        'WIN_DEF_ATP' { $RootLicence = 'Windows 10 Defender ATP' }
        'RIGHTSMANAGEMENT' { $RootLicence = 'Rights Management' }
        'INFOPROTECTION_P2' { $RootLicence = 'AIP Premium P2' }
        'COMMUNICATIONS_COMPLIANCE' { $RootLicence = 'Communications Compliance' }

        #Skype
        'MCOSTANDARD_GOV' { $RootLicence = 'Lync Plan 2G' }
        'MCOLITE' { $RootLicence = 'Lync Online (Plan 1)' }
        'MCOSTANDARD_MIDMARKET' { $RootLicence = 'Lync Online (Plan 1)' }
        'MCOSTANDARD' { $RootLicence = 'SFBO Plan 2' }
        'VIDEO_INTEROP' { $RootLicence = 'Polycom Video Interop' }

        #SharePoint
        'SHAREPOINTSTORAGE' { $RootLicence = 'SharePoint storage' }
        'SHAREPOINTDESKLESS_GOV' { $RootLicence = 'SharePoint Online Kiosk' }
        'SHAREPOINTENTERPRISE_GOV' { $RootLicence = 'SharePoint Plan 2G' }
        'SHAREPOINTDESKLESS' { $RootLicence = 'SharePoint Online Kiosk' }
        'SHAREPOINTLITE' { $RootLicence = 'SharePoint Online (Plan 1)' }
        'SHAREPOINTENTERPRISE_MIDMARKET'    { $RootLicence = 'SharePoint Online (Plan 1)' }

        #Teams
        'TEAMS_FREE' { $RootLicence = 'Microsoft Teams (Free)' }
        'TEAMS_EXPLORATORY' { $RootLicence = 'Microsoft Teams Exploratory' }
        'TEAMS_COMMERCIAL_TRIAL' { $RootLicence = 'Teams Commercial Cloud' }
        'MS_TEAMS_IW' { $RootLicence = 'Microsoft Teams Trial' }
        'MEETING_ROOM' { $RootLicence = 'Meeting Room' }
        'MCOCAP' { $RootLicence = 'Common Area Phone' }

        #Telephony
        'MCOMEETADV' { $RootLicence = 'PSTN conferencing' }
        'MCOPSTN1' { $RootLicence = 'Dom Calling Plan (1200 mins)' }
        'MCOPSTN1_FACULTY' { $RootLicence = 'Dom Calling Plan (1200 mins) Faculty' }
        'MCOPSTN2' { $RootLicence = 'Dom and Intl Calling Plan' }
        'MCOEV' { $RootLicence = 'Microsoft Phone System' }
        'MCOPSTN_5' { $RootLicence = 'Dom Calling Plan (120mins)' }
        'PHONESYSTEM_VIRTUALUSER' { $RootLicence = 'Phone System Virtual User' }
        'PHONESYSTEM_VIRTUALUSER_FACULTY'    { $RootLicence = 'Phone System V User Faculty' }
        'MCOMEETACPEA' { $RootLicence = 'M365 Audio Conf PPM' }

        #Visio
        'VISIOCLIENT' { $RootLicence = 'Visio Pro Online' }
        'VISIOONLINE_PLAN1' { $RootLicence = 'Visio Online Plan 1' }

        #Windows 10
        'Win10_VDA_E3' { $RootLicence = 'Windows 10 E3' }
        'Win10_VDA_E5' { $RootLicence = 'Windows 10 E5' }
        'WINDOWS_STORE' { $RootLicence = 'Windows Store for Business' }
        'SMB_APPS' { $RootLicence = 'Microsoft Business Apps' }
        'MICROSOFT_BUSINESS_CENTER' { $RootLicence = 'Microsoft Business Center' }

        #Yammer
        'YAMMER_ENTERPRISE' { $RootLicence = 'Yammer Enterprise' }
        'YAMMER_MIDSIZE' { $RootLicence = 'Yammer' }
        default { $RootLicence = $licensesku }
    }
    Write-Output $RootLicence
}
#Helper function for tidier select of Groups for Group Based Licensing
Function Invoke-GroupGuidConversion
{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [String[]]
        $GroupGuid,
        [Parameter(Mandatory)]
        [hashtable]
        $LicenseGroups
    )
    $output = New-Object System.Collections.Generic.List[System.Object]
    foreach ($guid in $GroupGuid)
    {
        $temp = [PSCustomObject]@{
            DisplayName = $LicenseGroups[$guid]
        }
        $output.Add($temp)
        Remove-Variable temp
    }
    Write-Output $output
}
#Helper function for merging CSV files and conditional formatting etc.
Function Merge-CSVFile
{
    $csvFiles = Get-ChildItem ("$CSVPath\*") -Include *.csv
    $Excel = New-Object -ComObject excel.application
    $Excel.visible = $false
    $Excel.sheetsInNewWorkbook = $csvFiles.Count
    $workbooks = $excel.Workbooks.Add()
    $CSVSheet = 1
    Foreach ($CSV in $Csvfiles)
    {
        $worksheets = $workbooks.worksheets
        $CSVFullPath = $CSV.FullName
        $SheetName = ($CSV.name -split '\.')[0]
        $worksheet = $worksheets.Item($CSVSheet)
        $worksheet.Name = $SheetName
        $TxtConnector = ('TEXT;' + $CSVFullPath)
        $CellRef = $worksheet.Range('A1')
        $Connector = $worksheet.QueryTables.add($TxtConnector, $CellRef)
        $worksheet.QueryTables.item($Connector.name).TextFileOtherDelimiter = "`t"
        $worksheet.QueryTables.item($Connector.name).TextFileParseType = 1
        $worksheet.QueryTables.item($Connector.name).Refresh()
        $worksheet.QueryTables.item($Connector.name).delete()
        $CSVSheet++
    }
    $worksheets = $workbooks.worksheets
    $xlTextString = [Microsoft.Office.Interop.Excel.XlFormatConditionType]::xlTextString
    $xlContains = [Microsoft.Office.Interop.Excel.XlContainsOperator]::xlContains
    foreach ($worksheet in $worksheets)
    {
        Write-Information ('Freezing panes on ' + $Worksheet.Name)
        $worksheet.Select()
        $worksheet.application.activewindow.splitcolumn = 1
        $worksheet.application.activewindow.splitrow = 1
        $worksheet.application.activewindow.freezepanes = $true
        $rows = $worksheet.UsedRange.Rows.count
        $columns = $worksheet.UsedRange.Columns.count
        $Selection = $worksheet.Range($worksheet.Cells(2, 5), $worksheet.Cells($rows, 6))
        [void]$Selection.Cells.Replace(';', "`n", [Microsoft.Office.Interop.Excel.XlLookAt]::xlPart)
        $Selection = $worksheet.Range($worksheet.Cells(1, 1), $worksheet.Cells($rows, $columns))
        $Selection.Font.Name = 'Segoe UI'
        $Selection.Font.Size = 9
        if ($Worksheet.Name -ne 'AllLicences')
        {
            Write-Information ('Setting Conditional Formatting on ' + $Worksheet.Name)
            $Selection = $worksheet.Range($worksheet.Cells(2, 6), $worksheet.Cells($rows, $columns))
            $Selection.FormatConditions.Add($xlTextString, '', $xlContains, 'Success', 'Success', 0, 0) | Out-Null
            $Selection.FormatConditions.Item(1).Interior.ColorIndex = 35
            $Selection.FormatConditions.Item(1).Font.ColorIndex = 51
            $Selection.FormatConditions.Add($xlTextString, '', $xlContains, 'Pending', 'Pending', 0, 0) | Out-Null
            $Selection.FormatConditions.Item(2).Interior.ColorIndex = 36
            $Selection.FormatConditions.Item(2).Font.ColorIndex = 12
            $Selection.FormatConditions.Add($xlTextString, '', $xlContains, 'Disabled', 'Disabled', 0, 0) | Out-Null
            $Selection.FormatConditions.Item(3).Interior.ColorIndex = 38
            $Selection.FormatConditions.Item(3).Font.ColorIndex = 30
        }
        else
        {
            foreach ($Item in (Import-Csv $CSVPath\AllLicences.csv -Delimiter "`t"))
            {
                if ($NoNameTranslation)
                {
                    $SearchString = $Item.'AccountLicenseSKU'
                    $Selection = $worksheet.Range('A2').EntireColumn
                    $Search = $Selection.find($SearchString, [Type]::Missing, [Type]::Missing, 1)
                    $ResultCell = "A$($Search.Row)"
                    $worksheet.Hyperlinks.Add($worksheet.Range($ResultCell), '', "`'$($SearchString)`'!A1", "$($SearchString)", $worksheet.Range($ResultCell).text)
                }
                else
                {
                    $SearchString = $Item.'AccountLicenseSKU(Friendly)'
                    $Selection = $worksheet.Range('A2').EntireColumn
                    $Search = $Selection.find($SearchString, [Type]::Missing, [Type]::Missing, 1)
                    $ResultCell = "A$($Search.Row)"
                    $worksheet.Hyperlinks.Add($worksheet.Range($ResultCell), '', "`'$($SearchString)`'!A1", "$($SearchString)", $worksheet.Range($ResultCell).text)
                }
            }
            $worksheet.Move($worksheets.Item(1))
        }
        [void]$worksheet.UsedRange.Autofilter()
        $worksheet.UsedRange.EntireColumn.AutoFit()
    }
    $workbooks.Worksheets.Item('AllLicences').Select()

    $workbooks.SaveAs($XLOutput, 51)
    $workbooks.Saved = $true
    $workbooks.Close()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbooks) | Out-Null
    $excel.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

$date = Get-Date -Format yyyyMMdd
$OutputPath = Get-Item $OutputPath
if ($OutputPath.FullName -notmatch '\\$')
{
    $excelfilepath = $OutputPath.FullName + '\'
}
else
{
    $excelfilepath = $OutputPath.FullName
}
$XLOutput = $excelfilepath + "$CompanyName - $date.xlsx" ## Output file name
$CSVPath = $excelfilepath + $( -join ((65..90) + (97..122) | Get-Random -Count 14 | ForEach-Object { [char]$_ }))
$CSVPath = (New-Item -Type Directory -Path $csvpath).FullName
Write-Information 'Checking Connection to Office 365'
$test365 = Get-MsolCompanyInformation -ErrorAction silentlycontinue
if ($null -eq $test365)
{
    do
    {
        if ($Office365Credentials)
        {
            Connect-MsolService -Credential $Office365Credentials
        }
        else
        {
            Connect-MsolService
        }
        $test365 = Get-MsolCompanyInformation -ErrorAction silentlycontinue
    } while ($null -eq $test365)
}
Write-Information 'Connected to Office 365'
# Get a list of all licences that exist within the tenant
$licenseType = Get-MsolAccountSku
# Replace the above with the below if only a single SKU is required
#$licenseType = Get-MsolAccountSku | Where-Object {$_.AccountSkuID -like "*Power*"}
# Get all licences for a summary view
if ($NoNameTranslation)
{
    $licenseType | Where-Object { $_.TargetClass -eq 'User' } | Select-Object @{Name = 'AccountLicenseSKU'; Expression = { $($_.SkuPartNumber) } }, ActiveUnits, ConsumedUnits | Sort-Object 'AccountLicenseSKU' | Export-Csv $CSVPath\AllLicences.csv -NoTypeInformation -Delimiter `t
}
else
{
    $licenseType | Where-Object { $_.TargetClass -eq 'User' } | Select-Object @{Name = 'AccountLicenseSKU(Friendly)'; Expression = { $(RootLicenceswitch($_.SkuPartNumber)) } }, ActiveUnits, ConsumedUnits | Sort-Object 'AccountLicenseSKU(Friendly)' | Export-Csv $CSVPath\AllLicences.csv -NoTypeInformation -Delimiter `t
}
$licenseType = $licenseType | Where-Object { $_.ConsumedUnits -ge 1 }
#get all users with licence
Write-Information 'Retrieving all licensed users - this may take a while.'
$alllicensedusers = Get-MsolUser -All | Where-Object { $_.isLicensed -eq $true }
$licensedGroups = @{}
# Loop through all licence types found in the tenant
foreach ($license in $licenseType)
{
    Write-Information ('Gathering users with the following subscription: ' + $license.accountskuid)
    # Gather users for this particular AccountSku from pre-existing array of users
    $users = $alllicensedusers | Where-Object { $_.licenses.accountskuid -contains $license.accountskuid }
    if ($NoNameTranslation)
    {
        $rootLicence = ($($license.SkuPartNumber))
    }
    else
    {
        $rootLicence = RootLicenceswitch($($license.SkuPartNumber))
    }
    #$logFile = $CompanyName + "-" +$rootLicence + ".csv"
    $logFile = $CSVpath + '\' + $rootLicence + '.csv'
    $licensedUsers = New-Object System.Collections.Generic.List[System.Object]
    # Loop through all users and write them to the CSV file
    foreach ($user in $users)
    {
        Write-Verbose ('Processing ' + $user.displayname)
        $thislicense = $user.licenses | Where-Object { $_.accountskuid -eq $license.accountskuid }
        if ($user.BlockCredential -eq $true)
        {
            $enabled = $false
        }
        else
        {
            $enabled = $true
        }
        $userHashTable = @{
            DisplayName = $user.DisplayName
            UserPrincipalName = $user.UserPrincipalName
            AccountEnabled = $enabled
            AccountSKU = $rootLicence
        }
        if ($thislicense.GroupsAssigningLicense.Count -eq 0)
        {
            $userHashTable['DirectAssigned'] = $true
            $userHashTable['GroupsAssigning'] = $false
        }
        else
        {
            if ($thislicense.GroupsAssigningLicense -contains $user.ObjectID)
            {
                $groups = $thislicense.groupsassigninglicense.guid | Where-Object { $_ -notlike $user.objectid }
                if ($null -eq $groups)
                {
                    $groups = $false
                }
                else
                {
                    foreach ($group in $groups)
                    {
                        if ($null -eq $licensedGroups[$group])
                        {
                            $getGroup = Get-MsolGroup -ObjectId $group
                            $licensedGroups[$group] = $getGroup.DisplayName
                        }
                    }
                    $groups = (Invoke-GroupGuidConversion -GroupGuid $groups -LicenseGroups $licensedGroups).DisplayName -Join ';'
                }
                $userHashTable['DirectAssigned'] = $true
                $userHashTable['GroupsAssigning'] = $groups
            }
            else
            {
                $groups = $thislicense.groupsassigninglicense.guid
                if ($null -eq $groups)
                {
                    $groups = $false
                }
                else
                {
                    foreach ($group in $groups)
                    {
                        if ($null -eq $licensedGroups[$group])
                        {
                            $getGroup = Get-MsolGroup -ObjectId $group
                            $licensedGroups[$group] = $getGroup.DisplayName
                        }
                    }
                    $groups = (Invoke-GroupGuidConversion -GroupGuid $groups -LicenseGroups $licensedGroups).DisplayName -Join ';'
                }
                $userHashTable['DirectAssigned'] = $false
                $userHashTable['GroupsAssigning'] = $groups
            }
        }
        foreach ($row in $($thislicense.ServiceStatus))
        {
            $serviceName = componentlicenseswitch([string]($row.ServicePlan.ServiceName))
            $userHashTable[$serviceName] = ($thislicense.ServiceStatus | Where-Object { $_.ServicePlan.ServiceName -eq $row.ServicePlan.ServiceName }).ProvisioningStatus
        }
        $licensedUsers.Add([PSCustomObject]$userHashTable) | Out-Null
    }
    $licensedUsers | Select-Object DisplayName, UserPrincipalName, AccountEnabled, AccountSKU, DirectAssigned, GroupsAssigning, * -ErrorAction SilentlyContinue | Export-Csv -Path $logFile -Delimiter "`t" -Encoding UTF8 -NoClobber -NoTypeInformation
}
Write-Information ('Merging CSV Files')
Merge-CSVFile -CSVPath $CSVPath -XLOutput $XLOutput | Out-Null
Write-Information ('Tidying up - deleting CSV Files')
Remove-Item $CSVPath -Recurse -Confirm:$false -Force
Write-Information ('CSV Files Deleted')
Write-Information ("Script Completed.  Results available in $XLOutput")
$InformationPreference = $initialInformationPreference
