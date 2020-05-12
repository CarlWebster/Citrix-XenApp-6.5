#This File is in Unicode format.  Do not edit in an ASCII editor.

#region help text

<#
.SYNOPSIS
	Creates an inventory of a Citrix XenApp 6.5 farm.
.DESCRIPTION
	Creates an inventory of a Citrix XenApp 6.5 Farm using Microsoft PowerShell, Word,
	plain text or HTML.
	
	Script runs fastest in PowerShell version 5.

	Word is NOT needed to run the script. This script will output in Text and HTML.
	
	You do NOT have to run this script on a Collector. This script was developed and run 
	from a Windows 7 VM. Unfortunately, Citrix did not add remoting support to the Group
	Policy module. If Policy information is required, the script will need to be run on 
	a Collector.
	
	You can run this script remotely using the –AdminAddress (AA) parameter.

	Creates an output file named after the XenApp 6.5 Farm.
	
	Word and PDF Document includes a Cover Page, Table of Contents and Footer.
	Includes support for the following language versions of Microsoft Word:
		Catalan
		Chinese
		Danish
		Dutch
		English
		Finnish
		French
		German
		Norwegian
		Portuguese
		Spanish
		Swedish
		
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
.PARAMETER CompanyAddress
	Company Address to use for the Cover Page, if the Cover Page has the Address field.  
		The following Cover Pages have an Address field:
			Banded (Word 2013/2016)
			Contrast (Word 2010)
			Exposure (Word 2010)
			Filigree (Word 2013/2016)
			Ion (Dark) (Word 2013/2016)
			Retrospect (Word 2013/2016)
			Semaphore (Word 2013/2016)
			Tiles (Word 2010)
			ViewMaster (Word 2013/2016)
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CA.
.PARAMETER CompanyEmail
	Company Email to use for the Cover Page, if the Cover Page has the Email field.  
		The following Cover Pages have an Email field:
			Facet (Word 2013/2016)
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CE.
.PARAMETER CompanyFax
	Company Fax to use for the Cover Page, if the Cover Page has the Fax field.  
		The following Cover Pages have a Fax field:
			Contrast (Word 2010)
			Exposure (Word 2010)
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CF.
.PARAMETER CompanyPhone
	Company Phone to use for the Cover Page, if the Cover Page has the Phone field.  
		The following Cover Pages have a Phone field:
			Contrast (Word 2010)
			Exposure (Word 2010)
	This parameter is only valid with the MSWORD and PDF output parameters.
	This parameter has an alias of CPh.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	Only Word 2010, 2013 and 2016 are supported.
	(default cover pages in Word en-US)
	
	Valid input is:
		Alphabet (Word 2010. Works)
		Annual (Word 2010. Doesn't work well for this report)
		Austere (Word 2010. Works)
		Austin (Word 2010/2013/2016. Doesn't work in 2013 or 2016, mostly works in 2010 but 
						Subtitle/Subject & Author fields need to be moved 
						after title box is moved up)
		Banded (Word 2013/2016. Works)
		Conservative (Word 2010. Works)
		Contrast (Word 2010. Works)
		Cubicles (Word 2010. Works)
		Exposure (Word 2010. Works if you like looking sideways)
		Facet (Word 2013/2016. Works)
		Filigree (Word 2013/2016. Works)
		Grid (Word 2010/2013/2016. Works in 2010)
		Integral (Word 2013/2016. Works)
		Ion (Dark) (Word 2013/2016. Top date doesn't fit, box needs to be manually resized or font 
						changed to 8 point)
		Ion (Light) (Word 2013/2016. Top date doesn't fit, box needs to be manually resized or font 
						changed to 8 point)
		Mod (Word 2010. Works)
		Motion (Word 2010/2013/2016. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2010. Works)
		Puzzle (Word 2010. Top date doesn't fit, box needs to be manually resized or font 
					changed to 14 point)
		Retrospect (Word 2013/2016. Works)
		Semaphore (Word 2013/2016. Works)
		Sideline (Word 2010/2013/2016. Doesn't work in 2013 or 2016, works in 2010)
		Slice (Dark) (Word 2013/2016. Doesn't work)
		Slice (Light) (Word 2013/2016. Doesn't work)
		Stacks (Word 2010. Works)
		Tiles (Word 2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2010. Works)
		ViewMaster (Word 2013/2016. Works)
		Whisp (Word 2013/2016. Works)
		
	Default value is Sideline.
	This parameter has an alias of CP.
	This parameter is only valid with the MSWORD and PDF output parameters.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER AdminAddress
	Specifies the address of a XenApp Collector the PowerShell snapins will connect to. 
	The Collector cannot be a Session-Host only server.
	This can be provided as a host name or an IP address. 
	This parameter defaults to nothing to allow the connection to be set outside the script.
	This parameter has an alias of AA.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER Text
	Creates a formatted text file with a .txt extension.
	This parameter is disabled by default.
.PARAMETER MSWord
	SaveAs DOCX file
	This parameter is set True if no other output format is selected.
.PARAMETER HTML
	Creates an HTML file with an .html extension.
	This parameter is disabled by default.
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network 
	Interface Cards

	This parameter may require the script be run from an elevated PowerShell session 
	using an account with permission to retrieve hardware information (i.e. Domain Admin or 
	Local Administrator).

	Selecting this parameter will add to both the time it takes to run the script and size 
	of the report.

	This parameter is disabled by default.
	This parameter has an alias of HW.
.PARAMETER Software
	Gather software installed by querying the registry.  
	Use SoftwareExclusions.txt to exclude software from the report.
	SoftwareExclusions.txt must exist, and be readable, in the same folder as this script.
	SoftwareExclusions.txt can be an empty file to have no installed applications excluded.
	See Get-Help About-Wildcards for help on formatting the lines to exclude applications.
	This parameter is disabled by default.
	This parameter has an alias of SW.
.PARAMETER StartDate
	Start date, in MM/DD/YYYY HH:MM format, for the Configuration Logging report.
	Default is today's date minus seven days.
	If the StartDate is entered as 01/01/2019, the date becomes 01/01/2019 00:00:00.
	This parameter has an alias of SD.
.PARAMETER EndDate
	End date, in MM/DD/YYYY HH:MM format, for the Configuration Logging report.
	Default is today's date.
	If the EndDate is entered as 01/01/2019, the date becomes 01/01/2019 00:00:00.
	This parameter has an alias of ED.
.PARAMETER Summary
	Only give summary information, no details.
	This parameter is disabled by default.
	This parameter cannot be used with either the Hardware, Software, StartDate or EndDate parameters.
.PARAMETER AddDateTime
	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2019 at 6PM is 2019-06-01_1800.
	Output filename will be ReportName_2019-06-01_1800.docx (or .pdf).
	This parameter is disabled by default.
	This parameter has an alias of ADT.
.PARAMETER Section
	Processes a specific section of the report.
	Valid options are:
		Admins (Administrators)
		Apps (Applications)
		ConfigLog (Configuration Logging)
		LBPolicies (Load Balancing Policies)
		LoadEvals (Load Evaluators)
		Policies
		Servers
		WGs (Worker Groups)
		Zones
		All
	This parameter defaults to All sections.
.PARAMETER NoPolicies
	Excludes all Farm and Citrix AD based policy information from the output document.
	
	Using the NoPolicies parameter will cause the Policies section to be set to False.
	
	This parameter is disabled by default.
	This parameter has an alias of NP.
.PARAMETER NoADPolicies
	Excludes all Citrix AD based policy information from the output document.
	Includes only Farm policies created in AppCenter.
	
	This switch is useful in large AD environments, where there may be thousands
	of policies, to keep SYSVOL from being searched.
	
	This parameter is disabled by default.
	This parameter has an alias of NoAD.
.PARAMETER Policies
	Give detailed information for both Site and Citrix AD based Policies.
	
	Using the Policies parameter can cause the report to take a very long time 
	to complete and can generate an extremely long report.
	
	Note: The Citrix Group Policy PowerShell module will not load from an elevated 
	PowerShell session. 
	If the module is manually imported, the module is not detected from an elevated 
	PowerShell session.
	
	There are three related parameters: Policies, NoPolicies, and NoADPolicies.
	
	Policies and NoPolicies are mutually exclusive and priority is given to NoPolicies.
	
	This parameter is disabled by default.
	This parameter has an alias of Pol.
.PARAMETER Folder
	Specifies the optional output folder to save the output report. 
.PARAMETER MaxDetails
	Adds maximum detail to the report.
	
	This is the same as using the following parameters:
		Administrators
		Applications
		HardWare
		Logging
		Policies
		Software

	Does not change the value of NoADPolicies.
	
	WARNING: Using this parameter can create an extremely large report and 
	can take a very long time to run.

	This parameter has an alias of MAX.
.PARAMETER Dev
	Clears errors at the beginning of the script.
	Outputs all errors to a text file at the end of the script.
	
	This is used when the script developer requests more troubleshooting data.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
.PARAMETER ScriptInfo
	Outputs information about the script to a text file.
	The text file is placed in the same folder from where the script is run.
	
	This parameter is disabled by default.
	This parameter has an alias of SI.
.PARAMETER Log
	Generates a log file for troubleshooting.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -AdminAddress XA65ZDC
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	XA65ZDC as the remote Collector to run the script against.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -PDF 
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -TEXT

	This parameter is reserved for a future update and no output is created at this time.
	
	Will use all default values and save the document as a formatted text file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -HTML

	This parameter is reserved for a future update and no output is created at this time.
	
	Will use all default values and save the document as an HTML file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -Summary
	
	Creates a Summary report with no detail.
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -PDF -Summary 
	
	Creates a Summary report with no detail.
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -Hardware 
	
	Will use all Default values and add additional information for each server about its 
	hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -Software 
	
	Will use all Default values and add additional information for each server about its 
	installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -StartDate "01/01/2019" -EndDate "01/02/2019" 
	
	Will use all Default values and add additional information for each server about its 
	installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will return all Configuration Logging entries from "01/01/2019 00:00:00" through "01/02/2019 "00:00:00".
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -StartDate "01/01/2019" -EndDate "01/01/2019" 
	
	Will use all Default values and add additional information for each server about its 
	installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will return all Configuration Logging entries from "01/01/2019 00:00:00" through 
	"01/01/2019 "00:00:00".  In other words, nothing is returned.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -StartDate "01/01/2019 21:00:00" 
	-EndDate "01/01/2019 22:00:00" 
	
	Will use all Default values and add additional information for each server about its 
	installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will return all Configuration Logging entries from 9PM to 10PM on 01/01/2019.
.EXAMPLE
	PS C:\PSScript .\XA65_Inventory_V5.ps1 -CompanyName "Carl Webster Consulting" 
	-CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\XA65_Inventory_V5.ps1 -CN "Carl Webster Consulting" -CP "Mod" 
	-UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -Section Policies
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Processes only the Policies section of the report.
	Includes both Farm and AD policies.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_v5.ps1 -NoADPolicies
	
	Creates a report with full details on Farm policies created in AppCenter but 
	no Citrix AD based Policy information.
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_v5.ps1 -Section Policies -NoADPolicies
	
	Creates a report with full details on Farm policies created in AppCenter but 
	no Citrix AD based Policy information.
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -AddDateTime
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2019 at 6PM is 2019-06-01_1800.
	Output filename will be XA65FarmName_2019-06-01_1800.docx
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -PDF -AddDateTime
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will display verbose messages as the script is running.

	Adds a date time stamp to the end of the file name.
	Time stamp is in the format of yyyy-MM-dd_HHmm.
	June 1, 2019 at 6PM is 2019-06-01_1800.
	Output filename will be XA65FarmName_2019-06-01_1800.pdf
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -Folder \\FileServer\ShareName
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Output file will be saved in the path \\FileServer\ShareName
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V5.ps1 -Dev -ScriptInfo -Log
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl 
	Webster" or 
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	
	Creates a text file named XA65V5InventoryScriptErrors_yyyy-MM-dd_HHmm.txt that 
	contains up to the last 250 errors reported by the script.
	
	Creates a text file named XA65V5InventoryScriptInfo_yyyy-MM-dd_HHmm.txt that 
	contains all the script parameters and other basic information.
	
	Creates a text file for transcript logging named 
	XA65V5DocScriptTranscript_yyyy-MM-dd_HHmm.txt.
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word or PDF document.
.NOTES
	NAME: XA65_Inventory_V5.ps1
	VERSION: 5.01
	AUTHOR: Carl Webster (with a lot of help from Michael B. Smith, Jeff Wouters and Iain Brighton)
	LASTEDIT: April 21, 2019
#>

#endregion

#region script parameters
#thanks to @jeffwouters and Michael B. Smith for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "Word") ]

Param(
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[Switch]$HTML=$False,

	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[Switch]$PDF=$False,

	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[Switch]$Text=$False,

	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Switch]$MSWord=$False,

	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('ADT')]
	[Switch]$AddDateTime=$False,
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[ValidateNotNullOrEmpty()]
	[Alias('AA')]
	[string]$AdminAddress='Localhost',

	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('Admins')]
	[Switch]$Administrators=$False,	
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('Apps')]
	[Switch]$Applications=$False,	
	
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('CA')]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyAddress='',
    
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('CE')]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyEmail='',
    
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('CF')]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyFax='',
    
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('CN')]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName='',
    
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('CPh')]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyPhone='',
    
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage='Sideline', 

	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Switch]$Dev=$False,
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('ED')]
	[Datetime]$EndDate = (Get-Date -displayhint date),
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[string]$Folder='',
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('HW')]
	[Switch]$Hardware=$False,

	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Switch]$Log=$False,
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Switch]$Logging=$False,	
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('MAX')]
	[Switch]$MaxDetails=$False,

	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('NoAD')]
	[Switch]$NoADPolicies=$False,	
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('NP')]
	[Switch]$NoPolicies=$False,	
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('Pol')]
	[Switch]$Policies=$False,	
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('SI')]
	[Switch]$ScriptInfo=$False,
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[string]$Section='All',
	
	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('SW')]
	[Switch]$Software=$False,

	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('SD')]
	[Datetime]$StartDate = ((Get-Date -displayhint date).AddDays(-7)),

	[parameter(ParameterSetName='HTML',Mandatory=$False)] 
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Text',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Switch]$Summary=$False,	
	
	[parameter(ParameterSetName='PDF',Mandatory=$False)] 
	[parameter(ParameterSetName='Word',Mandatory=$False)] 
	[Alias('UN')]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username

	)

#endregion

#region script change log	
#Original Script created 8/17/2010 by Michael Bogobowicz, Citrix Systems.
#To contact, please message @mikebogo on Twitter
#This script is designed to be run on a XenApp 6.5 server

#Modifications by Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#Version 5.00 created on June 1, 2015

#V5.01 21-Apr-2019
#	If Policies parameter is used, check to see if PowerShell session is elevated. If it is,
#		abort the script. This is the #2 support email.
#		Added a Note to the Help Text and ReadMe file about the Citrix.GroupPolicy.Commands module:
#		Note: The Citrix Group Policy PowerShell module will not load from an elevated PowerShell session. 
#		If the module is manually imported, the module is not detected from an elevated PowerShell session.

##V5.00 released to the community 14-Dec-2018
#	Removed minimum requirement for PowerShell V3
#	Fixed all code to make it work in PowerShell V2
#	Removed all SMTP related code as we could not could that code to work with PowerShell V2
#	Added HTML and Text output options
#	Added parameters to bring the code up to the same standard as the other documentation scripts
#		AdminAddress
#		MaxDetails
#		Dev
#		ScriptInfo
#		Log
#		NoPolicies
#		NoADPolicies
#endregion

#region initial variable testing and setup
Set-StrictMode -Version 2

#force  on
$SaveEAPreference = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'

If($Null -eq $AdminAddress)
{
	$AdminAddress = "LocalHost"
}
If($Null -eq $CompanyAddress)
{
	$CompanyAddress = ""
}
If($Null -eq $CompanyEmail)
{
	$CompanyEmail = ""
}
If($Null -eq $CompanyFax)
{
	$CompanyFax = ""
}
If($Null -eq $CompanyName)
{
	$CompanyName = ""
}
If($Null -eq $CompanyPhone)
{
	$CompanyPhone = ""
}
If($Null -eq $CoverPage)
{
	$CoverPage="Sideline"
}
If($Null -eq $UserName)
{
	$UserName=$env:username
}
If($Null -eq $HTML)
{
	$HTML = $False
}
If($Null -eq $MSWord)
{
	$MSWord = $False
}
If($Null -eq $PDF)
{
	$PDF = $False
}
If($Null -eq $Text)
{
	$Text = $False
}
If($Null -eq $Administrators)
{
	$Administrators=$False
}
If($Null -eq $Applications)
{
	$Applications=$False
}
If($Null -eq $Logging)
{
	$Logging=$False
}
If($Null -eq $StartDate)
{
	$StartDate = ((Get-Date -displayhint date).AddDays(-7))
}
If($Null -eq $EndDate)
{
	$EndDate = (Get-Date -displayhint date)
}
If($Null -eq $Summary)
{
	$Summary = $False
}
If($Null -eq $MaxDetails)
{
	$MaxDetails=$False
}
If($Null -eq $Policies)
{
	$Policies = $False
}
If($Null -eq $NoPolicies)
{
	$NoPolicies = $False
}
If($Null -eq $NoADPolicies)
{
	$NoADPolicies = $False
}
If($Null -eq $AddDateTime)
{
	$AddDateTime = $False
}
If($Null -eq $Folder)
{
	$Folder = ""
}
If($Null -eq $Hardware)
{
	$Hardware = $False
}
If($Null -eq $Software)
{
	$Software = $False
}
If($Null -eq $Section)
{
	$Section = "All"
}
If($Null -eq $Dev)
{
	$Dev = $False
}
If($Null -eq $ScriptInfo)
{
	$ScriptInfo = $False
}
If($Null -eq $Log)
{
	$Log = $False
}

If(!(Test-Path Variable:AdminAddress))
{
	$AdminAddress = "LocalHost"
}
If(!(Test-Path Variable:CompanyAddress))
{
	$CompanyAddress = ""
}
If(!(Test-Path Variable:CompanyEmail))
{
	$CompanyEmail = ""
}
If(!(Test-Path Variable:CompanyFax))
{
	$CompanyFax = ""
}
If(!(Test-Path Variable:CompanyName))
{
	$CompanyName = ""
}
If(!(Test-Path Variable:CompanyPhone))
{
	$CompanyPhone = ""
}
If(!(Test-Path Variable:CoverPage))
{
	$CoverPage="Sideline"
}
If(!(Test-Path Variable:UserName))
{
	$UserName=$env:username
}
If(!(Test-Path Variable:HTML))
{
	$HTML = $False
}
If(!(Test-Path Variable:MSWord))
{
	$MSWord = $False
}
If(!(Test-Path Variable:PDF))
{
	$PDF = $False
}
If(!(Test-Path Variable:Text))
{
	$Text = $False
}
If(!(Test-Path Variable:Administrators))
{
	$Administrators=$False
}
If(!(Test-Path Variable:Applications))
{
	$Applications=$False
}
If(!(Test-Path Variable:Logging))
{
	$Logging=$False
}
If(!(Test-Path Variable:StartDate))
{
	$StartDate = ((Get-Date -displayhint date).AddDays(-7))
}
If(!(Test-Path Variable:EndDate))
{
	$EndDate = (Get-Date -displayhint date)
}
If(!(Test-Path Variable:Summary))
{
	$Summary = $False
}
If(!(Test-Path Variable:MaxDetails))
{
	$MaxDetails=$False
}
If(!(Test-Path Variable:Policies))
{
	$Policies = $False
}
If(!(Test-Path Variable:NoPolicies))
{
	$NoPolicies = $False
}
If(!(Test-Path Variable:NoADPolicies))
{
	$NoADPolicies = $False
}
If(!(Test-Path Variable:AddDateTime))
{
	$AddDateTime = $False
}
If(!(Test-Path Variable:Folder))
{
	$Folder = ""
}
If(!(Test-Path Variable:Hardware))
{
	$Hardware = $False
}
If(!(Test-Path Variable:Software))
{
	$Software = $False
}
If(!(Test-Path Variable:Section))
{
	$Section = "All"
}
If(!(Test-Path Variable:Dev))
{
	$Dev = $False
}
If(!(Test-Path Variable:ScriptInfo))
{
	$ScriptInfo = $False
}
If(!(Test-Path Variable:Log))
{
	$Log = $False
}

If($Null -eq $MSWord)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

If($Log) 
{
	#start transcript logging
	$Script:ThisScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Definition
	$Script:LogPath = "$Script:ThisScriptPath\XA65V5DocScriptTranscript_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	
	try 
	{
		Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
		Write-Host "$(Get-Date): Transcript/log started at $Script:LogPath" -BackgroundColor Black -ForegroundColor Yellow
		$Script:StartLog = $true
	} 
	catch 
	{
		Write-Host "$(Get-Date): Transcript/log failed at $Script:LogPath" -BackgroundColor Black -ForegroundColor Yellow
		$Script:StartLog = $false
	}
}

If($Dev)
{
	$Error.Clear()
	$Script:DevErrorFile = "$($pwd.Path)\XA65V5InventoryScriptErrors_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
}

If($Null -eq $MSWord)
{
	If($Text -or $HTML -or $PDF)
	{
		$MSWord = $False
	}
	Else
	{
		$MSWord = $True
	}
}

If($MSWord -eq $False -and $PDF -eq $False -and $Text -eq $False -and $HTML -eq $False)
{
	$MSWord = $True
}

Write-Host "$(Get-Date): Testing output parameters" -BackgroundColor Black -ForegroundColor Yellow

If($MSWord)
{
	Write-Host "$(Get-Date): MSWord is set" -BackgroundColor Black -ForegroundColor Yellow
}
ElseIf($PDF)
{
	Write-Host "$(Get-Date): PDF is set" -BackgroundColor Black -ForegroundColor Yellow
}
ElseIf($Text)
{
	Write-Host "$(Get-Date): Text is set" -BackgroundColor Black -ForegroundColor Yellow
}
ElseIf($HTML)
{
	Write-Host "$(Get-Date): HTML is set" -BackgroundColor Black -ForegroundColor Yellow
}
Else
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Host "$(Get-Date): Unable to determine output parameter" -BackgroundColor Black -ForegroundColor Yellow
	If($Null -eq $MSWord)
	{
		Write-Host "$(Get-Date): MSWord is Null" -BackgroundColor Black -ForegroundColor Yellow
	}
	ElseIf($Null -eq $PDF)
	{
		Write-Host "$(Get-Date): PDF is Null" -BackgroundColor Black -ForegroundColor Yellow
	}
	ElseIf($Null -eq $Text)
	{
		Write-Host "$(Get-Date): Text is Null" -BackgroundColor Black -ForegroundColor Yellow
	}
	ElseIf($Null -eq $HTML)
	{
		Write-Host "$(Get-Date): HTML is Null" -BackgroundColor Black -ForegroundColor Yellow
	}
	Else
	{
		Write-Host "$(Get-Date): MSWord is $($MSWord)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): PDF is $($PDF)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): Text is $($Text)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): HTML is $($HTML)" -BackgroundColor Black -ForegroundColor Yellow
	}
	Write-Error "Unable to determine output parameter.  Script cannot continue"
	Exit
}

#If the MaxDetails parameter is used, set a bunch of stuff true and some stuff false
If($MaxDetails)
{
	$Administrators		= $True
	$Applications		= $True
	$HardWare			= $True
	$Logging			= $True
	$Policies			= $True
	$Software			= $True
	
	$NoPolicies			= $False
	$Section			= "All"
}

If($NoPolicies)
{
	$Policies = $False
}

If($NoPolicies -and $Section -eq "Policies")
{
	#conflict
	$ErrorActionPreference = $SaveEAPreference
	Write-Error -Message "`n`tYou specified conflicting parameters.`n`n`tYou specified the $($Section) section but also selected NoPolicies.`n`n`tPlease change one of these options and rerun the script.`n`n
	Script cannot continue."
	Exit
}

$ValidSection = $False
Switch ($Section)
{
	"Admins"		{$ValidSection = $True; Break}
	"Apps"			{$ValidSection = $True; Break}
	"ConfigLog"		{$ValidSection = $True; $Logging = $True; Break}	#force $logging true if the config logging section is specified}
	"LBPolicies"	{$ValidSection = $True; Break}
	"LoadEvals"		{$ValidSection = $True; Break}
	"Policies"		{$ValidSection = $True; $Policies = $True; Break} #force $policies true if the policies section is specified
	"Servers"		{$ValidSection = $True; Break}
	"WGs"			{$ValidSection = $True; Break}
	"Zones"			{$ValidSection = $True; Break}
	"All"			{$ValidSection = $True; Break}
}

If($ValidSection -eq $False)
{
	$ErrorActionPreference = $SaveEAPreference
	Write-Error -Message "`n`tThe Section parameter specified, $($Section), is an invalid Section option.`n`tValid options are:
	
	`t`tAdmins
	`t`tApps
	`t`tConfigLog
	`t`tLBPolicies
	`t`tLoadEvals
	`t`tPolicies
	`t`tServers
	`t`tWGs
	`t`tZones
	`t`tAll
	
	`tScript cannot continue."
	Exit
}

If($Folder -ne "")
{
	Write-Host "$(Get-Date): Testing folder path" -BackgroundColor Black -ForegroundColor Yellow
	#does it exist
	If(Test-Path $Folder -EA 0)
	{
		#it exists, now check to see if it is a folder and not a file
		If(Test-Path $Folder -pathType Container -EA 0)
		{
			#it exists and it is a folder
			Write-Host "$(Get-Date): Folder path $Folder exists and is a folder" -BackgroundColor Black -ForegroundColor Yellow
		}
		Else
		{
			#it exists but it is a file not a folder
			Write-Error "Folder $Folder is a file, not a folder.  Script cannot continue"
			Exit
		}
	}
	Else
	{
		#does not exist
		Write-Error "Folder $Folder does not exist.  Script cannot continue"
		Exit
	}
}

#V5.01  Add check if $Policies -eq $True, see if PowerShell session is elevated
#		If session is elevated, abort the script
Function ElevatedSession
{
	#added in V5.01
	$currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

	If($currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator ))
	{
		Write-Verbose "$(Get-Date): This is an elevated PowerShell session"
		Return $True
	}
	Else
	{
		Write-Verbose "$(Get-Date): This is NOT an elevated PowerShell session" -Foreground White
		Return $False
	}
}

If($Policies -eq $True)
{
	Write-Verbose "$(Get-Date): Testing for elevated PowerShell session."
	#see if session is elevated
	$Elevated = ElevatedSession
	
	If($Elevated -eq $True)
	{
		#abort script
		Write-Error "
		`n
		`n
		`tThe Citrix Group Policy module cannot be loaded or found in an elevated PowerShell session.
		`n
		`n
		`tThe Policies parameter was used and this is an elevated PowerShell session.
		`n
		`n
		`tRerun the script from a non-elevated PowerShell session. The script will now close.
		`n
		`n"
		Write-Verbose "$(Get-Date): "
		Exit
	}
}
#endregion

#region initialize variables for Word, HTML, and text
[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption

If($MSWord -or $PDF)
{
	#try and fix the issue with the $CompanyName variable
	$Script:CoName = $CompanyName
	Write-Host "$(Get-Date): CoName is $($Script:CoName)" -BackgroundColor Black -ForegroundColor Yellow
	
	#the following values were attained from 
	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
	#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
	[int]$wdAlignPageNumberRight = 2
	[long]$wdColorGray15 = 14277081
	[int]$wdMove = 0
	[int]$wdSeekMainDocument = 0
	[int]$wdSeekPrimaryFooter = 4
	[int]$wdStory = 6
	[int]$wdColorRed = 255
	[int]$wdWord2007 = 12
	[int]$wdWord2010 = 14
	[int]$wdWord2013 = 15
	[int]$wdWord2016 = 16
	[int]$wdFormatDocumentDefault = 16
	[int]$wdFormatPDF = 17
	#http://blogs.technet.com/b/heyscriptingguy/archive/2006/03/01/how-can-i-right-align-a-single-column-in-a-word-table.aspx
	#http://msdn.microsoft.com/en-us/library/office/ff835817%28v=office.15%29.aspx
	[int]$wdAlignParagraphLeft = 0
	[int]$wdAlignParagraphCenter = 1
	[int]$wdAlignParagraphRight = 2
	#http://msdn.microsoft.com/en-us/library/office/ff193345%28v=office.15%29.aspx
	[int]$wdCellAlignVerticalTop = 0
	[int]$wdCellAlignVerticalCenter = 1
	[int]$wdCellAlignVerticalBottom = 2
	#http://msdn.microsoft.com/en-us/library/office/ff844856%28v=office.15%29.aspx
	[int]$wdAutoFitFixed = 0
	[int]$wdAutoFitContent = 1
	[int]$wdAutoFitWindow = 2
	#http://msdn.microsoft.com/en-us/library/office/ff821928%28v=office.15%29.aspx
	[int]$wdAdjustNone = 0
	[int]$wdAdjustProportional = 1
	[int]$wdAdjustFirstColumn = 2
	[int]$wdAdjustSameWidth = 3

	[int]$PointsPerTabStop = 36
	[int]$Indent0TabStops = 0 * $PointsPerTabStop
	[int]$Indent1TabStops = 1 * $PointsPerTabStop
	[int]$Indent2TabStops = 2 * $PointsPerTabStop
	[int]$Indent3TabStops = 3 * $PointsPerTabStop
	[int]$Indent4TabStops = 4 * $PointsPerTabStop

	# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
	[int]$wdStyleHeading1 = -2
	[int]$wdStyleHeading2 = -3
	[int]$wdStyleHeading3 = -4
	[int]$wdStyleHeading4 = -5
	[int]$wdStyleNoSpacing = -158
	[int]$wdTableGrid = -155
	[int]$wdTableLightListAccent3 = -206

	#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/org/codehaus/groovy/scriptom/tlb/office/word/WdLineStyle.html
	[int]$wdLineStyleNone = 0
	[int]$wdLineStyleSingle = 1

	[int]$wdHeadingFormatTrue = -1
	[int]$wdHeadingFormatFalse = 0 
	
	[string]$Script:RunningOS = (Get-WmiObject -class Win32_OperatingSystem -EA 0).Caption
}
Else
{
	$Script:CoName = ""
}

If($HTML)
{
    Set-Variable htmlredmask         -Option AllScope -Value "#FF0000"
    Set-Variable htmlcyanmask        -Option AllScope -Value "#00FFFF"
    Set-Variable htmlbluemask        -Option AllScope -Value "#0000FF"
    Set-Variable htmldarkbluemask    -Option AllScope -Value "#0000A0"
    Set-Variable htmllightbluemask   -Option AllScope -Value "#ADD8E6"
    Set-Variable htmlpurplemask      -Option AllScope -Value "#800080"
    Set-Variable htmlyellowmask      -Option AllScope -Value "#FFFF00"
    Set-Variable htmllimemask        -Option AllScope -Value "#00FF00"
    Set-Variable htmlmagentamask     -Option AllScope -Value "#FF00FF"
    Set-Variable htmlwhitemask       -Option AllScope -Value "#FFFFFF"
    Set-Variable htmlsilvermask      -Option AllScope -Value "#C0C0C0"
    Set-Variable htmlgraymask        -Option AllScope -Value "#808080"
    Set-Variable htmlblackmask       -Option AllScope -Value "#000000"
    Set-Variable htmlorangemask      -Option AllScope -Value "#FFA500"
    Set-Variable htmlmaroonmask      -Option AllScope -Value "#800000"
    Set-Variable htmlgreenmask       -Option AllScope -Value "#008000"
    Set-Variable htmlolivemask       -Option AllScope -Value "#808000"

    Set-Variable htmlbold        -Option AllScope -Value 1
    Set-Variable htmlitalics     -Option AllScope -Value 2
    Set-Variable htmlred         -Option AllScope -Value 4
    Set-Variable htmlcyan        -Option AllScope -Value 8
    Set-Variable htmlblue        -Option AllScope -Value 16
    Set-Variable htmldarkblue    -Option AllScope -Value 32
    Set-Variable htmllightblue   -Option AllScope -Value 64
    Set-Variable htmlpurple      -Option AllScope -Value 128
    Set-Variable htmlyellow      -Option AllScope -Value 256
    Set-Variable htmllime        -Option AllScope -Value 512
    Set-Variable htmlmagenta     -Option AllScope -Value 1024
    Set-Variable htmlwhite       -Option AllScope -Value 2048
    Set-Variable htmlsilver      -Option AllScope -Value 4096
    Set-Variable htmlgray        -Option AllScope -Value 8192
    Set-Variable htmlolive       -Option AllScope -Value 16384
    Set-Variable htmlorange      -Option AllScope -Value 32768
    Set-Variable htmlmaroon      -Option AllScope -Value 65536
    Set-Variable htmlgreen       -Option AllScope -Value 131072
    Set-Variable htmlblack       -Option AllScope -Value 262144
}

If($TEXT)
{
	$Script:output = ""
}
#endregion

#region code for -hardware switch
Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com
	# modified 1-May-2014 to work in trusted AD Forests and using different domain admin credentials	
	# modified 17-Aug-2016 to fix a few issues with Text and HTML output
	# modified 29-Apr-2018 to change from Arrays to New-Object System.Collections.ArrayList

	#Get Computer info
	Write-Host "$(Get-Date): `t`tProcessing WMI Computer information" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): `t`t`tHardware information" -BackgroundColor Black -ForegroundColor Yellow
	If($MSWord -or $PDF)
	{
		WriteWordLine 3 0 "Computer Information: $($RemoteComputerName)"
		WriteWordLine 4 0 "General Computer"
	}
	ElseIf($Text)
	{
		Line 0 "Computer Information: $($RemoteComputerName)"
		Line 1 "General Computer"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 3 0 "Computer Information: $($RemoteComputerName)"
	}
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
	}
	
	Catch
	{
		$Results = $Null
	}
	
	If($? -and $Null -ne $Results)
	{
		$ComputerItems = $Results | Select-Object Manufacturer, Model, Domain, `
		@{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}, `
		NumberOfProcessors, NumberOfLogicalProcessors
		$Results = $Null
		[string]$ComputerOS = (Get-WmiObject -class Win32_OperatingSystem -computername $RemoteComputerName -EA 0).Caption

		ForEach($Item in $ComputerItems)
		{
			OutputComputerItem $Item $ComputerOS
		}
	}
	ElseIf(!$?)
	{
		Write-Host "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
			Line 2 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Host "$(Get-Date): No results Returned for Computer information" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Computer information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Computer information" "" $Null 0 $False $True
		}
	}
	
	#Get Disk info
	Write-Host "$(Get-Date): `t`t`tDrive information" -BackgroundColor Black -ForegroundColor Yellow

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Drive(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Drive(s)"
	}
	ElseIf($HTML)
	{
	}

	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$drives = $Results | Select-Object caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				OutputDriveItem $drive
			}
		}
	}
	ElseIf(!$?)
	{
		Write-Host "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Host "$(Get-Date): No results Returned for Drive information" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Drive information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Drive information" "" $Null 0 $False $True
		}
	}
	

	#Get CPU's and stepping
	Write-Host "$(Get-Date): `t`t`tProcessor information" -BackgroundColor Black -ForegroundColor Yellow

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Processor(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Processor(s)"
	}
	ElseIf($HTML)
	{
	}

	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Processors = $Results | Select-Object availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
		ForEach($processor in $processors)
		{
			OutputProcessorItem $processor
		}
	}
	ElseIf(!$?)
	{
		Write-Host "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Host "$(Get-Date): No results Returned for Processor information" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for Processor information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for Processor information" "" $Null 0 $False $True
		}
	}

	#Get Nics
	Write-Host "$(Get-Date): `t`t`tNIC information" -BackgroundColor Black -ForegroundColor Yellow

	If($MSWord -or $PDF)
	{
		WriteWordLine 4 0 "Network Interface(s)"
	}
	ElseIf($Text)
	{
		Line 1 "Network Interface(s)"
	}
	ElseIf($HTML)
	{
	}

	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration
	}
	
	Catch
	{
		$Results = $Null
	}

	If($? -and $Null -ne $Results)
	{
		$Nics = $Results | Where-Object {$Null -ne $_.ipaddress}
		$Results = $Null

		If($Null -eq $Nics) 
		{ 
			$GotNics = $False 
		} 
		Else 
		{ 
			$GotNics = !($Nics.__PROPERTY_COUNT -eq 0) 
		} 
	
		If($GotNics)
		{
			ForEach($nic in $nics)
			{
				Try
				{
					$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | Where-Object {$_.index -eq $nic.index}
				}
				
				Catch 
				{
					$ThisNic = $Null
				}
				
				If($? -and $Null -ne $ThisNic)
				{
					OutputNicItem $Nic $ThisNic
				}
				ElseIf(!$?)
				{
					Write-Warning "$(Get-Date): Error retrieving NIC information"
					Write-Host "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" -BackgroundColor Black -ForegroundColor Yellow
					Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "Error retrieving NIC information"
						Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
						Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
						Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
						Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "Error retrieving NIC information" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
						WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
				}
				Else
				{
					Write-Host "$(Get-Date): No results Returned for NIC information" -BackgroundColor Black -ForegroundColor Yellow
					If($MSWORD -or $PDF)
					{
						WriteWordLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
					ElseIf($Text)
					{
						Line 2 "No results Returned for NIC information"
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 2 "No results Returned for NIC information" "" $Null 0 $False $True
					}
				}
			}
		}	
	}
	ElseIf(!$?)
	{
		Write-Warning "$(Get-Date): Error retrieving NIC configuration information"
		Write-Host "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteWordLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteWordLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteWordLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteWordLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "Error retrieving NIC configuration information"
			Line 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
			Line 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository"
			Line 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may"
			Line 2 "need to rerun the script with Domain Admin credentials from the trusted Forest."
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "Error retrieving NIC configuration information" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "and winmgmt /salvagerepository.  If this is a trusted Forest, you may" "" $Null 0 $False $True
			WriteHTMLLine 0 2 "need to rerun the script with Domain Admin credentials from the trusted Forest." "" $Null 0 $False $True
		}
	}
	Else
	{
		Write-Host "$(Get-Date): No results Returned for NIC configuration information" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWORD -or $PDF)
		{
			WriteWordLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
		ElseIf($Text)
		{
			Line 2 "No results Returned for NIC configuration information"
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 2 "No results Returned for NIC configuration information" "" $Null 0 $False $True
		}
	}
	
	If($MSWORD -or $PDF)
	{
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 0 ""
	}
}

Function OutputComputerItem
{
	Param([object]$Item, [string]$OS)
	
	If($MSWord -or $PDF)
	{
		$ItemInformation = New-Object System.Collections.ArrayList
		$ItemInformation += (@{ Data = "Manufacturer"; Value = $Item.manufacturer; }) 
		$ItemInformation += (@{ Data = "Model"; Value = $Item.model; }) 
		$ItemInformation += (@{ Data = "Domain"; Value = $Item.domain; }) 
		$ItemInformation += (@{ Data = "Operating System"; Value = $OS; }) 
		$ItemInformation += (@{ Data = "Total Ram"; Value = "$($Item.totalphysicalram) GB"; }) 
		$ItemInformation += (@{ Data = "Physical Processors (sockets)"; Value = $Item.NumberOfProcessors; }) 
		$ItemInformation += (@{ Data = "Logical Processors (cores w/HT)"; Value = $Item.NumberOfLogicalProcessors; }) 
		$Table = AddWordTable -Hashtable $ItemInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustNone)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Manufacturer`t`t`t: " $Item.manufacturer
		Line 2 "Model`t`t`t`t: " $Item.model
		Line 2 "Domain`t`t`t`t: " $Item.domain
		Line 2 "Operating System`t`t: " $OS
		Line 2 "Total Ram`t`t`t: $($Item.totalphysicalram) GB"
		Line 2 "Physical Processors (sockets)`t: " $Item.NumberOfProcessors
		Line 2 "Logical Processors (cores w/HT)`t: " $Item.NumberOfLogicalProcessors
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Manufacturer",($htmlsilver -bor $htmlbold),$Item.manufacturer,$htmlwhite)
		$rowdata += @(,('Model',($htmlsilver -bor $htmlbold),$Item.model,$htmlwhite))
		$rowdata += @(,('Domain',($htmlsilver -bor $htmlbold),$Item.domain,$htmlwhite))
		$rowdata += @(,('Operating System',($htmlsilver -bor $htmlbold),$OS,$htmlwhite))
		$rowdata += @(,('Total Ram',($htmlsilver -bor $htmlbold),"$($Item.totalphysicalram) GB",$htmlwhite))
		$rowdata += @(,('Physical Processors (sockets)',($htmlsilver -bor $htmlbold),$Item.NumberOfProcessors,$htmlwhite))
		$rowdata += @(,('Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Item.NumberOfLogicalProcessors,$htmlwhite))

		$msg = "General Computer"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputDriveItem
{
	Param([object]$Drive)
	
	$xDriveType = ""
	Switch ($drive.drivetype)
	{
		0	{$xDriveType = "Unknown"; Break}
		1	{$xDriveType = "No Root Directory"; Break}
		2	{$xDriveType = "Removable Disk"; Break}
		3	{$xDriveType = "Local Disk"; Break}
		4	{$xDriveType = "Network Drive"; Break}
		5	{$xDriveType = "Compact Disc"; Break}
		6	{$xDriveType = "RAM Disk"; Break}
		Default {$xDriveType = "Unknown"; Break}
	}
	
	$xVolumeDirty = ""
	If(![String]::IsNullOrEmpty($drive.volumedirty))
	{
		If($drive.volumedirty)
		{
			$xVolumeDirty = "Yes"
		}
		Else
		{
			$xVolumeDirty = "No"
		}
	}

	If($MSWORD -or $PDF)
	{
		$DriveInformation = New-Object System.Collections.ArrayList
		$DriveInformation += (@{ Data = "Caption"; Value = $Drive.caption; }) 
		$DriveInformation += (@{ Data = "Size"; Value = "$($drive.drivesize) GB"; }) 
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$DriveInformation += (@{ Data = "File System"; Value = $Drive.filesystem; }) 
		}
		$DriveInformation += (@{ Data = "Free Space"; Value = "$($drive.drivefreespace) GB"; }) 
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$DriveInformation += (@{ Data = "Volume Name"; Value = $Drive.volumename; }) 
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$DriveInformation += (@{ Data = "Volume is Dirty"; Value = $xVolumeDirty; }) 
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$DriveInformation += (@{ Data = "Volume Serial Number"; Value = $Drive.volumeserialnumber; }) 
		}
		$DriveInformation += (@{ Data = "Drive Type"; Value = $xDriveType; }) 
		$Table = AddWordTable -Hashtable $DriveInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells `
		-Bold `
		-BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 2 ""
	}
	ElseIf($Text)
	{
		Line 2 "Caption`t`t: " $drive.caption
		Line 2 "Size`t`t: $($drive.drivesize) GB"
		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			Line 2 "File System`t: " $drive.filesystem
		}
		Line 2 "Free Space`t: $($drive.drivefreespace) GB"
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			Line 2 "Volume Name`t: " $drive.volumename
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			Line 2 "Volume is Dirty`t: " $xVolumeDirty
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			Line 2 "Volume Serial #`t: " $drive.volumeserialnumber
		}
		Line 2 "Drive Type`t: " $xDriveType
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Caption",($htmlsilver -bor $htmlbold),$Drive.caption,$htmlwhite)
		$rowdata += @(,('Size',($htmlsilver -bor $htmlbold),"$($drive.drivesize) GB",$htmlwhite))

		If(![String]::IsNullOrEmpty($drive.filesystem))
		{
			$rowdata += @(,('File System',($htmlsilver -bor $htmlbold),$Drive.filesystem,$htmlwhite))
		}
		$rowdata += @(,('Free Space',($htmlsilver -bor $htmlbold),"$($drive.drivefreespace) GB",$htmlwhite))
		If(![String]::IsNullOrEmpty($drive.volumename))
		{
			$rowdata += @(,('Volume Name',($htmlsilver -bor $htmlbold),$Drive.volumename,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumedirty))
		{
			$rowdata += @(,('Volume is Dirty',($htmlsilver -bor $htmlbold),$xVolumeDirty,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
		{
			$rowdata += @(,('Volume Serial Number',($htmlsilver -bor $htmlbold),$Drive.volumeserialnumber,$htmlwhite))
		}
		$rowdata += @(,('Drive Type',($htmlsilver -bor $htmlbold),$xDriveType,$htmlwhite))

		$msg = "Drive(s)"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputProcessorItem
{
	Param([object]$Processor)
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"; Break}
		2	{$xAvailability = "Unknown"; Break}
		3	{$xAvailability = "Running or Full Power"; Break}
		4	{$xAvailability = "Warning"; Break}
		5	{$xAvailability = "In Test"; Break}
		6	{$xAvailability = "Not Applicable"; Break}
		7	{$xAvailability = "Power Off"; Break}
		8	{$xAvailability = "Off Line"; Break}
		9	{$xAvailability = "Off Duty"; Break}
		10	{$xAvailability = "Degraded"; Break}
		11	{$xAvailability = "Not Installed"; Break}
		12	{$xAvailability = "Install Error"; Break}
		13	{$xAvailability = "Power Save - Unknown"; Break}
		14	{$xAvailability = "Power Save - Low Power Mode"; Break}
		15	{$xAvailability = "Power Save - Standby"; Break}
		16	{$xAvailability = "Power Cycle"; Break}
		17	{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	If($MSWORD -or $PDF)
	{
		$ProcessorInformation = New-Object System.Collections.ArrayList
		$ProcessorInformation += (@{ Data = "Name"; Value = $Processor.name; }) 
		$ProcessorInformation += (@{ Data = "Description"; Value = $Processor.description; }) 
		$ProcessorInformation += (@{ Data = "Max Clock Speed"; Value = "$($processor.maxclockspeed) MHz"; }) 
		If($processor.l2cachesize -gt 0)
		{
			$ProcessorInformation += (@{ Data = "L2 Cache Size"; Value = "$($processor.l2cachesize) KB"; }) 
		}
		If($processor.l3cachesize -gt 0)
		{
			$ProcessorInformation += (@{ Data = "L3 Cache Size"; Value = "$($processor.l3cachesize) KB"; }) 
		}
		If($processor.numberofcores -gt 0)
		{
			$ProcessorInformation += (@{ Data = "Number of Cores"; Value = $Processor.numberofcores; }) 
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$ProcessorInformation += (@{ Data = "Number of Logical Processors (cores w/HT)"; Value = $Processor.numberoflogicalprocessors; }) 
		}
		$ProcessorInformation += (@{ Data = "Availability"; Value = $xAvailability; }) 
		$Table = AddWordTable -Hashtable $ProcessorInformation `
		-Columns Data,Value `
		-List `
		-AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t`t: " $processor.name
		Line 2 "Description`t`t`t: " $processor.description
		Line 2 "Max Clock Speed`t`t`t: $($processor.maxclockspeed) MHz"
		If($processor.l2cachesize -gt 0)
		{
			Line 2 "L2 Cache Size`t`t`t: $($processor.l2cachesize) KB"
		}
		If($processor.l3cachesize -gt 0)
		{
			Line 2 "L3 Cache Size`t`t`t: $($processor.l3cachesize) KB"
		}
		If($processor.numberofcores -gt 0)
		{
			Line 2 "# of Cores`t`t`t: " $processor.numberofcores
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			Line 2 "# of Logical Procs (cores w/HT)`t: " $processor.numberoflogicalprocessors
		}
		Line 2 "Availability`t`t`t: " $xAvailability
		Line 2 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Processor.name,$htmlwhite)
		$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Processor.description,$htmlwhite))

		$rowdata += @(,('Max Clock Speed',($htmlsilver -bor $htmlbold),"$($processor.maxclockspeed) MHz",$htmlwhite))
		If($processor.l2cachesize -gt 0)
		{
			$rowdata += @(,('L2 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l2cachesize) KB",$htmlwhite))
		}
		If($processor.l3cachesize -gt 0)
		{
			$rowdata += @(,('L3 Cache Size',($htmlsilver -bor $htmlbold),"$($processor.l3cachesize) KB",$htmlwhite))
		}
		If($processor.numberofcores -gt 0)
		{
			$rowdata += @(,('Number of Cores',($htmlsilver -bor $htmlbold),$Processor.numberofcores,$htmlwhite))
		}
		If($processor.numberoflogicalprocessors -gt 0)
		{
			$rowdata += @(,('Number of Logical Processors (cores w/HT)',($htmlsilver -bor $htmlbold),$Processor.numberoflogicalprocessors,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))

		$msg = "Processor(s)"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputNicItem
{
	Param([object]$Nic, [object]$ThisNic)
	
	$powerMgmt = Get-WmiObject MSPower_DeviceEnable -Namespace root\wmi | Where-Object {$_.InstanceName -match [regex]::Escape($ThisNic.PNPDeviceID)}

	If($? -and $Null -ne $powerMgmt)
	{
		If($powerMgmt.Enable -eq $True)
		{
			$PowerSaving = "Enabled"
		}
		Else
		{
			$PowerSaving = "Disabled"
		}
	}
	Else
	{
        $PowerSaving = "N/A"
	}
	
	$xAvailability = ""
	Switch ($processor.availability)
	{
		1	{$xAvailability = "Other"; Break}
		2	{$xAvailability = "Unknown"; Break}
		3	{$xAvailability = "Running or Full Power"; Break}
		4	{$xAvailability = "Warning"; Break}
		5	{$xAvailability = "In Test"; Break}
		6	{$xAvailability = "Not Applicable"; Break}
		7	{$xAvailability = "Power Off"; Break}
		8	{$xAvailability = "Off Line"; Break}
		9	{$xAvailability = "Off Duty"; Break}
		10	{$xAvailability = "Degraded"; Break}
		11	{$xAvailability = "Not Installed"; Break}
		12	{$xAvailability = "Install Error"; Break}
		13	{$xAvailability = "Power Save - Unknown"; Break}
		14	{$xAvailability = "Power Save - Low Power Mode"; Break}
		15	{$xAvailability = "Power Save - Standby"; Break}
		16	{$xAvailability = "Power Cycle"; Break}
		17	{$xAvailability = "Power Save - Warning"; Break}
		Default	{$xAvailability = "Unknown"; Break}
	}

	$xIPAddress = @()
	ForEach($IPAddress in $Nic.ipaddress)
	{
		$xIPAddress += "$($IPAddress)"
	}

	$xIPSubnet = @()
	ForEach($IPSubnet in $Nic.ipsubnet)
	{
		$xIPSubnet += "$($IPSubnet)"
	}

	If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
	{
		$nicdnsdomainsuffixsearchorder = $nic.dnsdomainsuffixsearchorder
		$xnicdnsdomainsuffixsearchorder = New-Object System.Collections.ArrayList
		ForEach($DNSDomain in $nicdnsdomainsuffixsearchorder)
		{
			$xnicdnsdomainsuffixsearchorder += ("$($DNSDomain)") 
		}
	}
	
	If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
	{
		$nicdnsserversearchorder = $nic.dnsserversearchorder
		$xnicdnsserversearchorder = @()
		ForEach($DNSServer in $nicdnsserversearchorder)
		{
			$xnicdnsserversearchorder += "$($DNSServer)"
		}
	}

	$xdnsenabledforwinsresolution = ""
	If($nic.dnsenabledforwinsresolution)
	{
		$xdnsenabledforwinsresolution = "Yes"
	}
	Else
	{
		$xdnsenabledforwinsresolution = "No"
	}
	
	$xTcpipNetbiosOptions = ""
	Switch ($nic.TcpipNetbiosOptions)
	{
		0	{$xTcpipNetbiosOptions = "Use NetBIOS setting from DHCP Server"; Break}
		1	{$xTcpipNetbiosOptions = "Enable NetBIOS"; Break}
		2	{$xTcpipNetbiosOptions = "Disable NetBIOS"; Break}
		Default	{$xTcpipNetbiosOptions = "Unknown"; Break}
	}
	
	$xwinsenablelmhostslookup = ""
	If($nic.winsenablelmhostslookup)
	{
		$xwinsenablelmhostslookup = "Yes"
	}
	Else
	{
		$xwinsenablelmhostslookup = "No"
	}

	If($MSWORD -or $PDF)
	{
		$NicInformation = New-Object System.Collections.ArrayList
		$NicInformation += (@{ Data = "Name"; Value = $ThisNic.Name; }) 
		If($ThisNic.Name -ne $nic.description)
		{
			$NicInformation += (@{ Data = "Description"; Value = $Nic.description; }) 
		}
		$NicInformation += (@{ Data = "Connection ID"; Value = $ThisNic.NetConnectionID; }) 
		If(validObject $Nic Manufacturer)
		{
			$NicInformation += (@{ Data = "Manufacturer"; Value = $Nic.manufacturer; }) 
		}
		$NicInformation += (@{ Data = "Availability"; Value = $xAvailability; }) 
		$NicInformation += (@{ Data = "Allow the computer to turn off this device to save power"; Value = $PowerSaving; }) 
		$NicInformation += (@{ Data = "Physical Address"; Value = $Nic.macaddress; }) 
		If($xIPAddress.Count -gt 1)
		{
			$NicInformation += (@{ Data = "IP Address"; Value = $xIPAddress[0]; }) 
			$NicInformation += (@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) 
			$NicInformation += (@{ Data = "Subnet Mask"; Value = $xIPSubnet[0]; }) 
			$cnt = -1
			ForEach($tmp in $xIPAddress)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += (@{ Data = "IP Address"; Value = $tmp; }) 
					$NicInformation += (@{ Data = "Subnet Mask"; Value = $xIPSubnet[$cnt]; }) 
				}
			}
		}
		Else
		{
			$NicInformation += (@{ Data = "IP Address"; Value = $xIPAddress; }) 
			$NicInformation += (@{ Data = "Default Gateway"; Value = $Nic.Defaultipgateway; }) 
			$NicInformation += (@{ Data = "Subnet Mask"; Value = $xIPSubnet; }) 
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$NicInformation += (@{ Data = "DHCP Enabled"; Value = $Nic.dhcpenabled; }) 
			$NicInformation += (@{ Data = "DHCP Lease Obtained"; Value = $dhcpleaseobtaineddate; }) 
			$NicInformation += (@{ Data = "DHCP Lease Expires"; Value = $dhcpleaseexpiresdate; }) 
			$NicInformation += (@{ Data = "DHCP Server"; Value = $Nic.dhcpserver; }) 
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$NicInformation += (@{ Data = "DNS Domain"; Value = $Nic.dnsdomain; }) 
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$NicInformation += (@{ Data = "DNS Search Suffixes"; Value = $xnicdnsdomainsuffixsearchorder[0]; }) 
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += (@{ Data = ""; Value = $tmp; }) 
				}
			}
		}
		$NicInformation += (@{ Data = "DNS WINS Enabled"; Value = $xdnsenabledforwinsresolution; }) 
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$NicInformation += (@{ Data = "DNS Servers"; Value = $xnicdnsserversearchorder[0]; }) 
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$NicInformation += (@{ Data = ""; Value = $tmp; }) 
				}
			}
		}
		$NicInformation += (@{ Data = "NetBIOS Setting"; Value = $xTcpipNetbiosOptions; }) 
		$NicInformation += (@{ Data = "WINS: Enabled LMHosts"; Value = $xwinsenablelmhostslookup; }) 
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$NicInformation += (@{ Data = "Host Lookup File"; Value = $Nic.winshostlookupfile; }) 
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$NicInformation += (@{ Data = "Primary Server"; Value = $Nic.winsprimaryserver; }) 
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$NicInformation += (@{ Data = "Secondary Server"; Value = $Nic.winssecondaryserver; }) 
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$NicInformation += (@{ Data = "Scope ID"; Value = $Nic.winsscopeid; }) 
		}
		$Table = AddWordTable -Hashtable $NicInformation -Columns Data,Value -List -AutoFit $wdAutoFitFixed;

		## Set first column format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 150;
		$Table.Columns.Item(2).Width = 200;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 2 "Name`t`t`t: " $ThisNic.Name
		If($ThisNic.Name -ne $nic.description)
		{
			Line 2 "Description`t`t: " $nic.description
		}
		Line 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
		If(validObject $Nic Manufacturer)
		{
			Line 2 "Manufacturer`t`t: " $Nic.manufacturer
		}
		Line 2 "Availability`t`t: " $xAvailability
		Line 2 "Allow computer to turn "
		Line 2 "off device to save power: " $PowerSaving
		Line 2 "Physical Address`t: " $nic.macaddress
		Line 2 "IP Address`t`t: " $xIPAddress[0]
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		Line 2 "Default Gateway`t`t: " $Nic.Defaultipgateway
		Line 2 "Subnet Mask`t`t: " $xIPSubnet[0]
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				Line 5 "  " $tmp
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			Line 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
			Line 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
			Line 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
			Line 2 "DHCP Server`t`t:" $nic.dhcpserver
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			Line 2 "DNS Domain`t`t: " $nic.dnsdomain
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Search Suffixes`t: " $xnicdnsdomainsuffixsearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "DNS WINS Enabled`t: " $xdnsenabledforwinsresolution
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			[int]$x = 1
			Line 2 "DNS Servers`t`t: " $xnicdnsserversearchorder[0]
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 5 "  " $tmp
				}
			}
		}
		Line 2 "NetBIOS Setting`t`t: " $xTcpipNetbiosOptions
		Line 2 "WINS:"
		Line 3 "Enabled LMHosts`t: " $xwinsenablelmhostslookup
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			Line 3 "Host Lookup File`t: " $nic.winshostlookupfile
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			Line 3 "Primary Server`t: " $nic.winsprimaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			Line 3 "Secondary Server`t: " $nic.winssecondaryserver
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			Line 3 "Scope ID`t`t: " $nic.winsscopeid
		}
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$rowdata = @()
		$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$ThisNic.Name,$htmlwhite)
		If($ThisNic.Name -ne $nic.description)
		{
			$rowdata += @(,('Description',($htmlsilver -bor $htmlbold),$Nic.description,$htmlwhite))
		}
		$rowdata += @(,('Connection ID',($htmlsilver -bor $htmlbold),$ThisNic.NetConnectionID,$htmlwhite))
		If(validObject $Nic Manufacturer)
		{
			$rowdata += @(,('Manufacturer',($htmlsilver -bor $htmlbold),$Nic.manufacturer,$htmlwhite))
		}
		$rowdata += @(,('Availability',($htmlsilver -bor $htmlbold),$xAvailability,$htmlwhite))
		$rowdata += @(,('Allow the computer to turn off this device to save power',($htmlsilver -bor $htmlbold),$PowerSaving,$htmlwhite))
		$rowdata += @(,('Physical Address',($htmlsilver -bor $htmlbold),$Nic.macaddress,$htmlwhite))
		$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$xIPAddress[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPAddress)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('IP Address',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		$rowdata += @(,('Default Gateway',($htmlsilver -bor $htmlbold),$Nic.Defaultipgateway[0],$htmlwhite))
		$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$xIPSubnet[0],$htmlwhite))
		$cnt = -1
		ForEach($tmp in $xIPSubnet)
		{
			$cnt++
			If($cnt -gt 0)
			{
				$rowdata += @(,('Subnet Mask',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
			}
		}
		If($nic.dhcpenabled)
		{
			$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
			$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
			$rowdata += @(,('DHCP Enabled',($htmlsilver -bor $htmlbold),$Nic.dhcpenabled,$htmlwhite))
			$rowdata += @(,('DHCP Lease Obtained',($htmlsilver -bor $htmlbold),$dhcpleaseobtaineddate,$htmlwhite))
			$rowdata += @(,('DHCP Lease Expires',($htmlsilver -bor $htmlbold),$dhcpleaseexpiresdate,$htmlwhite))
			$rowdata += @(,('DHCP Server',($htmlsilver -bor $htmlbold),$Nic.dhcpserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.dnsdomain))
		{
			$rowdata += @(,('DNS Domain',($htmlsilver -bor $htmlbold),$Nic.dnsdomain,$htmlwhite))
		}
		If($Null -ne $nic.dnsdomainsuffixsearchorder -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Search Suffixes',($htmlsilver -bor $htmlbold),$xnicdnsdomainsuffixsearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsdomainsuffixsearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('DNS WINS Enabled',($htmlsilver -bor $htmlbold),$xdnsenabledforwinsresolution,$htmlwhite))
		If($Null -ne $nic.dnsserversearchorder -and $nic.dnsserversearchorder.length -gt 0)
		{
			$rowdata += @(,('DNS Servers',($htmlsilver -bor $htmlbold),$xnicdnsserversearchorder[0],$htmlwhite))
			$cnt = -1
			ForEach($tmp in $xnicdnsserversearchorder)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
			}
		}
		$rowdata += @(,('NetBIOS Setting',($htmlsilver -bor $htmlbold),$xTcpipNetbiosOptions,$htmlwhite))
		$rowdata += @(,('WINS: Enabled LMHosts',($htmlsilver -bor $htmlbold),$xwinsenablelmhostslookup,$htmlwhite))
		If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
		{
			$rowdata += @(,('Host Lookup File',($htmlsilver -bor $htmlbold),$Nic.winshostlookupfile,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
		{
			$rowdata += @(,('Primary Server',($htmlsilver -bor $htmlbold),$Nic.winsprimaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
		{
			$rowdata += @(,('Secondary Server',($htmlsilver -bor $htmlbold),$Nic.winssecondaryserver,$htmlwhite))
		}
		If(![String]::IsNullOrEmpty($nic.winsscopeid))
		{
			$rowdata += @(,('Scope ID',($htmlsilver -bor $htmlbold),$Nic.winsscopeid,$htmlwhite))
		}

		$msg = "Network Interface(s)"
		$columnWidths = @("150px","200px")
		FormatHTMLTable $msg -rowarray $rowdata -columnArray $columnheaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region word specific functions
Function SetWordHashTable
{
	Param([string]$CultureCode)

	#optimized by Michael B. Smith
	
	# DE and FR translations for Word 2010 by Vladimir Radojevic
	# Vladimir.Radojevic@Commerzreal.com

	# DA translations for Word 2010 by Thomas Daugaard
	# Citrix Infrastructure Specialist at edgemo A/S

	# CA translations by Javier Sanchez 
	# CEO & Founder 101 Consulting

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese
	
	[string]$toc = $(
		Switch ($CultureCode)
		{
			'ca-'	{ 'Taula automática 2'; Break }
			'da-'	{ 'Automatisk tabel 2'; Break }
			'de-'	{ 'Automatische Tabelle 2'; Break }
			'en-'	{ 'Automatic Table 2'; Break }
			'es-'	{ 'Tabla automática 2'; Break }
			'fi-'	{ 'Automaattinen taulukko 2'; Break }
#			'fr-'	{ 'Sommaire Automatique 2'; Break }
			'fr-'	{ 'Table automatique 2'; Break } #changed 10-feb-2017 david roquier and samuel legrand
			'nb-'	{ 'Automatisk tabell 2'; Break }
			'nl-'	{ 'Automatische inhoudsopgave 2'; Break }
			'pt-'	{ 'Sumário Automático 2'; Break }
			'sv-'	{ 'Automatisk innehållsförteckning2'; Break }
			'zh-'	{ '自动目录 2'; Break }
		}
	)

	$Script:myHash                      = @{}
	$Script:myHash.Word_TableOfContents = $toc
	$Script:myHash.Word_NoSpacing       = $wdStyleNoSpacing
	$Script:myHash.Word_Heading1        = $wdStyleheading1
	$Script:myHash.Word_Heading2        = $wdStyleheading2
	$Script:myHash.Word_Heading3        = $wdStyleheading3
	$Script:myHash.Word_Heading4        = $wdStyleheading4
	$Script:myHash.Word_TableGrid       = $wdTableGrid
}

Function GetCulture
{
	Param([int]$WordValue)
	
	#codes obtained from http://support.microsoft.com/kb/221435
	#http://msdn.microsoft.com/en-us/library/bb213877(v=office.12).aspx
	$CatalanArray = 1027
	$ChineseArray = 2052,3076,5124,4100
	$DanishArray = 1030
	$DutchArray = 2067, 1043
	$EnglishArray = 3081, 10249, 4105, 9225, 6153, 8201, 5129, 13321, 7177, 11273, 2057, 1033, 12297
	$FinnishArray = 1035
	$FrenchArray = 2060, 1036, 11276, 3084, 12300, 5132, 13324, 6156, 8204, 10252, 7180, 9228, 4108
	$GermanArray = 1031, 3079, 5127, 4103, 2055
	$NorwegianArray = 1044, 2068
	$PortugueseArray = 1046, 2070
	$SpanishArray = 1034, 11274, 16394, 13322, 9226, 5130, 7178, 12298, 17418, 4106, 18442, 19466, 6154, 15370, 10250, 20490, 3082, 14346, 8202
	$SwedishArray = 1053, 2077

	#ca - Catalan
	#da - Danish
	#de - German
	#en - English
	#es - Spanish
	#fi - Finnish
	#fr - French
	#nb - Norwegian
	#nl - Dutch
	#pt - Portuguese
	#sv - Swedish
	#zh - Chinese

	Switch ($WordValue)
	{
		{$CatalanArray -contains $_}	{$CultureCode = "ca-"}
		{$ChineseArray -contains $_}	{$CultureCode = "zh-"}
		{$DanishArray -contains $_}		{$CultureCode = "da-"}
		{$DutchArray -contains $_}		{$CultureCode = "nl-"}
		{$EnglishArray -contains $_}	{$CultureCode = "en-"}
		{$FinnishArray -contains $_}	{$CultureCode = "fi-"}
		{$FrenchArray -contains $_}		{$CultureCode = "fr-"}
		{$GermanArray -contains $_}		{$CultureCode = "de-"}
		{$NorwegianArray -contains $_}	{$CultureCode = "nb-"}
		{$PortugueseArray -contains $_}	{$CultureCode = "pt-"}
		{$SpanishArray -contains $_}	{$CultureCode = "es-"}
		{$SwedishArray -contains $_}	{$CultureCode = "sv-"}
		Default {$CultureCode = "en-"}
	}
	
	Return $CultureCode
}

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP, [string]$CultureCode)
	
	$xArray = ""
	
	Switch ($CultureCode)
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Austin", "En bandes", "Faceta", "Filigrana",
					"Integral", "Ió (clar)", "Ió (fosc)", "Línia lateral",
					"Moviment", "Quadrícula", "Retrospectiu", "Sector (clar)",
					"Sector (fosc)", "Semàfor", "Visualització", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Anual", "Austin", "Conservador",
					"Contrast", "Cubicles", "Diplomàtic", "Exposició",
					"Línia lateral", "Mod", "Mosiac", "Moviment", "Paper de diari",
					"Perspectiva", "Piles", "Quadrícula", "Sobri",
					"Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevægElse", "Brusen", "Facet", "Filigran", 
					"Gitter", "Integral", "Ion (lys)", "Ion (mørk)", 
					"Retro", "Semafor", "Sidelinje", "Stribet", 
					"Udsnit (lys)", "Udsnit (mørk)", "Visningsmaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("BevægElse", "Brusen", "Ion (lys)", "Filigran",
					"Retro", "Semafor", "Visningsmaster", "Integral",
					"Facet", "Gitter", "Stribet", "Sidelinje", "Udsnit (lys)",
					"Udsnit (mørk)", "Ion (mørk)", "Austin")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("BevægElse", "Moderat", "Perspektiv", "Firkanter",
					"Overskrid", "Alfabet", "Kontrast", "Stakke", "Fliser", "Gåde",
					"Gitter", "Austin", "Eksponering", "Sidelinje", "Enkel",
					"Nålestribet", "Årlig", "Avispapir", "Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Bewegung", "Facette", "Filigran", 
					"Gebändert", "Integral", "Ion (dunkel)", "Ion (hell)", 
					"Pfiff", "Randlinie", "Raster", "Rückblick", 
					"Segment (dunkel)", "Segment (hell)", "Semaphor", 
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Semaphor", "Segment (hell)", "Ion (hell)",
					"Raster", "Ion (dunkel)", "Filigran", "Rückblick", "Pfiff",
					"ViewMaster", "Segment (dunkel)", "Verbunden", "Bewegung",
					"Randlinie", "Austin", "Integral", "Facette")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Austin", "Bewegung", "Durchscheinend",
					"Herausgestellt", "Jährlich", "Kacheln", "Kontrast", "Kubistisch",
					"Modern", "Nadelstreifen", "Perspektive", "Puzzle", "Randlinie",
					"Raster", "Schlicht", "Stapel", "Traditionell", "Zeitungspapier")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
					"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
					"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
					"Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
					"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
					"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Con bandas", "Cortar (oscuro)", "Cuadrícula", 
					"Whisp", "Faceta", "Filigrana", "Integral", "Ion (claro)", 
					"Ion (oscuro)", "Línea lateral", "Movimiento", "Retrospectiva", 
					"Semáforo", "Slice (luz)", "Vista principal", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Whisp", "Vista principal", "Filigrana", "Austin",
					"Slice (luz)", "Faceta", "Semáforo", "Retrospectiva", "Cuadrícula",
					"Movimiento", "Cortar (oscuro)", "Línea lateral", "Ion (oscuro)",
					"Ion (claro)", "Integral", "Con bandas")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Austin", "Conservador",
					"Contraste", "Cuadrícula", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Papel periódico",
					"Perspectiva", "Pilas", "Puzzle", "Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kuiskaus", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("Filigraani", "Integraali", "Ioni (tumma)",
					"Ioni (vaalea)", "Opastin", "Pinta", "Retro", "Sektori (tumma)",
					"Sektori (vaalea)", "Vaihtuvavärinen", "ViewMaster", "Austin",
					"Kiehkura", "Liike", "Ruudukko", "Sivussa")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aakkoset", "Askeettinen", "Austin", "Kontrasti",
					"Laatikot", "Liike", "Liituraita", "Mod", "Osittain peitossa",
					"Palapeli", "Perinteinen", "Perspektiivi", "Pinot", "Ruudukko",
					"Ruudut", "Sanomalehtipaperi", "Sivussa", "Vuotuinen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("À bandes", "Austin", "Facette", "Filigrane", 
					"Guide", "Intégrale", "Ion (clair)", "Ion (foncé)", 
					"Lignes latérales", "Quadrillage", "Rétrospective", "Secteur (clair)", 
					"Secteur (foncé)", "Sémaphore", "ViewMaster", "Whisp")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Austin", 
					"Blocs empilés", "Classique", "Contraste", "Emplacements de bureau", 
					"Exposition", "Guide", "Ligne latérale", "Moderne", 
					"Mosaïques", "Mots croisés", "Papier journal", "Perspective",
					"Quadrillage", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "BevegElse", "Dempet", "Fasett", "Filigran",
					"Integral", "Ion (lys)", "Ion (mørk)", "Retrospekt", "Rutenett",
					"Sektor (lys)", "Sektor (mørk)", "Semafor", "Sidelinje", "Stripet",
					"ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabet", "Årlig", "Avistrykk", "Austin", "Avlukker",
					"BevegElse", "Engasjement", "Enkel", "Fliser", "Konservativ",
					"Kontrast", "Mod", "Perspektiv", "Puslespill", "Rutenett", "Sidelinje",
					"Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Beweging", "Facet", "Filigraan", "Gestreept",
					"Integraal", "Ion (donker)", "Ion (licht)", "Raster",
					"Segment (Light)", "Semafoor", "Slice (donker)", "Spriet",
					"Terugblik", "Terzijde", "ViewMaster")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Austin", "Bescheiden",
					"Beweging", "Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks",
					"Krantenpapier", "Krijtstreep", "Kubussen", "Mod", "Perspectief",
					"Puzzel", "Raster", "Stapels",
					"Tegels", "Terzijde")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Animação", "Austin", "Em Tiras", "Exibição Mestra",
					"Faceta", "Fatia (Clara)", "Fatia (Escura)", "Filete", "Filigrana", 
					"Grade", "Integral", "Íon (Claro)", "Íon (Escuro)", "Linha Lateral",
					"Retrospectiva", "Semáforo")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Austin", "Baias",
					"Conservador", "Contraste", "Exposição", "Grade", "Ladrilhos",
					"Linha Lateral", "Listras", "Mod", "Papel Jornal", "Perspectiva", "Pilhas",
					"Quebra-cabeça", "Transcend")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ("Austin", "Band", "Fasett", "Filigran", "Integrerad", "Jon (ljust)",
					"Jon (mörkt)", "Knippe", "Rutnät", "RörElse", "Sektor (ljus)", "Sektor (mörk)",
					"Semafor", "Sidlinje", "VisaHuvudsida", "Återblick")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Alfabetmönster", "Austin", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Perspektiv", "Plattor", "Pussel", "Rutnät",
					"RörElse", "Sidlinje", "Sobert", "Staplat", "Tidningspapper", "Årligt",
					"Övergående")
				}
			}

		'zh-'	{
				If($xWordVersion -eq $wdWord2010 -or $xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
				{
					$xArray = ('奥斯汀', '边线型', '花丝', '怀旧', '积分',
					'离子(浅色)', '离子(深色)', '母版型', '平面', '切片(浅色)',
					'切片(深色)', '丝状', '网格', '镶边', '信号灯',
					'运动型')
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013 -or $xWordVersion -eq $wdWord2016)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid",
						"Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect",
						"Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster",
						"Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		$xArray = $Null
		Return $True
	}
	Else
	{
		$xArray = $Null
		Return $False
	}
}

Function ConvertTo-ScriptBlock 
{
	#by Jeff Wouters, PowerShell MVP
	Param([string]$string)
	$ScriptBlock = $executioncontext.invokecommand.NewScriptBlock($string)
	Return $ScriptBlock
} 

Function SWExclusions 
{
	# original work by Shaun Ritchie
	# performance improvements by Jeff Wouters, PowerShell MVP
	# modified by Webster
	# modified 3-jan-2014 to add displayversion
	# bug found 30-jul-2014 by Sam Jacobs
	# this function did not work if the SoftwareExlusions.txt file contained only one line
	$var = ""
	$Tmp = '$InstalledApps | Where {'
	$Exclusions = Get-Content "$($pwd.path)\SoftwareExclusions.txt" -EA 0
	If($? -and $Null -ne $Exclusions)
	{
		If($Exclusions -is [array])	
		{
			ForEach($Exclusion in $Exclusions) 
			{
				$Tmp += "(`$`_.DisplayName -notlike ""$($Exclusion)"") -and "
			}
			$var += $Tmp.Substring(0,($Tmp.Length - 6))
			}
		Else
		{
			# added 30-jul-2014 to handle if the file contained only one line
			$tmp += "(`$`_.DisplayName -notlike ""$($Exclusions)"")"
			$var = $tmp
		}
		$var += "} | Select-Object DisplayName, DisplayVersion | Sort DisplayName -unique"
	}
	return $var
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`t`tThis script directly outputs to Microsoft Word, please install Microsoft Word`n`n" -BackgroundColor Black -ForegroundColor Yellow
		Exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "`n`n`tPlease close all instances of Microsoft Word before running this report.`n`n" -BackgroundColor Black -ForegroundColor Yellow
		Exit
	}
}

Function ValidateCompanyName
{
	[bool]$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	If($xResult)
	{
		Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
	}
	Else
	{
		$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		If($xResult)
		{
			Return Get-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "Company"
		}
		Else
		{
			Return ""
		}
	}
}

Function Set-DocumentProperty {
    <#
	.SYNOPSIS
	Function to set the Title Page document properties in MS Word
	.DESCRIPTION
	Long description
	.PARAMETER Document
	Current Document Object
	.PARAMETER DocProperty
	Parameter description
	.PARAMETER Value
	Parameter description
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value 'MyTitle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value 'MyCompany'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value 'Jim Moyle'
	.EXAMPLE
	Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value 'MySubjectTitle'
	.NOTES
	Function Created by Jim Moyle June 2017
	Twitter : @JimMoyle
	#>
    param (
        [object]$Document,
        [String]$DocProperty,
        [string]$Value
    )
    try {
        $binding = "System.Reflection.BindingFlags" -as [type]
        $builtInProperties = $Document.BuiltInDocumentProperties
        $property = [System.__ComObject].invokemember("item", $binding::GetProperty, $null, $BuiltinProperties, $DocProperty)
        [System.__ComObject].invokemember("value", $binding::SetProperty, $null, $property, $Value)
    }
    catch {
        Write-Warning "Failed to set $DocProperty to $Value"
    }
}

Function FindWordDocumentEnd
{
	#Return focus to main document    
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument
	#move to the end of the current document
	$Script:Selection.EndKey($wdStory,$wdMove) | Out-Null
}

Function SetupWord
{
	Write-Host "$(Get-Date): Setting up Word" -BackgroundColor Black -ForegroundColor Yellow
    
	# Setup word for output
	Write-Host "$(Get-Date): Create Word comObject." -BackgroundColor Black -ForegroundColor Yellow
	$Script:Word = New-Object -comobject "Word.Application" -EA 0
	
	If(!$? -or $Null -eq $Script:Word)
	{
		Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tThe Word object could not be created.  You may need to repair your Word installation.`n`n`t`tScript cannot continue.`n`n"
		Exit
	}

	Write-Host "$(Get-Date): Determine Word language value" -BackgroundColor Black -ForegroundColor Yellow
	If( ( validStateProp $Script:Word Language Value__ ) )
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language.Value__
	}
	Else
	{
		[int]$Script:WordLanguageValue = [int]$Script:Word.Language
	}

	If(!($Script:WordLanguageValue -gt -1))
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tUnable to determine the Word language value.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}
	Write-Host "$(Get-Date): Word language value is $($Script:WordLanguageValue)" -BackgroundColor Black -ForegroundColor Yellow
	
	$Script:WordCultureCode = GetCulture $Script:WordLanguageValue
	
	SetWordHashTable $Script:WordCultureCode
	
	[int]$Script:WordVersion = [int]$Script:Word.Version
	If($Script:WordVersion -eq $wdWord2016)
	{
		$Script:WordProduct = "Word 2016"
	}
	ElseIf($Script:WordVersion -eq $wdWord2013)
	{
		$Script:WordProduct = "Word 2013"
	}
	ElseIf($Script:WordVersion -eq $wdWord2010)
	{
		$Script:WordProduct = "Word 2010"
	}
	ElseIf($Script:WordVersion -eq $wdWord2007)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tMicrosoft Word 2007 is no longer supported.`n`n`t`tScript will end.`n`n"
		AbortScript
	}
	Else
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tYou are running an untested or unsupported version of Microsoft Word.`n`n`t`tScript will end.`n`n`t`tPlease send info on your version of Word to webster@carlwebster.com`n`n"
		AbortScript
	}

	#only validate CompanyName if the field is blank
	If([String]::IsNullOrEmpty($Script:CoName))
	{
		Write-Host "$(Get-Date): Company name is blank.  Retrieve company name from registry." -BackgroundColor Black -ForegroundColor Yellow
		$TmpName = ValidateCompanyName
		
		If([String]::IsNullOrEmpty($TmpName))
		{
			Write-Warning "`n`n`t`tCompany Name is blank so Cover Page will not show a Company Name."
			Write-Warning "`n`t`tCheck HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
			Write-Warning "`n`t`tYou may want to use the -CompanyName parameter if you need a Company Name on the cover page.`n`n"
		}
		Else
		{
			$Script:CoName = $TmpName
			Write-Host "$(Get-Date): Updated company name to $($Script:CoName)" -BackgroundColor Black -ForegroundColor Yellow
		}
	}

	If($Script:WordCultureCode -ne "en-")
	{
		Write-Host "$(Get-Date): Check Default Cover Page for $($WordCultureCode)" -BackgroundColor Black -ForegroundColor Yellow
		[bool]$CPChanged = $False
		Switch ($Script:WordCultureCode)
		{
			'ca-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línia lateral"
						$CPChanged = $True
					}
				}

			'da-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'de-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Randlinie"
						$CPChanged = $True
					}
				}

			'es-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Línea lateral"
						$CPChanged = $True
					}
				}

			'fi-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sivussa"
						$CPChanged = $True
					}
				}

			'fr-'	{
					If($CoverPage -eq "Sideline")
					{
						If($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
						{
							$CoverPage = "Lignes latérales"
							$CPChanged = $True
						}
						Else
						{
							$CoverPage = "Ligne latérale"
							$CPChanged = $True
						}
					}
				}

			'nb-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidelinje"
						$CPChanged = $True
					}
				}

			'nl-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Terzijde"
						$CPChanged = $True
					}
				}

			'pt-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Linha Lateral"
						$CPChanged = $True
					}
				}

			'sv-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "Sidlinje"
						$CPChanged = $True
					}
				}

			'zh-'	{
					If($CoverPage -eq "Sideline")
					{
						$CoverPage = "边线型"
						$CPChanged = $True
					}
				}
		}

		If($CPChanged)
		{
			Write-Host "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)" -BackgroundColor Black -ForegroundColor Yellow
		}
	}

	Write-Host "$(Get-Date): Validate cover page $($CoverPage) for culture code $($Script:WordCultureCode)" -BackgroundColor Black -ForegroundColor Yellow
	[bool]$ValidCP = $False
	
	$ValidCP = ValidateCoverPage $Script:WordVersion $CoverPage $Script:WordCultureCode
	
	If(!$ValidCP)
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Host "$(Get-Date): Word language value $($Script:WordLanguageValue)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): Culture code $($Script:WordCultureCode)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Error "`n`n`t`tFor $($Script:WordProduct), $($CoverPage) is not a valid Cover Page option.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	ShowScriptOptions

	$Script:Word.Visible = $False

	#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
	#using Jeff's Demo-WordReport.ps1 file for examples
	Write-Host "$(Get-Date): Load Word Templates" -BackgroundColor Black -ForegroundColor Yellow

	[bool]$Script:CoverPagesExist = $False
	[bool]$BuildingBlocksExist = $False

	$Script:Word.Templates.LoadBuildingBlocks()
	#word 2010/2013/2016
	$BuildingBlocksCollection = $Script:Word.Templates | Where-Object{$_.name -eq "Built-In Building Blocks.dotx"}

	Write-Host "$(Get-Date): Attempt to load cover page $($CoverPage)" -BackgroundColor Black -ForegroundColor Yellow
	$part = $Null

	$BuildingBlocksCollection | 
	ForEach-Object {
		If ($_.BuildingBlockEntries.Item($CoverPage).Name -eq $CoverPage) 
		{
			$BuildingBlocks = $_
		}
	}        

	If($Null -ne $BuildingBlocks)
	{
		$BuildingBlocksExist = $True

		Try 
		{
			$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
		}

		Catch
		{
			$part = $Null
		}

		If($Null -ne $part)
		{
			$Script:CoverPagesExist = $True
		}
	}

	If(!$Script:CoverPagesExist)
	{
		Write-Host "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist." -BackgroundColor Black -ForegroundColor Yellow
		Write-Warning "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
		Write-Warning "This report will not have a Cover Page."
	}

	Write-Host "$(Get-Date): Create empty word doc" -BackgroundColor Black -ForegroundColor Yellow
	$Script:Doc = $Script:Word.Documents.Add()
	If($Null -eq $Script:Doc)
	{
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn empty Word document could not be created.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	$Script:Selection = $Script:Word.Selection
	If($Null -eq $Script:Selection)
	{
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "`n`n`t`tAn unknown error happened selecting the entire Word document for default formatting options.`n`n`t`tScript cannot continue.`n`n"
		AbortScript
	}

	#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
	#36 = .50"
	$Script:Word.ActiveDocument.DefaultTabStop = 36

	#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
	Write-Host "$(Get-Date): Disable grammar and spell checking" -BackgroundColor Black -ForegroundColor Yellow
	#bug reported 1-Apr-2014 by Tim Mangan
	#save current options first before turning them off
	$Script:CurrentGrammarOption = $Script:Word.Options.CheckGrammarAsYouType
	$Script:CurrentSpellingOption = $Script:Word.Options.CheckSpellingAsYouType
	$Script:Word.Options.CheckGrammarAsYouType = $False
	$Script:Word.Options.CheckSpellingAsYouType = $False

	If($BuildingBlocksExist)
	{
		#insert new page, getting ready for table of contents
		Write-Host "$(Get-Date): Insert new page, getting ready for table of contents" -BackgroundColor Black -ForegroundColor Yellow
		$part.Insert($Script:Selection.Range,$True) | Out-Null
		$Script:Selection.InsertNewPage()

		#table of contents
		Write-Host "$(Get-Date): Table of Contents - $($Script:MyHash.Word_TableOfContents)" -BackgroundColor Black -ForegroundColor Yellow
		$toc = $BuildingBlocks.BuildingBlockEntries.Item($Script:MyHash.Word_TableOfContents)
		If($Null -eq $toc)
		{
			Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
			Write-Host "$(Get-Date): Table of Content - $($Script:MyHash.Word_TableOfContents) could not be retrieved." -BackgroundColor Black -ForegroundColor Yellow
			Write-Warning "This report will not have a Table of Contents."
		}
		Else
		{
			$toc.insert($Script:Selection.Range,$True) | Out-Null
		}
	}
	Else
	{
		Write-Host "$(Get-Date): Table of Contents are not installed." -BackgroundColor Black -ForegroundColor Yellow
		Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
	}

	#set the footer
	Write-Host "$(Get-Date): Set the footer" -BackgroundColor Black -ForegroundColor Yellow
	[string]$footertext = "Report created by $username"

	#get the footer
	Write-Host "$(Get-Date): Get the footer and format font" -BackgroundColor Black -ForegroundColor Yellow
	$Script:Doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
	#get the footer and format font
	$footers = $Script:Doc.Sections.Last.Footers
	ForEach ($footer in $footers) 
	{
		If($footer.exists) 
		{
			$footer.range.Font.name = "Calibri"
			$footer.range.Font.size = 8
			$footer.range.Font.Italic = $True
			$footer.range.Font.Bold = $True
		}
	} #end ForEach
	Write-Host "$(Get-Date): Footer text" -BackgroundColor Black -ForegroundColor Yellow
	$Script:Selection.HeaderFooter.Range.Text = $footerText

	#add page numbering
	Write-Host "$(Get-Date): Add page numbering" -BackgroundColor Black -ForegroundColor Yellow
	$Script:Selection.HeaderFooter.PageNumbers += ($wdAlignPageNumberRight) | Out-Null

	FindWordDocumentEnd
	Write-Host "$(Get-Date):" -BackgroundColor Black -ForegroundColor Yellow
	#end of Jeff Hicks 
}

Function UpdateDocumentProperties
{
	Param([string]$AbstractTitle, [string]$SubjectTitle)
	#updated 8-Jun-2017 with additional cover page fields
	#Update document properties
	If($MSWORD -or $PDF)
	{
		If($Script:CoverPagesExist)
		{
			Write-Host "$(Get-Date): Set Cover Page Properties" -BackgroundColor Black -ForegroundColor Yellow
			#8-Jun-2017 put these 4 items in alpha order
            Set-DocumentProperty -Document $Script:Doc -DocProperty Author -Value $UserName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Company -Value $Script:CoName
            Set-DocumentProperty -Document $Script:Doc -DocProperty Subject -Value $SubjectTitle
            Set-DocumentProperty -Document $Script:Doc -DocProperty Title -Value $Script:title

			#Get the Coverpage XML part
			$cp = $Script:Doc.CustomXMLParts | Where-Object{$_.NamespaceURI -match "coverPageProps$"}

			#get the abstract XML part
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "Abstract"}
			#set the text
			If([String]::IsNullOrEmpty($Script:CoName))
			{
				[string]$abstract = $AbstractTitle
			}
			Else
			{
				[string]$abstract = "$($AbstractTitle) for $($Script:CoName)"
			}
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyAddress"}
			#set the text
			[string]$abstract = $CompanyAddress
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyEmail"}
			#set the text
			[string]$abstract = $CompanyEmail
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyFax"}
			#set the text
			[string]$abstract = $CompanyFax
			$ab.Text = $abstract

			#added 8-Jun-2017
			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "CompanyPhone"}
			#set the text
			[string]$abstract = $CompanyPhone
			$ab.Text = $abstract

			$ab = $cp.documentelement.ChildNodes | Where-Object{$_.basename -eq "PublishDate"}
			#set the text
			[string]$abstract = (Get-Date -Format d).ToString()
			$ab.Text = $abstract

			Write-Host "$(Get-Date): Update the Table of Contents" -BackgroundColor Black -ForegroundColor Yellow
			#update the Table of Contents
			$Script:Doc.TablesOfContents.item(1).Update()
			$cp = $Null
			$ab = $Null
			$abstract = $Null
		}
	}
}
#endregion

#region registry functions
#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This Function just gets $True or $False
Function Test-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	$key -and $Null -ne $key.GetValue($name, $Null)
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue($path, $name)
{
	$key = Get-Item -LiteralPath $path -EA 0
	If($key)
	{
		$key.GetValue($name, $Null)
	}
	Else
	{
		$Null
	}
}

# Gets the specified registry value or $Null if it is missing
Function Get-RegistryValue2
{
	[CmdletBinding()]
	Param([string]$path, [string]$name, [string]$ComputerName)
	If($ComputerName -eq $env:computername)
	{
		$key = Get-Item -LiteralPath $path -EA 0
		If($key)
		{
			Return $key.GetValue($name, $Null)
		}
		Else
		{
			Return $Null
		}
	}
	Else
	{
		#path needed here is different for remote registry access
		$path = $path.SubString(6)
		$path2 = $path.Replace('\','\\')
		$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $ComputerName)
		$RegKey = $Reg.OpenSubKey($path2)
		If ($RegKey)
		{
			$Results = $RegKey.GetValue($name)

			If($Null -ne $Results)
			{
				Return $Results
			}
			Else
			{
				Return $Null
			}
		}
		Else
		{
			Return $Null
		}
	}
}

Function Get-RegKeyToObject 
{
	#function contributed by Andrew Williamson @ Fujitsu Services
    param([string]$RegPath,
    [string]$RegKey,
    [string]$ComputerName)
	
    $val = Get-RegistryValue2 $RegPath $RegKey $ComputerName
	
    $obj1 = New-Object -TypeName PSObject
	$obj1 | Add-Member -MemberType NoteProperty -Name RegKey	-Value $RegPath
	$obj1 | Add-Member -MemberType NoteProperty -Name RegValue	-Value $RegKey
    If($Null -eq $val) 
	{
        $obj1 | Add-Member -MemberType NoteProperty -Name Value	-Value "Not set"
    } 
	Else 
	{
	    $obj1 | Add-Member -MemberType NoteProperty -Name Value	-Value $val
    }
    $Script:ControllerRegistryItems += ($obj1) 
}
#endregion

#region Word, text, and HTML line output functions
Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
#for creating the formatted text report
#created March 2011
#updated March 2014
{
	Param( [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "`r`n", [switch]$nonewline )
	While( $tabs -gt 0 ) { $Script:Output += "`t"; $tabs--; }
	If( $nonewline )
	{
		$Script:Output += $name + $value
	}
	Else
	{
		$Script:Output += $name + $value + $newline
	}
}
	
Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
#updated 27-Mar-2014 to include font name, font size, italics and bold options
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName=$Null,
	[int]$fontSize=0,
	[bool]$italics=$False,
	[bool]$boldface=$False,
	[Switch]$nonewline)
	
	#Build output style
	[string]$output = ""
	Switch ($style)
	{
		0 {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
		1 {$Script:Selection.Style = $Script:MyHash.Word_Heading1; Break}
		2 {$Script:Selection.Style = $Script:MyHash.Word_Heading2; Break}
		3 {$Script:Selection.Style = $Script:MyHash.Word_Heading3; Break}
		4 {$Script:Selection.Style = $Script:MyHash.Word_Heading4; Break}
		Default {$Script:Selection.Style = $Script:MyHash.Word_NoSpacing; Break}
	}
	
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
 
	If(![String]::IsNullOrEmpty($fontName)) 
	{
		$Script:Selection.Font.name = $fontName
	} 

	If($fontSize -ne 0) 
	{
		$Script:Selection.Font.size = $fontSize
	} 
 
	If($italics -eq $True) 
	{
		$Script:Selection.Font.Italic = $True
	} 
 
	If($boldface -eq $True) 
	{
		$Script:Selection.Font.Bold = $True
	} 

	#output the rest of the parameters.
	$output += $name + $value
	$Script:Selection.TypeText($output)
 
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Script:Selection.TypeParagraph()
	}
}

#***********************************************************************************************************
# WriteHTMLLine
#***********************************************************************************************************

<#
.Synopsis
	Writes a line of output for HTML output
.DESCRIPTION
	This function formats an HTML line
.USAGE
	WriteHTMLLine <Style> <Tabs> <Name> <Value> <Font Name> <Font Size> <Options>

	0 for Font Size denotes using the default font size of 2 or 10 point

.EXAMPLE
	WriteHTMLLine 0 0 " "

	Writes a blank line with no style or tab stops, obviously none needed.

.EXAMPLE
	WriteHTMLLine 0 1 "This is a regular line of text indented 1 tab stops"

	Writes a line with 1 tab stop.

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in italics" "" $Null 0 $htmlitalics

	Writes a line omitting font and font size and setting the italics attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold" "" $Null 0 $htmlbold

	Writes a line omitting font and font size and setting the bold attribute

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in bold italics" "" $Null 0 ($htmlbold -bor $htmlitalics)

	Writes a line omitting font and font size and setting both italics and bold options

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of text in the default font in 10 point" "" $Null 2  # 10 point font

	Writes a line using 10 point font

.EXAMPLE
	WriteHTMLLine 0 0 "This is a regular line of text in Courier New font" "" "Courier New" 0 

	Writes a line using Courier New Font and 0 font point size (default = 2 if set to 0)

.EXAMPLE	
	WriteHTMLLine 0 0 "This is a regular line of RED text indented 0 tab stops with the computer name as data in 10 point Courier New bold italics: " $env:computername "Courier New" 2 ($htmlbold -bor $htmlred -bor $htmlitalics)

	Writes a line using Courier New Font with first and second string values to be used, also uses 10 point font with bold, italics and red color options set.

.NOTES

	Font Size - Unlike word, there is a limited set of font sizes that can be used in HTML.  They are:
		0 - default which actually gives it a 2 or 10 point.
		1 - 7.5 point font size
		2 - 10 point
		3 - 13.5 point
		4 - 15 point
		5 - 18 point
		6 - 24 point
		7 - 36 point
	Any number larger than 7 defaults to 7

	Style - Refers to the headers that are used with output and resemble the headers in word, HTML supports headers h1-h6 and h1-h4 are more commonly used.  Unlike word, H1 will not give you
	a blue colored font, you will have to set that yourself.

	Colors and Bold/Italics Flags are:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack       
#>

Function WriteHTMLLine
#Function created by Ken Avram
#Function created to make output to HTML easy in this script
#headings fixed 12-Oct-2016 by Webster
#errors with $HTMLStyle fixed 7-Dec-2017 by Webster
{
	Param([int]$style=0, 
	[int]$tabs = 0, 
	[string]$name = '', 
	[string]$value = '', 
	[string]$fontName="Calibri",
	[int]$fontSize=1,
	[int]$options=$htmlblack)


	#Build output style
	[string]$output = ""

	If([String]::IsNullOrEmpty($Name))	
	{
		$HTMLBody = "<p></p>"
	}
	Else
	{
		$color = CheckHTMLColor $options

		#build # of tabs

		While($tabs -gt 0)
		{ 
			$output += "&nbsp;&nbsp;&nbsp;&nbsp;"; $tabs--; 
		}

		$HTMLFontName = $fontName		

		$HTMLBody = ""

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "<i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "<b>"
		} 

		#output the rest of the parameters.
		$output += $name + $value

		Switch ($style)
		{
			1 {$HTMLStyle = "<h1>"; Break}
			2 {$HTMLStyle = "<h2>"; Break}
			3 {$HTMLStyle = "<h3>"; Break}
			4 {$HTMLStyle = "<h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		$HTMLBody += $HTMLStyle + $output

		Switch ($style)
		{
			1 {$HTMLStyle = "</h1>"; Break}
			2 {$HTMLStyle = "</h2>"; Break}
			3 {$HTMLStyle = "</h3>"; Break}
			4 {$HTMLStyle = "</h4>"; Break}
			Default {$HTMLStyle = ""; Break}
		}

		#added by webster 12-oct-2016
		#if a heading, don't add the <br>
		#moved to after the two switch statements on 7-Dec-2017 to fix $HTMLStyle has not been set error
		If($HTMLStyle -eq "")
		{
			$HTMLBody += "<br><font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		Else
		{
			$HTMLBody += "<font face='" + $HTMLFontName + "' " + "color='" + $color + "' size='"  + $fontsize + "'>"
		}
		
		$HTMLBody += $HTMLStyle +  "</font>"

		If($options -band $htmlitalics) 
		{
			$HTMLBody += "</i>"
		} 

		If($options -band $htmlbold) 
		{
			$HTMLBody += "</b>"
		} 

		#added by webster 12-oct-2016
		#if a heading, don't add the <br />
		#moved to inside the Else statement on 7-Dec-2017 to fix $HTMLStyle has not been set error
		If($HTMLStyle -eq "")
		{
			$HTMLBody += "<br />"
		}
	}

	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody
}
#endregion

#region HTML table functions
#***********************************************************************************************************
# AddHTMLTable - Called from FormatHTMLTable function
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************
Function AddHTMLTable
{
	Param([string]$fontName="Calibri",
	[int]$fontSize=2,
	[int]$colCount=0,
	[int]$rowCount=0,
	[object[]]$rowInfo=@(),
	[object[]]$fixedInfo=@())

	For($rowidx = $RowIndex;$rowidx -le $rowCount;$rowidx++)
	{
		$rd = @($rowInfo[$rowidx - 2])
		$htmlbody = $htmlbody + "<tr>"
		For($columnIndex = 0; $columnIndex -lt $colCount; $columnindex+=2)
		{
			$tmp = CheckHTMLColor $rd[$columnIndex+1]

			If($fixedInfo.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedInfo[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $rd[$columnIndex])
			{
				$cell = $rd[$columnIndex].tostring()
				If($cell -eq " " -or $cell.length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $cell.length;$i++)
					{
						If($cell[$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($cell[$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $cell
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($rd[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($rd[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	out-file -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 
}

#***********************************************************************************************************
# FormatHTMLTable 
# Created by Ken Avram
# modified by Jake Rutski
#***********************************************************************************************************

<#
.Synopsis
	Format table for HTML output document
.DESCRIPTION
	This function formats a table for HTML from an array of strings
.PARAMETER noBorder
	If set to $true, a table will be generated without a border (border='0')
.PARAMETER noHeadCols
	This parameter should be used when generating tables without column headers
	Set this parameter equal to the number of columns in the table
.PARAMETER rowArray
	This parameter contains the row data array for the table
.PARAMETER columnArray
	This parameter contains column header data for the table
.PARAMETER fixedWidth
	This parameter contains widths for columns in pixel format ("100px") to override auto column widths
	The variable should contain a width for each column you wish to override the auto-size setting
	For example: $columnWidths = @("100px","110px","120px","130px","140px")

.USAGE
	FormatHTMLTable <Table Header> <Table Format> <Font Name> <Font Size>

.EXAMPLE
	FormatHTMLTable "Table Heading" "auto" "Calibri" 3

	This example formats a table and writes it out into an html file.  All of the parameters are optional
	defaults are used if not supplied.

	for <Table format>, the default is auto which will autofit the text into the columns and adjust to the longest text in that column.  You can also use percentage i.e. 25%
	which will take only 25% of the line and will auto word wrap the text to the next line in the column.  Also, instead of using a percentage, you can use pixels i.e. 400px.

	FormatHTMLTable "Table Heading" "auto" -rowArray $rowData -columnArray $columnData

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, column header data from $columnData and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -noHeadCols 3

	This example creates an HTML table with a heading of 'Table Heading', auto column spacing, no header, and row data from $rowData

	FormatHTMLTable "Table Heading" -rowArray $rowData -fixedWidth $fixedColumns

	This example creates an HTML table with a heading of 'Table Heading, no header, row data from $rowData, and fixed columns defined by $fixedColumns

.NOTES
	In order to use the formatted table it first has to be loaded with data.  Examples below will show how to load the table:

	First, initialize the table array

	$rowdata = @()

	Then Load the array.  If you are using column headers then load those into the column headers array, otherwise the first line of the table goes into the column headers array
	and the second and subsequent lines go into the $rowdata table as shown below:

	$columnHeaders = @('Display Name',($htmlsilver -bor $htmlbold),'Status',($htmlsilver -bor $htmlbold),'Startup Type',($htmlsilver -bor $htmlbold))

	The first column is the actual name to display, the second are the attributes of the column i.e. color anded with bold or italics.  For the anding, parens are required or it will
	not format correctly.

	This is following by adding rowdata as shown below.  As more columns are added the columns will auto adjust to fit the size of the page.

	$rowdata = @()
	$columnHeaders = @("User Name",($htmlsilver -bor $htmlbold),$UserName,$htmlwhite)
	$rowdata += @(,('Save as PDF',($htmlsilver -bor $htmlbold),$PDF.ToString(),$htmlwhite))
	$rowdata += @(,('Save as TEXT',($htmlsilver -bor $htmlbold),$TEXT.ToString(),$htmlwhite))
	$rowdata += @(,('Save as WORD',($htmlsilver -bor $htmlbold),$MSWORD.ToString(),$htmlwhite))
	$rowdata += @(,('Save as HTML',($htmlsilver -bor $htmlbold),$HTML.ToString(),$htmlwhite))
	$rowdata += @(,('Add DateTime',($htmlsilver -bor $htmlbold),$AddDateTime.ToString(),$htmlwhite))
	$rowdata += @(,('Hardware Inventory',($htmlsilver -bor $htmlbold),$Hardware.ToString(),$htmlwhite))
	$rowdata += @(,('Computer Name',($htmlsilver -bor $htmlbold),$ComputerName,$htmlwhite))
	$rowdata += @(,('Filename1',($htmlsilver -bor $htmlbold),$Script:FileName1,$htmlwhite))
	$rowdata += @(,('OS Detected',($htmlsilver -bor $htmlbold),$Script:RunningOS,$htmlwhite))
	$rowdata += @(,('PSUICulture',($htmlsilver -bor $htmlbold),$PSCulture,$htmlwhite))
	$rowdata += @(,('PoSH version',($htmlsilver -bor $htmlbold),$Host.Version.ToString(),$htmlwhite))
	FormatHTMLTable "Example of Horizontal AutoFitContents HTML Table" -rowArray $rowdata

	The 'rowArray' paramater is mandatory to build the table, but it is not set as such in the function - if nothing is passed, the table will be empty.

	Colors and Bold/Italics Flags are shown below:

		htmlbold       
		htmlitalics    
		htmlred        
		htmlcyan        
		htmlblue       
		htmldarkblue   
		htmllightblue   
		htmlpurple      
		htmlyellow      
		htmllime       
		htmlmagenta     
		htmlwhite       
		htmlsilver      
		htmlgray       
		htmlolive       
		htmlorange      
		htmlmaroon      
		htmlgreen       
		htmlblack     

#>

Function FormatHTMLTable
{
	Param([string]$tableheader,
	[string]$tablewidth="auto",
	[string]$fontName="Calibri",
	[int]$fontSize=2,
	[switch]$noBorder=$false,
	[int]$noHeadCols=1,
	[object[]]$rowArray=@(),
	[object[]]$fixedWidth=@(),
	[object[]]$columnArray=@())

	$HTMLBody = "<b><font face='" + $fontname + "' size='" + ($fontsize + 1) + "'>" + $tableheader + "</font></b>"

	If($columnArray.Length -eq 0)
	{
		$NumCols = $noHeadCols + 1
	}  # means we have no column headers, just a table
	Else
	{
		$NumCols = $columnArray.Length
	}  # need to add one for the color attrib

	If($Null -ne $rowArray)
	{
		$NumRows = $rowArray.length + 1
	}
	Else
	{
		$NumRows = 1
	}

	If($noBorder)
	{
		$htmlbody += "<table border='0' width='" + $tablewidth + "'>"
	}
	Else
	{
		$htmlbody += "<table border='1' width='" + $tablewidth + "'>"
	}

	If(!($columnArray.Length -eq 0))
	{
		$htmlbody += "<tr>"

		For($columnIndex = 0; $columnIndex -lt $NumCols; $columnindex+=2)
		{
			$tmp = CheckHTMLColor $columnArray[$columnIndex+1]
			If($fixedWidth.Length -eq 0)
			{
				$htmlbody += "<td style=""background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}
			Else
			{
				$htmlbody += "<td style=""width:$($fixedWidth[$columnIndex/2]); background-color:$($tmp)""><font face='$($fontName)' size='$($fontSize)'>"
			}

			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "<b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "<i>"
			}
			If($Null -ne $columnArray[$columnIndex])
			{
				If($columnArray[$columnIndex] -eq " " -or $columnArray[$columnIndex].length -eq 0)
				{
					$htmlbody += "&nbsp;&nbsp;&nbsp;"
				}
				Else
				{
					For($i=0;$i -lt $columnArray[$columnIndex].length;$i+=2)
					{
						If($columnArray[$columnIndex][$i] -eq " ")
						{
							$htmlbody += "&nbsp;"
						}
						If($columnArray[$columnIndex][$i] -ne " ")
						{
							Break
						}
					}
					$htmlbody += $columnArray[$columnIndex]
				}
			}
			Else
			{
				$htmlbody += "&nbsp;&nbsp;&nbsp;"
			}
			If($columnArray[$columnIndex+1] -band $htmlbold)
			{
				$htmlbody += "</b>"
			}
			If($columnArray[$columnIndex+1] -band $htmlitalics)
			{
				$htmlbody += "</i>"
			}
			$htmlbody += "</font></td>"
		}
		$htmlbody += "</tr>"
	}
	$rowindex = 2
	If($Null -ne $rowArray)
	{
		AddHTMLTable $fontName $fontSize -colCount $numCols -rowCount $NumRows -rowInfo $rowArray -fixedInfo $fixedWidth
		$rowArray = @()
		$htmlbody = "</table>"
	}
	Else
	{
		$HTMLBody += "</table>"
	}	

	Out-File -FilePath $Script:FileName1 -Append -InputObject $HTMLBody 
}
#endregion

#region other HTML functions
#***********************************************************************************************************
# CheckHTMLColor - Called from AddHTMLTable WriteHTMLLine and FormatHTMLTable
#***********************************************************************************************************
Function CheckHTMLColor
{
	Param($hash)

	If($hash -band $htmlwhite)
	{
		Return $htmlwhitemask
	}
	If($hash -band $htmlred)
	{
		Return $htmlredmask
	}
	If($hash -band $htmlcyan)
	{
		Return $htmlcyanmask
	}
	If($hash -band $htmlblue)
	{
		Return $htmlbluemask
	}
	If($hash -band $htmldarkblue)
	{
		Return $htmldarkbluemask
	}
	If($hash -band $htmllightblue)
	{
		Return $htmllightbluemask
	}
	If($hash -band $htmlpurple)
	{
		Return $htmlpurplemask
	}
	If($hash -band $htmlyellow)
	{
		Return $htmlyellowmask
	}
	If($hash -band $htmllime)
	{
		Return $htmllimemask
	}
	If($hash -band $htmlmagenta)
	{
		Return $htmlmagentamask
	}
	If($hash -band $htmlsilver)
	{
		Return $htmlsilvermask
	}
	If($hash -band $htmlgray)
	{
		Return $htmlgraymask
	}
	If($hash -band $htmlblack)
	{
		Return $htmlblackmask
	}
	If($hash -band $htmlorange)
	{
		Return $htmlorangemask
	}
	If($hash -band $htmlmaroon)
	{
		Return $htmlmaroonmask
	}
	If($hash -band $htmlgreen)
	{
		Return $htmlgreenmask
	}
	If($hash -band $htmlolive)
	{
		Return $htmlolivemask
	}
}

Function SetupHTML
{
	Write-Host "$(Get-Date): Setting up HTML" -BackgroundColor Black -ForegroundColor Yellow
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).html"
	}

	$htmlhead = "<html><head><meta http-equiv='Content-Language' content='da'><title>" + $Script:Title + "</title></head><body>"
	out-file -FilePath $Script:Filename1 -Force -InputObject $HTMLHead
}
#endregion

#region Iain's Word table functions

<#
.Synopsis
	Add a table to a Microsoft Word document
.DESCRIPTION
	This function adds a table to a Microsoft Word document from either an array of
	Hashtables or an array of PSCustomObjects.

	Using this function is quicker than setting each table cell individually but can
	only utilise the built-in MS Word table autoformats. Individual tables cells can
	be altered after the table has been appended to the document (a table reference
	is Returned).
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. Column headers will display the key names as defined.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -List

	This example adds table to the MS Word document, utilising all key/value pairs in
	the array of hashtables. No column headers will be added, in a ListView format.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray

	This example adds table to the MS Word document, utilising all note property names
	the array of PSCustomObjects. Column headers will display the note property names.
	Note: the columns might not be displayed in the order that they were defined. To
	ensure columns are displayed in the required order utilise the -Columns parameter.
.EXAMPLE
	AddWordTable -Hashtable $HashtableArray -Columns FirstName,LastName,EmailAddress

	This example adds a table to the MS Word document, but only using the specified
	key names: FirstName, LastName and EmailAddress. If other keys are present in the
	array of Hashtables they will be ignored.
.EXAMPLE
	AddWordTable -CustomObject $PSCustomObjectArray -Columns FirstName,LastName,EmailAddress -Headers "First Name","Last Name","Email Address"

	This example adds a table to the MS Word document, but only using the specified
	PSCustomObject note properties: FirstName, LastName and EmailAddress. If other note
	properties are present in the array of PSCustomObjects they will be ignored. The
	display names for each specified column header has been overridden to display a
	custom header. Note: the order of the header names must match the specified columns.
#>

Function AddWordTable
{
	[CmdletBinding()]
	Param
	(
		# Array of Hashtable (including table headers)
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='Hashtable', Position=0)]
		[ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Hashtable,
		# Array of PSCustomObjects
		[Parameter(Mandatory=$True, ValueFromPipelineByPropertyName=$True, ParameterSetName='CustomObject', Position=0)]
		[ValidateNotNullOrEmpty()] [PSCustomObject[]] $CustomObject,
		# Array of Hashtable key names or PSCustomObject property names to include, in display order.
		# If not supplied then all Hashtable keys or all PSCustomObject properties will be displayed.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Columns = $Null,
		# Array of custom table header strings in display order.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [string[]] $Headers = $Null,
		# AutoFit table behavior.
		[Parameter(ValueFromPipelineByPropertyName=$True)] [AllowNull()] [int] $AutoFit = -1,
		# List view (no headers)
		[Switch] $List,
		# Grid lines
		[Switch] $NoGridLines,
		[Switch] $NoInternalGridLines,
		# Built-in Word table formatting style constant
		# Would recommend only $wdTableFormatContempory for normal usage (possibly $wdTableFormatList5 for List view)
		[Parameter(ValueFromPipelineByPropertyName=$True)] [int] $Format = 0
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'" -f $PSCmdlet.ParameterSetName);
		## Check if -Columns wasn't specified but -Headers were (saves some additional parameter sets!)
		If(($Null -eq $Columns) -and ($Null -eq $Headers)) 
		{
			Write-Warning "No columns specified and therefore, specified headers will be ignored.";
			$Columns = $Null;
		}
		ElseIf(($Null -ne $Columns) -and ($Null -ne $Headers)) 
		{
			## Check if number of specified -Columns matches number of specified -Headers
			If($Columns.Length -ne $Headers.Length) 
			{
				Write-Error "The specified number of columns does not match the specified number of headers.";
			}
		} ## end elseif
	} ## end Begin

	Process
	{
		## Build the Word table data string to be converted to a range and then a table later.
		[System.Text.StringBuilder] $WordRangeString = New-Object System.Text.StringBuilder;

		Switch ($PSCmdlet.ParameterSetName) 
		{
			'CustomObject' 
			{
				If($Null -eq $Columns) 
				{
					## Build the available columns from all availble PSCustomObject note properties
					[string[]] $Columns = @();
					## Add each NoteProperty name to the array
					ForEach($Property in ($CustomObject | Get-Member -MemberType NoteProperty)) 
					{ 
						$Columns += $Property.Name; 
					}
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{ 
                        [ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}

				## Iterate through each PSCustomObject
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Object in $CustomObject) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Object.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach
				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f ($CustomObject.Count));
			} ## end CustomObject

			Default 
			{   ## Hashtable
				If($Null -eq $Columns) 
				{
					## Build the available columns from all available hashtable keys. Hopefully
					## all Hashtables have the same keys (they should for a table).
					$Columns = $Hashtable[0].Keys;
				}

				## Add the table headers from -Headers or -Columns (except when in -List(view)
				If(-not $List) 
				{
					Write-Debug ("$(Get-Date): `t`tBuilding table headers");
					If($Null -ne $Headers) 
					{ 
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Headers));
					}
					Else 
					{
						[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $Columns));
					}
				}
                
				## Iterate through each Hashtable
				Write-Debug ("$(Get-Date): `t`tBuilding table rows");
				ForEach($Hash in $Hashtable) 
				{
					$OrderedValues = @();
					## Add each row item in the specified order
					ForEach($Column in $Columns) 
					{ 
						$OrderedValues += $Hash.$Column; 
					}
					## Use the ordered list to add each column in specified order
					[ref] $Null = $WordRangeString.AppendFormat("{0}`n", [string]::Join("`t", $OrderedValues));
				} ## end ForEach

				Write-Debug ("$(Get-Date): `t`t`tAdded '{0}' table rows" -f $Hashtable.Count);
			} ## end default
		} ## end switch

		## Create a MS Word range and set its text to our tab-delimited, concatenated string
		Write-Debug ("$(Get-Date): `t`tBuilding table range");
		$WordRange = $Script:Doc.Application.Selection.Range;
		$WordRange.Text = $WordRangeString.ToString();

		## Create hash table of named arguments to pass to the ConvertToTable method
		$ConvertToTableArguments = @{ Separator = [Microsoft.Office.Interop.Word.WdTableFieldSeparator]::wdSeparateByTabs; }

		## Negative built-in styles are not supported by the ConvertToTable method
		If($Format -ge 0) 
		{
			$ConvertToTableArguments.Add("Format", $Format);
			$ConvertToTableArguments.Add("ApplyBorders", $True);
			$ConvertToTableArguments.Add("ApplyShading", $True);
			$ConvertToTableArguments.Add("ApplyFont", $True);
			$ConvertToTableArguments.Add("ApplyColor", $True);
			If(!$List) 
			{ 
				$ConvertToTableArguments.Add("ApplyHeadingRows", $True); 
			}
			$ConvertToTableArguments.Add("ApplyLastRow", $True);
			$ConvertToTableArguments.Add("ApplyFirstColumn", $True);
			$ConvertToTableArguments.Add("ApplyLastColumn", $True);
		}

		## Invoke ConvertToTable method - with named arguments - to convert Word range to a table
		## See http://msdn.microsoft.com/en-us/library/office/aa171893(v=office.11).aspx
		Write-Debug ("$(Get-Date): `t`tConverting range to table");
		## Store the table reference just in case we need to set alternate row coloring
		$WordTable = $WordRange.GetType().InvokeMember(
			"ConvertToTable",                               # Method name
			[System.Reflection.BindingFlags]::InvokeMethod, # Flags
			$Null,                                          # Binder
			$WordRange,                                     # Target (self!)
			([Object[]]($ConvertToTableArguments.Values)),  ## Named argument values
			$Null,                                          # Modifiers
			$Null,                                          # Culture
			([String[]]($ConvertToTableArguments.Keys))     ## Named argument names
		);

		## Implement grid lines (will wipe out any existing formatting
		If($Format -lt 0) 
		{
			Write-Debug ("$(Get-Date): `t`tSetting table format");
			$WordTable.Style = $Format;
		}

		## Set the table autofit behavior
		If($AutoFit -ne -1) 
		{ 
			$WordTable.AutoFitBehavior($AutoFit); 
		}

		If(!$List)
		{
			#the next line causes the heading row to flow across page breaks
			$WordTable.Rows.First.Headingformat = $wdHeadingFormatTrue;
		}

		If(!$NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleSingle;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}
		If($NoGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleNone;
		}
		If($NoInternalGridLines) 
		{
			$WordTable.Borders.InsideLineStyle = $wdLineStyleNone;
			$WordTable.Borders.OutsideLineStyle = $wdLineStyleSingle;
		}

		Return $WordTable;

	} ## end Process
}

<#
.Synopsis
	Sets the format of one or more Word table cells
.DESCRIPTION
	This function sets the format of one or more table cells, either from a collection
	of Word COM object cell references, an individual Word COM object cell reference or
	a hashtable containing Row and Column information.

	The font name, font size, bold, italic , underline and shading values can be used.
.EXAMPLE
	SetWordCellFormat -Hashtable $Coordinates -Table $TableReference -Bold

	This example sets all text to bold that is contained within the $TableReference
	Word table, using an array of hashtables. Each hashtable contain a pair of co-
	ordinates that is used to select the required cells. Note: the hashtable must
	contain the .Row and .Column key names. For example:
	@ { Row = 7; Column = 3 } to set the cell at row 7 and column 3 to bold.
.EXAMPLE
	$RowCollection = $Table.Rows.First.Cells
	SetWordCellFormat -Collection $RowCollection -Bold -Size 10

	This example sets all text to size 8 and bold for all cells that are contained
	within the first row of the table.
	Note: the $Table.Rows.First.Cells Returns a collection of Word COM cells objects
	that are in the first table row.
.EXAMPLE
	$ColumnCollection = $Table.Columns.Item(2).Cells
	SetWordCellFormat -Collection $ColumnCollection -BackgroundColor 255

	This example sets the background (shading) of all cells in the table's second
	column to red.
	Note: the $Table.Columns.Item(2).Cells Returns a collection of Word COM cells objects
	that are in the table's second column.
.EXAMPLE
	SetWordCellFormat -Cell $Table.Cell(17,3) -Font "Tahoma" -Color 16711680

	This example sets the font to Tahoma and the text color to blue for the cell located
	in the table's 17th row and 3rd column.
	Note: the $Table.Cell(17,3) Returns a single Word COM cells object.
#>

Function SetWordCellFormat 
{
	[CmdletBinding(DefaultParameterSetName='Collection')]
	Param (
		# Word COM object cell collection reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, ParameterSetName='Collection', Position=0)] [ValidateNotNullOrEmpty()] $Collection,
		# Word COM object individual cell reference
		[Parameter(Mandatory=$True, ParameterSetName='Cell', Position=0)] [ValidateNotNullOrEmpty()] $Cell,
		# Hashtable of cell co-ordinates
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=0)] [ValidateNotNullOrEmpty()] [System.Collections.Hashtable[]] $Coordinates,
		# Word COM object table reference
		[Parameter(Mandatory=$True, ParameterSetName='Hashtable', Position=1)] [ValidateNotNullOrEmpty()] $Table,
		# Font name
		[Parameter()] [AllowNull()] [string] $Font = $Null,
		# Font color
		[Parameter()] [AllowNull()] $Color = $Null,
		# Font size
		[Parameter()] [ValidateNotNullOrEmpty()] [int] $Size = 0,
		# Cell background color
		[Parameter()] [AllowNull()] $BackgroundColor = $Null,
		# Force solid background color
		[Switch] $Solid,
		[Switch] $Bold,
		[Switch] $Italic,
		[Switch] $Underline
	)

	Begin 
	{
		Write-Debug ("Using parameter set '{0}'." -f $PSCmdlet.ParameterSetName);
	}

	Process 
	{
		Switch ($PSCmdlet.ParameterSetName) 
		{
			'Collection' {
				ForEach($Cell in $Collection) 
				{
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				} # end ForEach
			} # end Collection
			'Cell' 
			{
				If($Bold) { $Cell.Range.Font.Bold = $True; }
				If($Italic) { $Cell.Range.Font.Italic = $True; }
				If($Underline) { $Cell.Range.Font.Underline = 1; }
				If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
				If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
				If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
				If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
				If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
			} # end Cell
			'Hashtable' 
			{
				ForEach($Coordinate in $Coordinates) 
				{
					$Cell = $Table.Cell($Coordinate.Row, $Coordinate.Column);
					If($Bold) { $Cell.Range.Font.Bold = $True; }
					If($Italic) { $Cell.Range.Font.Italic = $True; }
					If($Underline) { $Cell.Range.Font.Underline = 1; }
					If($Null -ne $Font) { $Cell.Range.Font.Name = $Font; }
					If($Null -ne $Color) { $Cell.Range.Font.Color = $Color; }
					If($Size -ne 0) { $Cell.Range.Font.Size = $Size; }
					If($Null -ne $BackgroundColor) { $Cell.Shading.BackgroundPatternColor = $BackgroundColor; }
					If($Solid) { $Cell.Shading.Texture = 0; } ## wdTextureNone
				}
			} # end Hashtable
		} # end switch
	} # end process
}

<#
.Synopsis
	Sets alternate row colors in a Word table
.DESCRIPTION
	This function sets the format of alternate rows within a Word table using the
	specified $BackgroundColor. This function is expensive (in performance terms) as
	it recursively sets the format on alternate rows. It would be better to pick one
	of the predefined table formats (if one exists)? Obviously the more rows, the
	longer it takes :'(

	Note: this function is called by the AddWordTable function if an alternate row
	format is specified.
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 255

	This example sets every-other table (starting with the first) row and sets the
	background color to red (wdColorRed).
.EXAMPLE
	SetWordTableAlternateRowColor -Table $TableReference -BackgroundColor 39423 -Seed Second

	This example sets every other table (starting with the second) row and sets the
	background color to light orange (weColorLightOrange).
#>

Function SetWordTableAlternateRowColor 
{
	[CmdletBinding()]
	Param (
		# Word COM object table reference
		[Parameter(Mandatory=$True, ValueFromPipeline=$True, Position=0)] [ValidateNotNullOrEmpty()] $Table,
		# Alternate row background color
		[Parameter(Mandatory=$True, Position=1)] [ValidateNotNull()] [int] $BackgroundColor,
		# Alternate row starting seed
		[Parameter(ValueFromPipelineByPropertyName=$True, Position=2)] [ValidateSet('First','Second')] [string] $Seed = 'First'
	)

	Process 
	{
		$StartDateTime = Get-Date;
		Write-Debug ("{0}: `t`tSetting alternate table row colors.." -f $StartDateTime);

		## Determine the row seed (only really need to check for 'Second' and default to 'First' otherwise
		If($Seed.ToLower() -eq 'second') 
		{ 
			$StartRowIndex = 2; 
		}
		Else 
		{ 
			$StartRowIndex = 1; 
		}

		For($AlternateRowIndex = $StartRowIndex; $AlternateRowIndex -lt $Table.Rows.Count; $AlternateRowIndex += 2) 
		{ 
			$Table.Rows.Item($AlternateRowIndex).Shading.BackgroundPatternColor = $BackgroundColor;
		}

		## I've put verbose calls in here we can see how expensive this functionality actually is.
		$EndDateTime = Get-Date;
		$ExecutionTime = New-TimeSpan -Start $StartDateTime -End $EndDateTime;
		Write-Debug ("{0}: `t`tDone setting alternate row style color in '{1}' seconds" -f $EndDateTime, $ExecutionTime.TotalSeconds);
	}
}
#endregion

#region general script functions
Function Check-LoadedModule
#Function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#bug fixed by Peter Bosen
#This Function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param([parameter(Mandatory = $True)][alias("Module")][string]$ModuleName)
	#$LoadedModules = Get-Module | Select-Object Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module | ForEach-Object { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work if the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non Default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	
	[string]$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If($ModuleFound -ne $ModuleName) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0
		If($module -and $?)
		{
			# module imported properly
			Return $True
		}
		Else
		{
			# module import failed
			Return $False
		}
	}
	Else
	{
		#module already imported into current session
		Return $True
	}
}

Function SaveandCloseDocumentandShutdownWord
{
	#bug fix 1-Apr-2014
	#reset Grammar and Spelling options back to their original settings
	$Script:Word.Options.CheckGrammarAsYouType = $Script:CurrentGrammarOption
	$Script:Word.Options.CheckSpellingAsYouType = $Script:CurrentSpellingOption

	Write-Host "$(Get-Date): Save and Close document and Shutdown Word" -BackgroundColor Black -ForegroundColor Yellow
	If($Script:WordVersion -eq $wdWord2010)
	{
		#the $saveFormat below passes StrictMode 2
		#I found this at the following two links
		#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
		#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
		If($PDF)
		{
			Write-Host "$(Get-Date): Saving as DOCX file first before saving to PDF" -BackgroundColor Black -ForegroundColor Yellow
		}
		Else
		{
			Write-Host "$(Get-Date): Saving DOCX file" -BackgroundColor Black -ForegroundColor Yellow
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Host "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)" -BackgroundColor Black -ForegroundColor Yellow
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$Script:Doc.SaveAs([REF]$Script:FileName1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Host "$(Get-Date): Now saving as PDF" -BackgroundColor Black -ForegroundColor Yellow
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$saveFormat)
		}
	}
	ElseIf($Script:WordVersion -eq $wdWord2013 -or $Script:WordVersion -eq $wdWord2016)
	{
		If($PDF)
		{
			Write-Host "$(Get-Date): Saving as DOCX file first before saving to PDF" -BackgroundColor Black -ForegroundColor Yellow
		}
		Else
		{
			Write-Host "$(Get-Date): Saving DOCX file" -BackgroundColor Black -ForegroundColor Yellow
		}
		If($AddDateTime)
		{
			$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).docx"
			If($PDF)
			{
				$Script:FileName2 += "_$(Get-Date -f yyyy-MM-dd_HHmm).pdf"
			}
		}
		Write-Host "$(Get-Date): Running $($Script:WordProduct) and detected operating system $($Script:RunningOS)" -BackgroundColor Black -ForegroundColor Yellow
		$Script:Doc.SaveAs2([REF]$Script:FileName1, [ref]$wdFormatDocumentDefault)
		If($PDF)
		{
			Write-Host "$(Get-Date): Now saving as PDF" -BackgroundColor Black -ForegroundColor Yellow
			$Script:Doc.SaveAs([REF]$Script:FileName2, [ref]$wdFormatPDF)
		}
	}

	Write-Host "$(Get-Date): Closing Word" -BackgroundColor Black -ForegroundColor Yellow
	$Script:Doc.Close()
	$Script:Word.Quit()
	If($PDF)
	{
		[int]$cnt = 0
		While(Test-Path $Script:FileName1)
		{
			$cnt++
			If($cnt -gt 1)
			{
				Write-Host "$(Get-Date): Waiting another 10 seconds to allow Word to fully close (try # $($cnt))" -BackgroundColor Black -ForegroundColor Yellow
				Start-Sleep -Seconds 10
				$Script:Word.Quit()
				If($cnt -gt 2)
				{
					#kill the winword process

					#find out our session (usually "1" except on TS/RDC or Citrix)
					$SessionID = (Get-Process -PID $PID).SessionId
					
					#Find out if winword is running in our session
					$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}).Id
					If($wordprocess -gt 0)
					{
						Write-Host "$(Get-Date): Attempting to stop WinWord process # $($wordprocess)" -BackgroundColor Black -ForegroundColor Yellow
						Stop-Process $wordprocess -EA 0
					}
				}
			}
			Write-Host "$(Get-Date): Attempting to delete $($Script:FileName1) since only $($Script:FileName2) is needed (try # $($cnt))" -BackgroundColor Black -ForegroundColor Yellow
			Remove-Item $Script:FileName1 -EA 0
		}
	}
	Write-Host "$(Get-Date): System Cleanup" -BackgroundColor Black -ForegroundColor Yellow
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
	If(Test-Path variable:global:word)
	{
		Remove-Variable -Name word -Scope Global
	}
	$SaveFormat = $Null
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	
	#is the winword process still running? kill it

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId

	#Find out if winword is running in our session
	$wordprocess = $Null
	$wordprocess = ((Get-Process 'WinWord' -ea 0) | Where-Object {$_.SessionId -eq $SessionID}).Id
	If($null -ne $wordprocess -and $wordprocess -gt 0)
	{
		Write-Host "$(Get-Date): WinWord process is still running. Attempting to stop WinWord process # $($wordprocess)" -BackgroundColor Black -ForegroundColor Yellow
		Stop-Process $wordprocess -EA 0
	}
}

Function SaveandCloseTextDocument
{
	If($AddDateTime)
	{
		$Script:FileName1 += "_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
	}

	Write-Output $Script:Output | Out-File $Script:Filename1
}

Function SaveandCloseHTMLDocument
{
	Out-File -FilePath $Script:FileName1 -Append -InputObject "<p></p></body></html>"
}

Function SetFileName1andFileName2
{
	Param([string]$OutputFileName)
	
	If($Folder -eq "")
	{
		$pwdpath = $pwd.Path
	}
	Else
	{
		$pwdpath = $Folder
	}

	If($pwdpath.EndsWith("\"))
	{
		#remove the trailing \
		$pwdpath = $pwdpath.SubString(0, ($pwdpath.Length - 1))
	}

	#set $Script:Filename1 and $Script:Filename2 with no file extension
	If($AddDateTime)
	{
		[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName)"
		If($PDF)
		{
			[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName)"
		}
	}

	If($MSWord -or $PDF)
	{
		CheckWordPreReq
		
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).docx"
			If($PDF)
			{
				[string]$Script:FileName2 = "$($pwdpath)\$($OutputFileName).pdf"
			}
		}

		SetupWord
	}
	ElseIf($Text)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).txt"
		}
		ShowScriptOptions
	}
	ElseIf($HTML)
	{
		If(!$AddDateTime)
		{
			[string]$Script:FileName1 = "$($pwdpath)\$($OutputFileName).html"
		}
		SetupHTML
		ShowScriptOptions
	}
}

Function ProcessDocumentOutput
{
	If($MSWORD -or $PDF)
	{
		SaveandCloseDocumentandShutdownWord
	}
	ElseIf($Text)
	{
		SaveandCloseTextDocument
	}
	ElseIf($HTML)
	{
		SaveandCloseHTMLDocument
	}

	$GotFile = $False

	If($PDF)
	{
		If(Test-Path "$($Script:FileName2)")
		{
			Write-Host "$(Get-Date): $($Script:FileName2) is ready for use" -BackgroundColor Black -ForegroundColor Yellow
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName2)"
			Write-Error "Unable to save the output file, $($Script:FileName2)"
		}
	}
	Else
	{
		If(Test-Path "$($Script:FileName1)")
		{
			Write-Host "$(Get-Date): $($Script:FileName1) is ready for use" -BackgroundColor Black -ForegroundColor Yellow
			$GotFile = $True
		}
		Else
		{
			Write-Warning "$(Get-Date): Unable to save the output file, $($Script:FileName1)"
			Write-Error "Unable to save the output file, $($Script:FileName1)"
		}
	}
}

Function ShowScriptOptions
{
	Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Add DateTime       : $($AddDateTime)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): AdminAddress       : $($AdminAddress)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Administrators     : $($Administrators)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Applications       : $($Applications)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Company Name       : $($Script:CoName)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Company Address    : $($CompanyAddress)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Company Email      : $($CompanyEmail)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Company Fax        : $($CompanyFax)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Company Phone      : $($CompanyPhone)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Cover Page         : $($CoverPage)" -BackgroundColor Black -ForegroundColor Yellow
	If($Dev)
	{
		Write-Host "$(Get-Date): DevErrorFile       : $($Script:DevErrorFile)" -BackgroundColor Black -ForegroundColor Yellow
	}
	Write-Host "$(Get-Date): Farm name          : $($Script:FarmName)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Filename1          : $($Script:filename1)" -BackgroundColor Black -ForegroundColor Yellow
	If($PDF)
	{
		Write-Host "$(Get-Date): Filename2          : $($Script:Filename2)" -BackgroundColor Black -ForegroundColor Yellow
	}
	Write-Host "$(Get-Date): Folder             : $($Folder)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): HW Inventory       : $($Hardware)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Log                : $($Log)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Logging            : $($Logging)" -BackgroundColor Black -ForegroundColor Yellow
	If($Logging)
	{
		Write-Host "$(Get-Date):    Start Date      : $($StartDate)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date):    End Date        : $($EndDate)" -BackgroundColor Black -ForegroundColor Yellow
	}
	Write-Host "$(Get-Date): MaxDetail          : $($MaxDetails)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): NoADPolicies       : $($NoADPolicies)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): NoPolicies         : $($NoPolicies)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Policies           : $($Policies)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Save As PDF        : $($PDF)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Save As HTML       : $($HTML)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Save As TEXT       : $($TEXT)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Save As WORD       : $($MSWORD)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): ScriptInfo         : $($ScriptInfo)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Section            : $($Section)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Summary            : $($Summary)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Title              : $($Script:Title)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): User Name          : $($UserName)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): OS Detected        : $($Script:RunningOS)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): PoSH version       : $($Host.Version)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): PSCulture          : $($PSCulture)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): PSUICulture        : $($PSUICulture)" -BackgroundColor Black -ForegroundColor Yellow
	If($MSWORD -or $PDF)
	{
		Write-Host "$(Get-Date): Word language      : $($Script:WordLanguageValue)" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): Word version       : $($Script:WordProduct)" -BackgroundColor Black -ForegroundColor Yellow
	}
	Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Script start       : $($Script:StartTime)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
}

Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	if( $object )
	{
		If( ( Get-Member -Name $topLevel -InputObject $object ) )
		{
			If( ( Get-Member -Name $secondLevel -InputObject $object.$topLevel ) )
			{
				Return $True
			}
		}
	}
	Return $False
}

Function AbortScript
{
	If($MSWord -or $PDF)
	{
		$Script:Word.quit()
		Write-Host "$(Get-Date): System Cleanup" -BackgroundColor Black -ForegroundColor Yellow
		[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Script:Word) | Out-Null
		If(Test-Path variable:global:word)
		{
			Remove-Variable -Name word -Scope Global
		}
	}
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Host "$(Get-Date): Script has been aborted" -BackgroundColor Black -ForegroundColor Yellow
	$ErrorActionPreference = $SaveEAPreference
	Exit
}

Function OutputWarning
{
	Param([string] $txt)
	Write-Warning $txt
	If($MSWord -or $PDF)
	{
		WriteWordLine 0 1 $txt
		WriteWordLIne 0 0 ""
	}
	ElseIf($Text)
	{
		Line 1 $txt
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 0 1 $txt
		WriteHTMLLine 0 0 " "
	}
}

Function TranscriptLogging
{
	If($Log) 
	{
		try 
		{
			If($Script:StartLog -eq $false)
			{
				Start-Transcript -Path $Script:LogPath -Force -Verbose:$false | Out-Null
			}
			Else
			{
				Start-Transcript -Path $Script:LogPath -Append -Verbose:$false | Out-Null
			}
			Write-Host "$(Get-Date): Transcript/log started at $Script:LogPath" -BackgroundColor Black -ForegroundColor Yellow
			$Script:StartLog = $true
		} 
		catch 
		{
			Write-Host "$(Get-Date): Transcript/log failed at $Script:LogPath" -BackgroundColor Black -ForegroundColor Yellow
			$Script:StartLog = $false
		}
	}
}

Function Get-IPAddress
{
	Param([string]$ComputerName)
	
	$IPAddress = "Unable to determine"
	
	Try
	{
		$IP = Test-Connection -ComputerName $ComputerName -Count 1 | Select-Object IPV4Address
	}
	
	Catch
	{
		$IP = $Null
	}

	If($? -and $Null -ne $IP)
	{
		$IPAddress = $IP.IPV4Address.IPAddressToString
	}
	
	Return $IPAddress
}
#endregion

#region script setup function
Function ProcessScriptSetup
{
	$script:startTime = Get-Date

	<#If(!(Check-NeededPSSnapins "Citrix.Common.Commands","Citrix.XenApp.Commands"))
	{
		#We're missing Citrix Snapins that we need
		$ErrorActionPreference = $SaveEAPreference
		Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Are you sure you are running this script on a XenApp 6.5 Server? Script will now close."
		Exit
	}
	#>
	Add-PSSnapin Citrix.Common.Commands -EA 0
	Add-PSSnapin Citrix.XenApp.Commands -EA 0
	
	$Script:DoPolicies = $True
	If($NoPolicies)
	{
		Write-Host "$(Get-Date): NoPolicies was specified so do not search for Citrix.GroupPolicy.Commands.psm1" -BackgroundColor Black -ForegroundColor Yellow
		$Script:DoPolicies = $False
	}
	ElseIf(!(Check-LoadedModule "Citrix.GroupPolicy.Commands") -and $Policies -eq $False)
	{
		Write-Warning "The Citrix Group Policy module Citrix.GroupPolicy.Commands.psm1 could not be loaded `n
		Please see the Prerequisites section in the ReadMe file (https://carlwebster.sharefile.com/d-s8e92231489542428). 
		`nCitrix Policy documentation will not take place"
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
		$Script:DoPolicies = $False
	}
	ElseIf(!(Check-LoadedModule "Citrix.GroupPolicy.Commands") -and $Policies -eq $True)
	{
		Write-Error "The Citrix Group Policy module Citrix.GroupPolicy.Commands.psm1 could not be loaded 
		`nPlease see the Prerequisites section in the ReadMe file (https://carlwebster.sharefile.com/d-s8e92231489542428). 
		`n
		`n
		`t`tBecause the Policies parameter was used the script will now close.
		`n
		`n"
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
		Exit
	}
	
	If($Policies -eq $False -and $NoPolicies -eq $False -and $NoADPolicies -eq $False)
	{
		#script defaults, so don't process policies
		$Script:DoPolicies = $False
	}
	If($NoPolicies -eq $True)
	{
		#don't process policies
		$Script:DoPolicies = $False
	}

	#if software inventory is specified then verify SoftwareExclusions.txt exists
	If($Software)
	{
		If(!(Test-Path "$($pwd.path)\SoftwareExclusions.txt"))
		{
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "Software inventory requested but $($pwd.path)\SoftwareExclusions.txt does not exist.  Script cannot continue."
			Exit
		}
		
		#file does exist but can we access it?
		$x = Get-Content "$($pwd.path)\SoftwareExclusions.txt" -EA 0
		If(!($?))
		{
			$ErrorActionPreference = $SaveEAPreference
			Write-Error "There was an error accessing or reading $($pwd.path)\SoftwareExclusions.txt.  Script cannot continue."
			Exit
		}
		$x = $Null
	}

	[bool]$Script:Remoting = $False
	[string]$Script:RemoteXAServer = ""
	
	If([String]::IsNullOrEmpty($AdminAddress))
	{
		#see if a remote connection was establish outside of the script
		$Script:RemoteXAServer = Get-XADefaultComputerName -EA 0
		If(![String]::IsNullOrEmpty($Script:RemoteXAServer))
		{
			$Script:Remoting = $True
		}
	}
	ElseIf($AdminAddress -ne "LocalHost")
	{
		#do nothing if $AdminAddress is LocalHost
		Set-XADefaultComputerName $AdminAddress -Scope LocalMachine -EA 0
		$Script:RemoteXAServer = Get-XADefaultComputerName -EA 0
		$Script:Remoting = $True
	}
	
	If($Script:Remoting)
	{
		Write-Host "$(Get-Date): Remoting is enabled to XenApp server $Script:RemoteXAServer" -BackgroundColor Black -ForegroundColor Yellow
		#now need to make sure the script is not being run against a session-only host
		$Server = Get-XAServer -ServerName $Script:RemoteXAServer -EA 0
		If($Server.ElectionPreference -eq "WorkerMode")
		{
			$ErrorActionPreference = $SaveEAPreference
			Write-Warning "This script cannot be run remotely against a Session-only Host Server."
			Write-Warning "Use Set-XADefaultComputerName XA65ControllerServerName or run the script on a controller."
			Write-Error "Script cannot continue.  See messages above."
			Exit
		}
	}
	Else
	{
		Write-Host "$(Get-Date): Remoting is not being used" -BackgroundColor Black -ForegroundColor Yellow
		
		#now need to make sure the script is not being run on a session-only host
		$ServerName = $env:computername
		$Server = Get-XAServer -ServerName $ServerName -EA 0
		If($Server.ElectionPreference -eq "WorkerMode")
		{
			$ErrorActionPreference = $SaveEAPreference
			Write-Warning "This script cannot be run on a Session-only Host Server if Remoting is not enabled."
			Write-Warning "Use Set-XADefaultComputerName XA65ControllerServerName or run the script on a controller."
			Write-Error "Script cannot continue.  See messages above."
			Exit
		}
	}

	# Get farm information
	Write-Host "$(Get-Date): Getting initial Farm data" -BackgroundColor Black -ForegroundColor Yellow
	$farm = Get-XAFarm -EA 0

	If($? -and $Null -ne $Farm)
	{
		Write-Host "$(Get-Date): Verify farm version" -BackgroundColor Black -ForegroundColor Yellow
		#first check to make sure this is a XenApp 6.5 farm
		If($Farm.ServerVersion.ToString().SubString(0,3) -eq "6.5")
		{
			#this is a XenApp 6.5 farm, script can proceed
		}
		Else
		{
			#this is not a XenApp 6.5 farm, script cannot proceed
			Write-Warning "This script is designed for XenApp 6.5 and should not be run on previous versions of XenApp"
			Return 1
		}
		[string]$Script:FarmName = $farm.FarmName
		[string]$Script:Title = "Inventory Report for the $($Script:FarmName) Farm"
	} 
	Else 
	{
		$ErrorActionPreference = $SaveEAPreference
		Write-Warning "Farm information could not be retrieved"
		If($Remoting)
		{
			Write-Error "A remote connection to $Script:RemoteXAServer could not be established.  Script cannot continue."
		}
		Else
		{
			Write-Error "Farm information could not be retrieved.  Script cannot continue."
		}
		Exit
	}
}
#endregion

#region configuration logging farm settings
Function ProcessConfigLogSettings
{
	If(!$Summary -and ($Section -eq "All" -or $Section -eq "ConfigLog"))
	{
		Write-Host "$(Get-Date): Processing Configuration Logging" -BackgroundColor Black -ForegroundColor Yellow
		[bool]$ConfigLog = $False
		$ConfigurationLogging = Get-XAConfigurationLog -EA 0

		If($? -and $Null -ne $ConfigurationLogging)
		{
			OutputConfigLogSettings $ConfigurationLogging
		}
		ElseIf($? -and $Null -eq $ConfigurationLogging) 
		{
			$txt = "No Configuration Logging settings"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Configuration Logging settings"
			OutputWarning $txt
		}
		Write-Host "$(Get-Date): Finished Configuration Logging" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputConfigLogSettings
{
	Param([object] $ConfigurationLogging )
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Configuration Logging"
	}
	ElseIf($Text)
	{
		Line 0 "Configuration Logging"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Configuration Logging"
	}
	
	If($ConfigurationLogging.LoggingEnabled) 
	{
		$Script:ConfigLog = $True
		[array]$ConString = $ConfigurationLogging.ConnectionString.Split(";")
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Configuration Logging"; Value = "Enabled"; }
			$ScriptInformation += @{ Data = "Allow changes to the farm when logging database is disconnected"; Value = $ConfigurationLogging.ChangesWhileDisconnectedAllowed; }
			$ScriptInformation += @{ Data = "Require administrator to enter credentials before clearing the log"; Value = $ConfigurationLogging.CredentialsOnClearLogRequired; }
			$ScriptInformation += @{ Data = "Database type"; Value = $ConfigurationLogging.DatabaseType; }
			$ScriptInformation += @{ Data = "Authentication mode"; Value = $ConfigurationLogging.AuthenticationMode; }
			$ScriptInformation += @{ Data = "Connection string"; Value = $ConString[0]; }
			$cnt = -1
			ForEach($tmp in $ConString)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$ScriptInformation += @{ Data = ""; Value = $tmp; }
				}
			}
			
			$ScriptInformation += @{ Data = "User name"; Value = $ConfigurationLogging.UserName; }
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 300;
			$Table.Columns.Item(2).Width = 200;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 "Configuration Logging`t`t`t`t`t`t`t: Enabled"
			Line 1 "Allow changes to the farm when logging database is disconnected`t`t: " $ConfigurationLogging.ChangesWhileDisconnectedAllowed
			Line 1 "Require administrator to enter credentials before clearing the log`t: " $ConfigurationLogging.CredentialsOnClearLogRequired
			Line 1 "Database type`t`t`t`t`t`t`t`t: " $ConfigurationLogging.DatabaseType
			Line 1 "Authentication mode`t`t`t`t`t`t`t: " $ConfigurationLogging.AuthenticationMode
			Line 1 "Connection string`t`t`t`t`t`t`t: " $ConString[0]
			$cnt = -1
			ForEach($tmp in $ConString)
			{
				$cnt++
				If($cnt -gt 0)
				{
					Line 10 "  " $tmp
				}
			}
			Line 1 "User name`t`t`t`t`t`t`t`t: " $ConfigurationLogging.UserName
			Line 0 ""
		}
		ElseIf($HTML)
		{
			$rowdata = @()
			$columnHeaders = @("Configuration Logging",($htmlsilver -bor $htmlbold),"Enabled",$htmlwhite)
			$rowdata += @(,('Allow changes to the farm when logging database is disconnected',($htmlsilver -bor $htmlbold),$ConfigurationLogging.ChangesWhileDisconnectedAllowed,$htmlwhite))	
			$rowdata += @(,('Require administrator to enter credentials before clearing the log',($htmlsilver -bor $htmlbold),$ConfigurationLogging.CredentialsOnClearLogRequired,$htmlwhite))	
			$rowdata += @(,('Database type',($htmlsilver -bor $htmlbold),$ConfigurationLogging.DatabaseType,$htmlwhite))	
			$rowdata += @(,('Authentication mode',($htmlsilver -bor $htmlbold),$ConfigurationLogging.AuthenticationMode,$htmlwhite))	
			$rowdata += @(,('Connection string',($htmlsilver -bor $htmlbold),$ConString[0],$htmlwhite))	
			$cnt = -1
			ForEach($tmp in $ConString)
			{
				$cnt++
				If($cnt -gt 0)
				{
					$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))	
				}
			}
			$rowdata += @(,('User name',($htmlsilver -bor $htmlbold),$ConfigurationLogging.UserName,$htmlwhite))	

			$msg = ""
			$columnWidths = @("300","200")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		}
	}
	Else 
	{
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 1 "Configuration Logging is disabled."
		}
		ElseIf($Text)
		{
			Line 1 "Configuration Logging is disabled."
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 0 1 "Configuration Logging is disabled."
			WriteHTMLLine 0 0 ""
		}
	}
}
#endregion

#region administrator functions
Function ProcessAdministrators
{
	If($Section -eq "All" -or $Section -eq "Admins")
	{
		Write-Host "$(Get-Date): Processing Administrators" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): `tSetting summary variables" -BackgroundColor Black -ForegroundColor Yellow
		[int]$Script:TotalFullAdmins = 0
		[int]$Script:TotalViewAdmins = 0
		[int]$Script:TotalCustomAdmins = 0
		[int]$Script:TotalAdmins = 0

		Write-Host "$(Get-Date): `tRetrieving Administrators" -BackgroundColor Black -ForegroundColor Yellow

		$Administrators = Get-XAAdministrator -EA 0| Sort-Object AdministratorName

		If($? -and $Null -ne $Administrators)
		{
			If($Summary)
			{
				OutputSummaryAdministrators $Administrators
			}
			Else
			{
				OutputAdministrators $Administrators
			}
		}
		ElseIf($? -and $Null -eq $Administrators)
		{
			$txt = "There are no Administrators"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Administrators"
			OutputWarning $txt
		}
		Write-Host "$(Get-Date): Finished Processing Administrators" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputSummaryAdministrators
{
	Param([object] $Administrators)

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Administrators"
		[System.Collections.Hashtable[]] $AdminsWordTable = @();
		[int] $CurrentServiceIndex = 2;
	}
	ElseIf($Text)
	{
		Line 0 "Administrators"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Administrators"
		$rowdata = @()
	}
	
	ForEach($Administrator in $Administrators)
	{
		Write-Host "$(Get-Date): `t`tProcessing administrator $($Administrator.AdministratorName)" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWord -or $PDF)
		{
			WriteWordLine 0 0 
			$WordTableRowHash = @{ 
			AdminName = $Administrator.AdministratorName;
			}
			$AdminsWordTable += $WordTableRowHash;
			$CurrentServiceIndex++;
		}
		ElseIf($Text)
		{
			Line 0 $Administrator.AdministratorName
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$Administrator.AdministratorName,$htmlwhite))
		}
		$Script:TotalAdmins++
	}
	
	If($MSWord -or $PDF)
	{
		$Table = AddWordTable -Hashtable $AdminsWordTable `
		-Columns AdminName `
		-Headers "Name" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold))
		
		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
	}
}

Function OutputAdministrators
{
	Param([object] $Administrators)

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Administrators"
	}
	ElseIf($Text)
	{
		Line 0 "Administrators"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Administrators"
	}
	
	ForEach($Administrator in $Administrators)
	{
		Write-Host "$(Get-Date): `t`tProcessing administrator $($Administrator.AdministratorName)" -BackgroundColor Black -ForegroundColor Yellow
		
		$xAdminType = ""
		Switch ($Administrator.AdministratorType)
		{
			"Unknown"  {$xAdminType = "Unknown"; Break}
			"Full"     {$xAdminType = "Full Administration"; $Script:TotalFullAdmins++; Break}
			"ViewOnly" {$xAdminType = "View Only"; $Script:TotalViewAdmins++; Break}
			"Custom"   {$xAdminType = "Custom"; $Script:TotalCustomAdmins++; Break}
			Default    {$xAdminType = "Administrator type could not be determined: $($Administrator.AdministratorType)"; Break}
		}
		
		$xAdminEnabled = ""
		If($Administrator.Enabled)
		{
			$xAdminEnabled = "Enabled" 
		} 
		Else
		{
			$xAdminEnabled = "Disabled" 
		}

		If($Administrator.AdministratorType -eq "Custom") 
		{
			$xFarmPrivileges = @()
			$xFolderPermissions = @()
			$AdministratorFarmPrivileges = $Administrator.FarmPrivileges
			ForEach($farmprivilege in $AdministratorFarmPrivileges) 
			{
				Write-Host "$(Get-Date): `t`t`tProcessing farm privilege $farmprivilege" -BackgroundColor Black -ForegroundColor Yellow
				Switch ($farmprivilege)
				{
					"Unknown"                   {$xFarmPrivileges += "Unknown"; Break}
					"ViewFarm"                  {$xFarmPrivileges += "View farm management"; Break}
					"EditZone"                  {$xFarmPrivileges += "Edit zones"; Break}
					"EditConfigurationLog"      {$xFarmPrivileges += "Configure logging for the farm"; Break}
					"EditFarmOther"             {$xFarmPrivileges += "Edit all other farm settings"; Break}
					"ViewAdmins"                {$xFarmPrivileges += "View Citrix administrators"; Break}
					"LogOnConsole"              {$xFarmPrivileges += "Log on to console"; Break}
					"LogOnWIConsole"            {$xFarmPrivileges += "Logon on to Web Interface console"; Break}
					"ViewLoadEvaluators"        {$xFarmPrivileges += "View load evaluators"; Break}
					"AssignLoadEvaluators"      {$xFarmPrivileges += "Assign load evaluators"; Break}
					"EditLoadEvaluators"        {$xFarmPrivileges += "Edit load evaluators"; Break}
					"ViewLoadBalancingPolicies" {$xFarmPrivileges += "View load balancing policies"; Break}
					"EditLoadBalancingPolicies" {$xFarmPrivileges += "Edit load balancing policies"; Break}
					"ViewPrinterDrivers"        {$xFarmPrivileges += "View printer drivers"; Break}
					"ReplicatePrinterDrivers"   {$xFarmPrivileges += "Replicate printer drivers"; Break}
					Default {$xFarmPrivileges += "Farm privileges could not be determined: $($farmprivilege)"; Break}
				}
			}
	
			Write-Host "$(Get-Date): `t`t`tProcessing folder privileges" -BackgroundColor Black -ForegroundColor Yellow
			$AdministratorFolderPrivileges = $Administrator.FolderPrivileges
			ForEach($folderprivilege in $AdministratorFolderPrivileges) 
			{
				#The Citrix PoSH cmdlet only returns data for three folders:
				#Servers
				#WorkerGroups
				#Applications
				
				Write-Host "$(Get-Date): `t`t`t`tProcessing folder permissions for $($FolderPrivilege.FolderPath)" -BackgroundColor Black -ForegroundColor Yellow
				$FolderPrivilegeFolderPrivileges = $FolderPrivilege.FolderPrivileges
				ForEach($FolderPermission in $FolderPrivilegeFolderPrivileges)
				{
					Switch ($folderpermission)
					{
						"Unknown"                          {$xFolderPermissions += "$($folderprivilege.FolderPath): Unknown"; Break}
						"ViewApplications"                 {$xFolderPermissions += "$($folderprivilege.FolderPath): View applications"; Break}
						"EditApplications"                 {$xFolderPermissions += "$($folderprivilege.FolderPath): Edit applications"; Break}
						"TerminateProcessApplication"      {$xFolderPermissions += "$($folderprivilege.FolderPath): Terminate process that is created as a result of launching a published application"; Break}
						"AssignApplicationsToServers"      {$xFolderPermissions += "$($folderprivilege.FolderPath): Assign applications to servers"; Break}
						"ViewServers"                      {$xFolderPermissions += "$($folderprivilege.FolderPath): View servers"; Break}
						"EditOtherServerSettings"          {$xFolderPermissions += "$($folderprivilege.FolderPath): Edit other server settings"; Break}
						"RemoveServer"                     {$xFolderPermissions += "$($folderprivilege.FolderPath): Remove a bad server from farm"; Break}
						"TerminateProcess"                 {$xFolderPermissions += "$($folderprivilege.FolderPath): Terminate processes on a server"; Break}
						"ViewSessions"                     {$xFolderPermissions += "$($folderprivilege.FolderPath): View ICA/RDP sessions"; Break}
						"ConnectSessions"                  {$xFolderPermissions += "$($folderprivilege.FolderPath): Connect sessions"; Break}
						"DisconnectSessions"               {$xFolderPermissions += "$($folderprivilege.FolderPath): Disconnect sessions"; Break}
						"LogOffSessions"                   {$xFolderPermissions += "$($folderprivilege.FolderPath): Log off sessions"; Break}
						"ResetSessions"                    {$xFolderPermissions += "$($folderprivilege.FolderPath): Reset sessions"; Break}
						"SendMessages"                     {$xFolderPermissions += "$($folderprivilege.FolderPath): Send messages to sessions"; Break}
						"ViewWorkerGroups"                 {$xFolderPermissions += "$($folderprivilege.FolderPath): View worker groups"; Break}
						"AssignApplicationsToWorkerGroups" {$xFolderPermissions += "$($folderprivilege.FolderPath): Assign applications to worker groups"; Break}
						Default {$xFolderPermissions += "Folder permission could not be determined: $($folderprivilege.FolderPath): $($folderpermission)"; Break}
					}
				}
			}
			$xFarmPrivileges = @($xFarmPrivileges | Sort-Object)
			$xFolderPermissions = @($xFolderPermissions | Sort-Object)
		}	
		
		If($MSWord -or $PDF)
		{
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Name"; Value = $Administrator.AdministratorName; }
			$ScriptInformation += @{ Data = "Type"; Value = $xAdminType; }
			$ScriptInformation += @{ Data = "Account"; Value = $xAdminEnabled; }

			If($Administrator.AdministratorType -eq "Custom") 
			{
				$ScriptInformation += @{ Data = "Privileges"; Value = $xFarmPrivileges[0]; }
				$cnt = -1
				ForEach($tmp in $xFarmPrivileges)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{ Data = ""; Value = $tmp; }
					}
				}
			
				$ScriptInformation += @{ Data = "Permissions"; Value = $xFolderPermissions[0]; }
				$cnt = -1
				ForEach($tmp in $xFolderPermissions)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$ScriptInformation += @{ Data = ""; Value = $tmp; }
					}
				}
			}
			
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 75;
			$Table.Columns.Item(2).Width = 425;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 1 "Name`t`t: " $Administrator.AdministratorName
			Line 1 "Type`t`t: " $xAdminType
			Line 1 "Account`t`t: " $xAdminEnabled
			If($Administrator.AdministratorType -eq "Custom") 
			{
				Line 1 "Privileges`t: " $xFarmPrivileges[0]
				$cnt = -1
				ForEach($tmp in $xFarmPrivileges)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 3 "  " $tmp
					}
				}
				Line 1 "Permissions`t: " $xFolderPermissions[0]
				$cnt = -1
				ForEach($tmp in $xFolderPermissions)
				{
					$cnt++
					If($cnt -gt 0)
					{
						Line 3 "  " $tmp
					}
				}
			}
			Line 0 ""				
		}
		ElseIf($HTML)
		{
			$columnHeaders = @()
			$rowdata = @()

			$columnHeaders = @("Name",($htmlsilver -bor $htmlbold),$Administrator.AdministratorName,$htmlwhite)
			$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$xAdminType,$htmlwhite))
			$rowdata += @(,('Account',($htmlsilver -bor $htmlbold),$xAdminEnabled,$htmlwhite))
			If($Administrator.AdministratorType -eq "Custom") 
			{
				$rowdata += @(,('Privileges',($htmlsilver -bor $htmlbold),$xFarmPrivileges[0],$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xFarmPrivileges)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					}
				}
				$rowdata += @(,('Permissions',($htmlsilver -bor $htmlbold),$xFolderPermissions[0],$htmlwhite))
				$cnt = -1
				ForEach($tmp in $xFolderPermissions)
				{
					$cnt++
					If($cnt -gt 0)
					{
						$rowdata += @(,('',($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					}
				}
			}
			$msg = ""
			$columnWidths = @("75","425")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			WriteHTMLLine 0 0 ""
		}
	}
	
}
#endregion

#region application functions
Function ProcessApplications
{
	If($Section -eq "All" -or $Section -eq "Apps")
	{
		Write-Host "$(Get-Date): Processing Applications" -BackgroundColor Black -ForegroundColor Yellow

		Write-Host "$(Get-Date): `tRetrieving Applications" -BackgroundColor Black -ForegroundColor Yellow
		If($Summary)
		{
			$Applications = Get-XAApplication -EA 0 | Sort-Object DisplayName
		}
		Else
		{
			$Applications = Get-XAApplication -EA 0 | Sort-Object FolderPath, DisplayName
		}

		If($? -and $Null -ne $Applications)
		{
			OutputApplications $Applications
		}
		ElseIf($Null -eq $Applications)
		{
			Write-Host "$(Get-Date): There are no Applications published" -BackgroundColor Black -ForegroundColor Yellow
		}
		Else 
		{
			Write-Warning "Application information could not be retrieved."
		}
		$Applications = $Null
		Write-Host "$(Get-Date): Finished Processing Applications" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputApplications
{
	Param([object] $Applications)

	Write-Host "$(Get-Date): `tSetting summary variables" -BackgroundColor Black -ForegroundColor Yellow
	[int]$Script:TotalPublishedApps = 0
	[int]$Script:TotalPublishedContent = 0
	[int]$Script:TotalPublishedDesktops = 0
	[int]$Script:TotalStreamedApps = 0
	[int]$Script:TotalApps = 0
	$Script:SessionSharingItems = @()

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Applications"
	}
	ElseIf($Text)
	{
		Line 0 "Applications"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Applications"
	}

	ForEach($Application in $Applications)
	{
		Write-Host "$(Get-Date): `t`tProcessing application $($Application.BrowserName)" -BackgroundColor Black -ForegroundColor Yellow
		
		If(!$Summary)
		{
			If($Application.ApplicationType -ne "ServerDesktop" -and $Application.ApplicationType -ne "Content")
			{
				#create array for appendix A
				#these items are taken from http://support.citrix.com/article/CTX159159
				#Some properties that must match on both Applications for Session Sharing to Function are:
				#
				#Color depth
				#Screen Size
				#Access Control Filters (for SmartAccess)
				#Sound (unexplained in article)
				#Drive Mapping (unexplained in article)
				#Printer Mapping (unexplained in article)
				#Encryption
				
				Write-Host "$(Get-Date): `t`t`tGather session sharing info for Appendix A" -BackgroundColor Black -ForegroundColor Yellow
				$obj = New-Object -TypeName PSObject
				$obj | Add-Member -MemberType NoteProperty -Name ApplicationName      -Value $Application.BrowserName
				$obj | Add-Member -MemberType NoteProperty -Name MaximumColorQuality  -Value $Application.ColorDepth
				$obj | Add-Member -MemberType NoteProperty -Name SessionWindowSize    -Value $Application.WindowType

				If($Application.AccessSessionConditionsEnabled)
				{
					$tmp = @()
					ForEach($filter in $Application.AccessSessionConditions)
					{
						$tmp += $filter
					}
					$obj | Add-Member -MemberType NoteProperty -Name AccessControlFilters -Value $tmp
				}
				Else
				{
					$obj | Add-Member -MemberType NoteProperty -Name AccessControlFilters -Value "None"
				}
				$tmp = $Null
				
				$obj | Add-Member -MemberType NoteProperty -Name Encryption           -Value $Application.EncryptionLevel
				$Script:SessionSharingItems += $obj
			}
			$AppServerInfoResults = $False
			$AppServerInfo = Get-XAApplicationReport -BrowserName $Application.BrowserName -EA 0
			If($? -and $Null -ne $AppServerInfo)
			{
				$AppServerInfoResults = $True
			}
			[bool]$streamedapp = $False
			If($Application.ApplicationType -Contains "streamedtoclient" -or $Application.ApplicationType -Contains "streamedtoserver")
			{
				$streamedapp = $True
			}
		}
		Else
		{
			$Script:TotalApps++
		}
		
		#name properties
		If(!$Summary)
		{
			#weird, if application is enabled, it is disabled!
			If($Application.Enabled) 
			{
				$ApplicationEnabled = "No"
			} 
			Else
			{
				$ApplicationEnabled = "Yes"
				If($Application.HideWhenDisabled)
				{
					$ApplicationHideWhenDisabled = "Yes"
				}
				Else
				{
					$ApplicationHideWhenDisabled = "No"
				}
			}
			
			If($MSWord -or $PDF)
			{
				WriteWordLine 2 0 $Application.DisplayName
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Application name"; Value = $Application.BrowserName; }
				$ScriptInformation += @{ Data = "Disable application"; Value = $ApplicationEnabled; }
				#weird, if application is enabled, it is disabled!
				If($ApplicationEnabled -eq "Yes") 
				{
					$ScriptInformation += @{ Data = "Hide disabled application"; Value = $ApplicationHideWhenDisabled; }
				}

				$ScriptInformation += @{ Data = "Application description"; Value = $Application.Description; }
				
				#type properties
				$tmp = ""
				Switch ($Application.ApplicationType)
				{
					"Unknown"                            	{$Tmp = "Unknown"; Break}
					"ServerInstalled"                    	{$Tmp = "Installed application"; $Script:TotalPublishedApps++; Break}
					"ServerDesktop"                      	{$Tmp = "Server desktop"; $Script:TotalPublishedDesktops++; Break}
					"Content"                            	{$Tmp = "Content"; $Script:TotalPublishedContent++; Break}
					"StreamedToServer"                   	{$Tmp = "Streamed to server"; $Script:TotalStreamedApps++; Break}
					"StreamedToClient"                   	{$Tmp = "Streamed to client"; $Script:TotalStreamedApps++; Break}
					"StreamedToClientOrInstalled"        	{$Tmp = "Streamed if possible, otherwise accessed from server as Installed application"; $Script:TotalStreamedApps++; Break}
					"StreamedToClientOrStreamedToServer" 	{$Tmp = "Streamed if possible, otherwise Streamed to server"; $Script:TotalStreamedApps++; Break}
					Default									{$Tmp = "Application Type could not be determined: $($Application.ApplicationType)"; Break}
				}
				$ScriptInformation += @{ Data = "Application Type"; Value = $tmp; }
				$ScriptInformation += @{ Data = "Folder path"; Value = $Application.FolderPath; }
				$ScriptInformation += @{ Data = "Content Address"; Value = $Application.ContentAddress; }
				$tmp = $Null
				
				#if a streamed app
				If($streamedapp)
				{
					$ScriptInformation += @{ Data = "Citrix streaming app profile address"; Value = $Application.ProfileLocation; }
					$ScriptInformation += @{ Data = "App to launch from Citrix stream app profile"; Value = $Application.ProfileProgramName; }
					$ScriptInformation += @{ Data = "Extra command line parameters"; Value = $Application.ProfileProgramArguments; }

					#if streamed, OffWriteWordLine 0 access properties
					If($Application.OfflineAccessAllowed)
					{
						If($Application.OfflineAccessAllowed)
						{
							$tmp = "Yes"
						}
						Else
						{
							$tmp = "No"
						}
						$ScriptInformation += @{ Data = "Enable offline access"; Value = $tmp; }
						$tmp = $Null
					}
					If($Application.CachingOption)
					{
						Switch ($Application.CachingOption)
						{
							"Unknown"   {$Tmp = "Unknown"; Break}
							"PreLaunch" {$Tmp = "Cache application prior to launching"; Break}
							"AtLaunch"  {$Tmp = "Cache application during launch"; Break}
							Default		{$Tmp = "Could not be determined: $($Application.CachingOption)"; Break}
						}
						$ScriptInformation += @{ Data = "Cache preference"; Value = $tmp; }
						$tmp = $Null
					}
				}
				
				#location properties
				If(!$streamedapp)
				{
					$ScriptInformation += @{ Data = "Command Line"; Value = $Application.CommandLineExecutable; }
					$ScriptInformation += @{ Data = "Working directory"; Value = $Application.WorkingDirectory; }
					
					#servers properties
					If($AppServerInfoResults)
					{
						If(![String]::IsNullOrEmpty($AppServerInfo.ServerNames))
						{
							$TempArray = @($AppServerInfo.ServerNames | Sort-Object ServerNames)
							$ScriptInformation += @{ Data = "Servers"; Value = $TempArray[0]; }
							$cnt = -1
							ForEach($Item in $TempArray)
							{
								$cnt++
								If($cnt -gt 0)
								{
									$ScriptInformation += @{ Data = ""; Value = $Item; }
								}
							}
							$TempArray = $Null
						}
						If(![String]::IsNullOrEmpty($AppServerInfo.WorkerGroupNames))
						{
							$TempArray = @($AppServerInfo.WorkerGroupNames | Sort-Object WorkerGroupNames)
							$ScriptInformation += @{ Data = "Worker Groups"; Value = $TempArray[0]; }
							$cnt = -1
							ForEach($Item in $TempArray)
							{
								$cnt++
								If($cnt -gt 0)
								{
									$ScriptInformation += @{ Data = ""; Value = $Item; }
								}
							}
							$TempArray = $Null
						}
					}
					Else
					{
						$ScriptInformation += @{ Data = "Unable to retrieve a list of Servers or Worker Groups for this application"; Value = ""; }
					}
				}
			
				#users properties
				If($Application.AnonymousConnectionsAllowed)
				{
					$ScriptInformation += @{ Data = "Allow anonymous users"; Value = $Application.AnonymousConnectionsAllowed; }
				}
				Else
				{
					If($AppServerInfoResults)
					{
						$TempArray = @($AppServerInfo.Accounts | Sort-Object AccountName)
						$ScriptInformation += @{ Data = "Users"; Value = $TempArray[0]; }
						$cnt = -1
						ForEach($Item in $TempArray)
						{
							$cnt++
							If($cnt -gt 0)
							{
								$ScriptInformation += @{ Data = ""; Value = $Item; }
							}
						}
						$TempArray = $Null
					}
					Else
					{
						$ScriptInformation += @{ Data = "Unable to retrieve a list of Users for this application"; Value = ""; }
					}
				}	

				#shortcut presentation properties
				#application icon is ignored
				$ScriptInformation += @{ Data = "Client application folder"; Value = $Application.ClientFolder; }
				If($Application.AddToClientStartMenu)
				{
					$ScriptInformation += @{ Data = "Add to client's start menu"; Value = ""; }
					If($Application.StartMenuFolder)
					{
						$ScriptInformation += @{ Data = "Start menu folder"; Value = $Application.StartMenuFolder; }
					}
				}
				If($Application.AddToClientDesktop)
				{
					$ScriptInformation += @{ Data = "Add shortcut to the client's desktop"; Value = ""; }
				}
			
				#access control properties
				If($Application.ConnectionsThroughAccessGatewayAllowed)
				{
					If($Application.ConnectionsThroughAccessGatewayAllowed)
					{
						$tmp = "Yes"
					} 
					Else
					{
						$tmp = "No"
					}
					$ScriptInformation += @{ Data = "Allow connections made through AGAE"; Value = $tmp; }
					$tmp = $Null
				}
				If($Application.OtherConnectionsAllowed)
				{
					If($Application.OtherConnectionsAllowed)
					{
						$tmp = "Yes"
					} 
					Else
					{
						$tmp = "No"
					}
					$ScriptInformation += @{ Data = "Any connection"; Value = $tmp; }
					$tmp = $Null
				}
				If($Application.AccessSessionConditionsEnabled)
				{
					$ScriptInformation += @{ Data = "Any connection that meets any of the following filters"; Value = $Application.AccessSessionConditionsEnabled; }
					$ScriptInformation += @{ Data = "     Access Gateway Filters"; Value = ""; }
					ForEach($AccessCondition in $Application.AccessSessionConditions)
					{
						[string]$Tmp = $AccessCondition
						[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
						[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
						$tmp = "Farm name: $AGFarm  Filter: $AGFilter"
						$ScriptInformation += @{ Data = ""; Value = $tmp; }
					}

					$tmp = $Null
					$AGFarm = $Null
					$AGFilter = $Null
				}
			
				#content redirection properties
				If($AppServerInfoResults)
				{
					If($AppServerInfo.FileTypes)
					{
						$TempArray = $AppServerInfo.FileTypes | Sort-Object FileTypes
						$ScriptInformation += @{ Data = "File type associations"; Value = $TempArray[0]; }
						$cnt = -1
						ForEach($Item in $TempArray)
						{
							$cnt++
							If($cnt -gt 0)
							{
								$ScriptInformation += @{ Data = ""; Value = $Item; }
							}
						}
						$TempArray = $Null
					}
					Else
					{
						$ScriptInformation += @{ Data = "File Type Associations for this application"; Value = "None"; }
					}
				}
				Else
				{
					$ScriptInformation += @{ Data = "Unable to retrieve the list of FTAs for this application"; Value = ""; }
				}
			
				#if streamed app, Alternate profiles
				If($streamedapp)
				{
					If($Application.AlternateProfiles)
					{
						$ScriptInformation += @{ Data = "Primary application profile location"; Value = $Application.AlternateProfiles; }
					}
				
					#if streamed app, User privileges properties
					If($Application.RunAsLeastPrivilegedUser)
					{
						$ScriptInformation += @{ Data = "Run app as a least-privileged user account"; Value = $Application.RunAsLeastPrivilegedUser; }
					}
				}
			
				#limits properties
				If($Application.InstanceLimit -eq -1)
				{
					$tmp = "No limit set"
				}
				Else
				{
					$tmp = $Application.InstanceLimit.ToString()
				}
				$ScriptInformation += @{ Data = "Limit instances allowed to run in server farm"; Value = $tmp; }
				$tmp = $Null

			
				If($Application.MultipleInstancesPerUserAllowed) 
				{
					$tmp = "No"
				} 
				Else
				{
					$tmp = "Yes"
				}
				$ScriptInformation += @{ Data = "Allow only 1 instance of app for each user"; Value = $tmp; }
				$tmp = $Null
			
			
				If($Application.CpuPriorityLevel)
				{
					Switch ($Application.CpuPriorityLevel)
					{
						"Unknown"     	{$Tmp = "Unknown"; Break}
						"BelowNormal" 	{$Tmp = "Below Normal"; Break}
						"Low"         	{$Tmp = "Low"; Break}
						"Normal"      	{$Tmp = "Normal"; Break}
						"AboveNormal" 	{$Tmp = "Above Normal"; Break}
						"High"        	{$Tmp = "High"; Break}
						Default			{$Tmp = "Application importance could not be determined: $($Application.CpuPriorityLevel)"; Break}
					}
					$ScriptInformation += @{ Data = "Application importance"; Value = $tmp; }
				}
				
				#client options properties
				Switch ($Application.AudioType)
				{
					"Unknown" 	{$Tmp = "Unknown"; Break}
					"None"    	{$Tmp = "Not Enabled"; Break}
					"Basic"   	{$Tmp = "Enabled"; Break}
					Default		{$Tmp = "Enable legacy audio could not be determined: $($Application.AudioType)"; Break}
				}
				$ScriptInformation += @{ Data = "Enable legacy audio"; Value = $tmp; }
				$tmp = $Null

				If($Application.AudioRequired)
				{
					$tmp = "Enabled"
				}
				Else
				{
					$tmp = "Disabled"
				}
				$ScriptInformation += @{ Data = "Minimum requirement"; Value = $tmp; }
				$tmp = $Null

				If($Application.SslConnectionEnabled)
				{
					If($Application.SslConnectionEnabled)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
				}
				$ScriptInformation += @{ Data = "Enable SSL and TLS protocols"; Value = $tmp; }
				$tmp = $Null

				If($Application.EncryptionLevel)
				{
					Switch ($Application.EncryptionLevel)
					{
						"Unknown" 	{$Tmp = "Unknown"; Break}
						"Basic"   	{$Tmp = "Basic"; Break}
						"LogOn"   	{$Tmp = "128-Bit Login Only (RC-5)"; Break}
						"Bits40"  	{$Tmp = "40-Bit (RC-5)"; Break}
						"Bits56"  	{$Tmp = "56-Bit (RC-5)"; Break}
						"Bits128" 	{$Tmp = "128-Bit (RC-5)"; Break}
						Default		{$Tmp = "Encryption could not be determined: $($Application.EncryptionLevel)"; Break}
					}
				}
				$ScriptInformation += @{ Data = "Encryption"; Value = $tmp; }
				$tmp = $Null
				
				If($Application.EncryptionRequired)
				{
					If($Application.EncryptionRequired)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
				}
				$ScriptInformation += @{ Data = "Minimum requirement"; Value = $tmp; }
				$tmp = $Null
			
				#another weird one, if True then this is Disabled
				If($Application.WaitOnPrinterCreation) 
				{
					$Tmp = "Disabled"
				} 
				Else
				{
					$Tmp = "Enabled"
				}
				$ScriptInformation += @{ Data = "Start app w/o waiting for printer creation"; Value = $tmp; }
				$tmp = $Null
				
				#appearance properties
				If($Application.WindowType)
				{
					$ScriptInformation += @{ Data = "Session window size"; Value = $Application.WindowType; }
				}
				If($Application.ColorDepth)
				{
					Switch ($Application.ColorDepth)
					{
						"Unknown"     	{$Tmp = "Unknown color depth"; Break}
						"Colors8Bit"  	{$Tmp = "256-color (8-bit)"; Break}
						"Colors16Bit" 	{$Tmp = "Better Speed (16-bit)"; Break}
						"Colors32Bit" 	{$Tmp = "Better Appearance (32-bit)"; Break}
						Default			{$Tmp = "Maximum color quality could not be determined: $($Application.ColorDepth)"; Break}
					}
					$ScriptInformation += @{ Data = "Maximum color quality"; Value = $tmp; }
					$tmp = $Null
				}
				If($Application.TitleBarHidden)
				{
					If($Application.TitleBarHidden)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
					$ScriptInformation += @{ Data = "Hide application title bar"; Value = $tmp; }
					$tmp = $Null
				}
				If($Application.MaximizedOnStartup)
				{
					If($Application.MaximizedOnStartup)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
					$ScriptInformation += @{ Data = "Maximize application at startup"; Value = $tmp; }
					$tmp = $Null
				}
				$AppServerInfo = $Null
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				## IB - Set the header row format
				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 0 $Application.DisplayName
				Line 1 "Application name`t`t: " $Application.BrowserName
				Line 1 "Disable application`t`t: " -NoNewLine
				#weird, if application is enabled, it is disabled!
				If($Application.Enabled) 
				{
					Line 0 "No"
				} 
				Else
				{
					Line 0 "Yes"
					Line 1 "Hide disabled application`t: " -nonewline
					If($Application.HideWhenDisabled)
					{
						Line 0 "Yes"
					}
					Else
					{
						Line 0 "No"
					}
				}

				If(![String]::IsNullOrEmpty($Application.Description))
				{
					Line 1 "Application description`t`t: " $Application.Description
				}
				
				#type properties
				Line 1 "Application Type`t`t: " -nonewline
				Switch ($Application.ApplicationType)
				{
					"Unknown"                            	{Line 0 "Unknown"; Break}
					"ServerInstalled"                    	{Line 0 "Installed application"; $Script:TotalPublishedApps++; Break}
					"ServerDesktop"                      	{Line 0 "Server desktop"; $Script:TotalPublishedDesktops++; Break}
					"Content"                            	{Line 0 "Content"; $Script:TotalPublishedContent++; Break}
					"StreamedToServer"                   	{Line 0 "Streamed to server"; $Script:TotalStreamedApps++; Break}
					"StreamedToClient"                   	{Line 0 "Streamed to client"; $Script:TotalStreamedApps++; Break}
					"StreamedToClientOrInstalled"        	{Line 0 "Streamed if possible, otherwise accessed from server as Installed application"; $Script:TotalStreamedApps++; Break}
					"StreamedToClientOrStreamedToServer" 	{Line 0 "Streamed if possible, otherwise Streamed to server"; $Script:TotalStreamedApps++; Break}
					Default 								{Line 0 "Application Type could not be determined: $($Application.ApplicationType)"; Break}
				}
				If(![String]::IsNullOrEmpty($Application.FolderPath))
				{
					Line 1 "Folder path`t`t`t: " $Application.FolderPath
				}
				If(![String]::IsNullOrEmpty($Application.ContentAddress))
				{1
				
					Line 1 "Content Address`t`t: " $Application.ContentAddress
				}
			
				#if a streamed app
				If($streamedapp)
				{
					Line 1 "Citrix streaming app profile address`t`t: " $Application.ProfileLocation
					Line 1 "App to launch from Citrix stream app profile`t: " $Application.ProfileProgramName
					If(![String]::IsNullOrEmpty($Application.ProfileProgramArguments))
					{
						Line 1 "Extra command line parameters`t`t`t: " $Application.ProfileProgramArguments
					}
					#if streamed, OffLine access properties
					If($Application.OfflineAccessAllowed)
					{
						Line 1 "Enable offline access`t`t`t`t: " -nonewline
						If($Application.OfflineAccessAllowed)
						{
							Line 0 "Yes"
						}
						Else
						{
							Line 0 "No"
						}
					}
					If($Application.CachingOption)
					{
						Line 1 "Cache preference`t`t`t`t: " -nonewline
						Switch ($Application.CachingOption)
						{
							"Unknown"   {Line 0 "Unknown"; Break}
							"PreLaunch" {Line 0 "Cache application prior to launching"; Break}
							"AtLaunch"  {Line 0 "Cache application during launch"; Break}
							Default		{Line 0 "Could not be determined: $($Application.CachingOption)"; Break}
						}
					}
				}
				
				#location properties
				If(!$streamedapp)
				{
					#requested by Pavel Stadler to put Command Line and Working Directory in a different sized font and make it bold
					If(![String]::IsNullOrEmpty($Application.CommandLineExecutable))
					{
						Line 1 "Command Line`t`t`t: " $Application.CommandLineExecutable 
					}
					If(![String]::IsNullOrEmpty($Application.WorkingDirectory))
					{
						Line 1 "Working directory`t`t: " $Application.WorkingDirectory
					}
					
					#servers properties
					If($AppServerInfoResults)
					{
						If(![String]::IsNullOrEmpty($AppServerInfo.ServerNames))
						{
							Line 1 "Servers:"
							$TempArray = $AppServerInfo.ServerNames | Sort-Object ServerNames
							ForEach($Item in $TempArray)
							{
								Line 2 $Item
							}
							$TempArray = $Null
						}
						If(![String]::IsNullOrEmpty($AppServerInfo.WorkerGroupNames))
						{
							Line 1 "Worker Groups:"
							$TempArray = $AppServerInfo.WorkerGroupNames | Sort-Object WorkerGroupNames
							ForEach($Item in $TempArray)
							{
								Line 2 $Item
							}
							$TempArray = $Null
						}
					}
					Else
					{
						Line 2 "Unable to retrieve a list of Servers or Worker Groups for this application"
					}
				}
			
				#users properties
				If($Application.AnonymousConnectionsAllowed)
				{
					Line 1 "Allow anonymous users: " $Application.AnonymousConnectionsAllowed
				}
				Else
				{
					If($AppServerInfoResults)
					{
						$TempArray = @($AppServerInfo.Accounts | Sort-Object Accounts)
						Line 1 "Users:"
						ForEach($user in $TempArray)
						{
							Line 2 $user
						}
					}
					Else
					{
						Line 2 "Unable to retrieve a list of Users for this application"
					}
				}	

				#shortcut presentation properties
				#application icon is ignored
				If(![String]::IsNullOrEmpty($Application.ClientFolder))
				{
					Line 1 "Client application folder`t`t`t: " $Application.ClientFolder
				}
				If($Application.AddToClientStartMenu)
				{
					Line 1 "Add to client's start menu"
					If($Application.StartMenuFolder)
					{
						Line 2 "Start menu folder`t`t`t: " $Application.StartMenuFolder
					}
				}
				If($Application.AddToClientDesktop)
				{
					Line 1 "Add shortcut to the client's desktop"
				}
			
				#access control properties
				If($Application.ConnectionsThroughAccessGatewayAllowed)
				{
					Line 1 "Allow connections made through AGAE`t`t: " -nonewline
					If($Application.ConnectionsThroughAccessGatewayAllowed)
					{
						Line 0 "Yes"
					} 
					Else
					{
						Line 0 "No"
					}
				}
				If($Application.OtherConnectionsAllowed)
				{
					Line 1 "Any connection`t`t`t`t`t: " -nonewline
					If($Application.OtherConnectionsAllowed)
					{
						Line 0 "Yes"
					} 
					Else
					{
						Line 0 "No"
					}
				}
				If($Application.AccessSessionConditionsEnabled)
				{
					Line 1 "Any connection that meets any of the following filters: " $Application.AccessSessionConditionsEnabled
					Line 1 "Access Gateway Filters:"
					ForEach($AccessCondition in $Application.AccessSessionConditions)
					{
						[string]$Tmp = $AccessCondition
						[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
						[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
						Line 2 "$($AGFarm) $($AGFilter)"
					}
					Line 0 ""
					$tmp = $Null
					$AGFarm = $Null
					$AGFilter = $Null
				}
			
				#content redirection properties
				If($AppServerInfoResults)
				{
					If($AppServerInfo.FileTypes)
					{
						Line 1 "File type associations:"
						ForEach($filetype in $AppServerInfo.FileTypes)
						{
							Line 2 $filetype
						}
					}
					Else
					{
						Line 1 "File Type Associations for this application`t: None"
					}
				}
				Else
				{
					Line 1 "Unable to retrieve the list of FTAs for this application"
				}
			
				#if streamed app, Alternate profiles
				If($streamedapp)
				{
					If($Application.AlternateProfiles)
					{
						Line 1 "Primary application profile location`t`t: " $Application.AlternateProfiles
					}
				
					#if streamed app, User privileges properties
					If($Application.RunAsLeastPrivilegedUser)
					{
						Line 1 "Run app as a least-privileged user account`t: " $Application.RunAsLeastPrivilegedUser
					}
				}
			
				#limits properties
				Line 1 "Limit instances allowed to run in server farm`t: " -NoNewLine

				If($Application.InstanceLimit -eq -1)
				{
					Line 0 "No limit set"
				}
				Else
				{
					Line 0 $Application.InstanceLimit
				}
			
				Line 1 "Allow only 1 instance of app for each user`t: " -NoNewLine
			
				If($Application.MultipleInstancesPerUserAllowed) 
				{
					Line 0 "No"
				} 
				Else
				{
					Line 0 "Yes"
				}
			
				If($Application.CpuPriorityLevel)
				{
					Line 1 "Application importance`t`t`t`t: " -nonewline
					Switch ($Application.CpuPriorityLevel)
					{
						"Unknown"     	{Line 0 "Unknown"; Break}
						"BelowNormal" 	{Line 0 "Below Normal"; Break}
						"Low"         	{Line 0 "Low"; Break}
						"Normal"      	{Line 0 "Normal"; Break}
						"AboveNormal" 	{Line 0 "Above Normal"; Break}
						"High"        	{Line 0 "High"; Break}
						Default			{Line 0 "Application importance could not be determined: $($Application.CpuPriorityLevel)"; Break}
					}
				}
				
				#client options properties
				Line 1 "Enable legacy audio`t`t`t`t: " -nonewline
				Switch ($Application.AudioType)
				{
					"Unknown" 	{Line 0 "Unknown"; Break}
					"None"    	{Line 0 "Not Enabled"; Break}
					"Basic"   	{Line 0 "Enabled"; Break}
					Default		{Line 0 "Enable legacy audio could not be determined: $($Application.AudioType)"; Break}
				}
				Line 1 "Minimum requirement`t`t`t`t: " -nonewline
				If($Application.AudioRequired)
				{
					Line 0 "Enabled"
				}
				Else
				{
					Line 0 "Disabled"
				}
				If($Application.SslConnectionEnabled)
				{
					Line 1 "Enable SSL and TLS protocols`t`t`t: " -nonewline
					If($Application.SslConnectionEnabled)
					{
						Line 0 "Enabled"
					}
					Else
					{
						Line 0 "Disabled"
					}
				}
				If($Application.EncryptionLevel)
				{
					Line 1 "Encryption`t`t`t`t`t: " -nonewline
					Switch ($Application.EncryptionLevel)
					{
						"Unknown" 	{Line 0 "Unknown"; Break}
						"Basic"   	{Line 0 "Basic"; Break}
						"LogOn"   	{Line 0 "128-Bit Login Only (RC-5)"; Break}
						"Bits40"  	{Line 0 "40-Bit (RC-5)"; Break}
						"Bits56"  	{Line 0 "56-Bit (RC-5)"; Break}
						"Bits128" 	{Line 0 "128-Bit (RC-5)"; Break}
						Default		{Line 0 "Encryption could not be determined: $($Application.EncryptionLevel)"; Break}
					}
				}
				If($Application.EncryptionRequired)
				{
					Line 1 "Minimum requirement`t`t`t`t: " -nonewline
					If($Application.EncryptionRequired)
					{
						Line 0 "Enabled"
					}
					Else
					{
						Line 0 "Disabled"
					}
				}
			
				Line 1 "Start app w/o waiting for printer creation`t: " -NoNewLine
				#another weird one, if True then this is Disabled
				If($Application.WaitOnPrinterCreation) 
				{
					Line 0 "Disabled"
				} 
				Else
				{
					Line 0 "Enabled"
				}
				
				#appearance properties
				If($Application.WindowType)
				{
					Line 1 "Session window size`t`t`t`t: " $Application.WindowType
				}
				If($Application.ColorDepth)
				{
					Line 1 "Maximum color quality`t`t`t`t: " -nonewline
					Switch ($Application.ColorDepth)
					{
						"Unknown"     	{Line 0 "Unknown color depth"; Break}
						"Colors8Bit"  	{Line 0 "256-color (8-bit)"; Break}
						"Colors16Bit" 	{Line 0 "Better Speed (16-bit)"; Break}
						"Colors32Bit" 	{Line 0 "Better Appearance (32-bit)"; Break}
						Default			{Line 0 "Maximum color quality could not be determined: $($Application.ColorDepth)"; Break}
					}
				}
				If($Application.TitleBarHidden)
				{
					Line 1 "Hide application title bar`t`t`t: " -nonewline
					If($Application.TitleBarHidden)
					{
						Line 0 "Enabled"
					}
					Else
					{
						Line 0 "Disabled"
					}
				}
				If($Application.MaximizedOnStartup)
				{
					Line 1 "Maximize application at startup`t`t`t: " -nonewline
					If($Application.MaximizedOnStartup)
					{
						Line 0 "Enabled"
					}
					Else
					{
						Line 0 "Disabled"
					}
				}
				Line 0 ""
				$AppServerInfo = $Null
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 $Application.DisplayName
				$columnHeaders = @()
				$rowdata = @()

				$columnHeaders = @("Application name",($htmlsilver -bor $htmlbold),$Application.BrowserName,$htmlwhite)
				$rowdata += @(,("Disable application",($htmlsilver -bor $htmlbold),$ApplicationEnabled,$htmlwhite))
				#weird, if application is enabled, it is disabled!
				If($ApplicationEnabled -eq "Yes") 
				{
					$rowdata += @(,("Hide disabled application",($htmlsilver -bor $htmlbold),$ApplicationHideWhenDisabled,$htmlwhite))
				}

				$rowdata += @(,("Application description",($htmlsilver -bor $htmlbold),$Application.Description,$htmlwhite))
				
				#type properties
				$tmp = ""
				Switch ($Application.ApplicationType)
				{
					"Unknown"                            	{$Tmp = "Unknown"; Break}
					"ServerInstalled"                    	{$Tmp = "Installed application"; $Script:TotalPublishedApps++; Break}
					"ServerDesktop"                      	{$Tmp = "Server desktop"; $Script:TotalPublishedDesktops++; Break}
					"Content"                            	{$Tmp = "Content"; $Script:TotalPublishedContent++; Break}
					"StreamedToServer"                   	{$Tmp = "Streamed to server"; $Script:TotalStreamedApps++; Break}
					"StreamedToClient"                   	{$Tmp = "Streamed to client"; $Script:TotalStreamedApps++; Break}
					"StreamedToClientOrInstalled"        	{$Tmp = "Streamed if possible, otherwise accessed from server as Installed application"; $Script:TotalStreamedApps++; Break}
					"StreamedToClientOrStreamedToServer" 	{$Tmp = "Streamed if possible, otherwise Streamed to server"; $Script:TotalStreamedApps++; Break}
					Default									{$Tmp = "Application Type could not be determined: $($Application.ApplicationType)"; Break}
				}
				$rowdata += @(,("Application Type",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$rowdata += @(,("Folder path",($htmlsilver -bor $htmlbold),$Application.FolderPath,$htmlwhite))
				$rowdata += @(,("Content Address",($htmlsilver -bor $htmlbold),$Application.ContentAddress,$htmlwhite))
				$tmp = $Null
				
				#if a streamed app
				If($streamedapp)
				{
					$rowdata += @(,("Citrix streaming app profile address",($htmlsilver -bor $htmlbold),$Application.ProfileLocation,$htmlwhite))
					$rowdata += @(,("App to launch from Citrix stream app profile",($htmlsilver -bor $htmlbold),$Application.ProfileProgramName,$htmlwhite))
					$rowdata += @(,("Extra command line parameters",($htmlsilver -bor $htmlbold),$Application.ProfileProgramArguments,$htmlwhite))

					#if streamed, OffWriteWordLine 0 access properties
					If($Application.OfflineAccessAllowed)
					{
						If($Application.OfflineAccessAllowed)
						{
							$tmp = "Yes"
						}
						Else
						{
							$tmp = "No"
						}
						$rowdata += @(,("Enable offline access",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
						$tmp = $Null
					}
					If($Application.CachingOption)
					{
						Switch ($Application.CachingOption)
						{
							"Unknown"   {$Tmp = "Unknown"; Break}
							"PreLaunch" {$Tmp = "Cache application prior to launching"; Break}
							"AtLaunch"  {$Tmp = "Cache application during launch"; Break}
							Default		{$Tmp = "Could not be determined: $($Application.CachingOption)"; Break}
						}
						$rowdata += @(,("Cache preference",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
						$tmp = $Null
					}
				}
				
				#location properties
				If(!$streamedapp)
				{
					$rowdata += @(,("Command Line",($htmlsilver -bor $htmlbold),$Application.CommandLineExecutable,$htmlwhite))
					$rowdata += @(,("Working directory",($htmlsilver -bor $htmlbold),$Application.WorkingDirectory,$htmlwhite))
					
					#servers properties
					If($AppServerInfoResults)
					{
						If(![String]::IsNullOrEmpty($AppServerInfo.ServerNames))
						{
							$TempArray = @($AppServerInfo.ServerNames | Sort-Object ServerNames)
							$rowdata += @(,("Servers",($htmlsilver -bor $htmlbold),$TempArray[0],$htmlwhite))
							$cnt = -1
							ForEach($Item in $TempArray)
							{
								$cnt++
								If($cnt -gt 0)
								{
									$rowdata += @(,("",($htmlsilver -bor $htmlbold),$Item,$htmlwhite))
								}
							}
							$TempArray = $Null
						}
						If(![String]::IsNullOrEmpty($AppServerInfo.WorkerGroupNames))
						{
							$TempArray = @($AppServerInfo.WorkerGroupNames | Sort-Object WorkerGroupNames)
							$rowdata += @(,("Worker Groups",($htmlsilver -bor $htmlbold),$TempArray[0],$htmlwhite))
							$cnt = -1
							ForEach($Item in $TempArray)
							{
								$cnt++
								If($cnt -gt 0)
								{
									$rowdata += @(,("",($htmlsilver -bor $htmlbold),$Item,$htmlwhite))
								}
							}
							$TempArray = $Null
						}
					}
					Else
					{
						$rowdata += @(,("Unable to retrieve a list of Servers or Worker Groups for this application",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					}
				}
			
				#users properties
				If($Application.AnonymousConnectionsAllowed)
				{
					$rowdata += @(,("Allow anonymous users",($htmlsilver -bor $htmlbold),$Application.AnonymousConnectionsAllowed,$htmlwhite))
				}
				Else
				{
					If($AppServerInfoResults)
					{
						$TempArray = @($AppServerInfo.Accounts | Sort-Object Accounts)
						$rowdata += @(,("Users",($htmlsilver -bor $htmlbold),$TempArray[0],$htmlwhite))
						$cnt = -1
						ForEach($Item in $TempArray)
						{
							$cnt++
							If($cnt -gt 0)
							{
								$rowdata += @(,("",($htmlsilver -bor $htmlbold),$Item,$htmlwhite))
							}
						}
						$TempArray = $Null
					}
					Else
					{
						$rowdata += @(,("Unable to retrieve a list of Users for this application",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					}
				}	

				#shortcut presentation properties
				#application icon is ignored
				$rowdata += @(,("Client application folder",($htmlsilver -bor $htmlbold),$Application.ClientFolder,$htmlwhite))
				If($Application.AddToClientStartMenu)
				{
					$rowdata += @(,("Add to client's start menu",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					If($Application.StartMenuFolder)
					{
						$rowdata += @(,("Start menu folder",($htmlsilver -bor $htmlbold),$Application.StartMenuFolder,$htmlwhite))
					}
				}
				If($Application.AddToClientDesktop)
				{
					$rowdata += @(,("Add shortcut to the client's desktop",($htmlsilver -bor $htmlbold),"",$htmlwhite))
				}
			
				#access control properties
				If($Application.ConnectionsThroughAccessGatewayAllowed)
				{
					If($Application.ConnectionsThroughAccessGatewayAllowed)
					{
						$tmp = "Yes"
					} 
					Else
					{
						$tmp = "No"
					}
					$rowdata += @(,("Allow connections made through AGAE",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					$tmp = $Null
				}
				If($Application.OtherConnectionsAllowed)
				{
					If($Application.OtherConnectionsAllowed)
					{
						$tmp = "Yes"
					} 
					Else
					{
						$tmp = "No"
					}
					$rowdata += @(,("Any connection",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					$tmp = $Null
				}
				If($Application.AccessSessionConditionsEnabled)
				{
					$rowdata += @(,("Any connection that meets any of the following filters",($htmlsilver -bor $htmlbold),$Application.AccessSessionConditionsEnabled,$htmlwhite))
					$rowdata += @(,("     Access Gateway Filters",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					ForEach($AccessCondition in $Application.AccessSessionConditions)
					{
						[string]$Tmp = $AccessCondition
						[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
						[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
						$tmp = "Farm name: $AGFarm  Filter: $AGFilter"
						$rowdata += @(,("",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					}

					$tmp = $Null
					$AGFarm = $Null
					$AGFilter = $Null
				}
			
				#content redirection properties
				If($AppServerInfoResults)
				{
					If($AppServerInfo.FileTypes)
					{
						$TempArray = @($AppServerInfo.FileTypes | Sort-Object)
						$rowdata += @(,("File type associations",($htmlsilver -bor $htmlbold),$TempArray[0],$htmlwhite))
						$cnt = -1
						ForEach($Item in $TempArray)
						{
							$cnt++
							If($cnt -gt 0)
							{
								$rowdata += @(,("",($htmlsilver -bor $htmlbold),$Item,$htmlwhite))
							}
						}
						$TempArray = $Null
					}
					Else
					{
						$rowdata += @(,("File Type Associations for this application",($htmlsilver -bor $htmlbold),"None",$htmlwhite))
					}
				}
				Else
				{
					$rowdata += @(,("Unable to retrieve the list of FTAs for this application",($htmlsilver -bor $htmlbold),"",$htmlwhite))
				}
			
				#if streamed app, Alternate profiles
				If($streamedapp)
				{
					If($Application.AlternateProfiles)
					{
						$rowdata += @(,("Primary application profile location",($htmlsilver -bor $htmlbold),$Application.AlternateProfiles,$htmlwhite))
					}
				
					#if streamed app, User privileges properties
					If($Application.RunAsLeastPrivilegedUser)
					{
						$rowdata += @(,("Run app as a least-privileged user account",($htmlsilver -bor $htmlbold),$Application.RunAsLeastPrivilegedUser,$htmlwhite))
					}
				}
			
				#limits properties
				If($Application.InstanceLimit -eq -1)
				{
					$tmp = "No limit set"
				}
				Else
				{
					$tmp = $Application.InstanceLimit.ToString()
				}
				$rowdata += @(,("Limit instances allowed to run in server farm",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$tmp = $Null

			
				If($Application.MultipleInstancesPerUserAllowed) 
				{
					$tmp = "No"
				} 
				Else
				{
					$tmp = "Yes"
				}
				$rowdata += @(,("Allow only 1 instance of app for each user",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$tmp = $Null
			
			
				If($Application.CpuPriorityLevel)
				{
					Switch ($Application.CpuPriorityLevel)
					{
						"Unknown"     	{$Tmp = "Unknown"; Break}
						"BelowNormal" 	{$Tmp = "Below Normal"; Break}
						"Low"         	{$Tmp = "Low"; Break}
						"Normal"      	{$Tmp = "Normal"; Break}
						"AboveNormal" 	{$Tmp = "Above Normal"; Break}
						"High"        	{$Tmp = "High"; Break}
						Default			{$Tmp = "Application importance could not be determined: $($Application.CpuPriorityLevel)"; Break}
					}
					$rowdata += @(,("Application importance",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				}
				
				#client options properties
				Switch ($Application.AudioType)
				{
					"Unknown" 	{$Tmp = "Unknown"; Break}
					"None"    	{$Tmp = "Not Enabled"; Break}
					"Basic"   	{$Tmp = "Enabled"; Break}
					Default		{$Tmp = "Enable legacy audio could not be determined: $($Application.AudioType)"; Break}
				}
				$rowdata += @(,("Enable legacy audio",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$tmp = $Null

				If($Application.AudioRequired)
				{
					$tmp = "Enabled"
				}
				Else
				{
					$tmp = "Disabled"
				}
				$rowdata += @(,("Minimum requirement",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$tmp = $Null

				If($Application.SslConnectionEnabled)
				{
					If($Application.SslConnectionEnabled)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
				}
				$rowdata += @(,("Enable SSL and TLS protocols",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$tmp = $Null

				If($Application.EncryptionLevel)
				{
					Switch ($Application.EncryptionLevel)
					{
						"Unknown" 	{$Tmp = "Unknown"; Break}
						"Basic"   	{$Tmp = "Basic"; Break}
						"LogOn"   	{$Tmp = "128-Bit Login Only (RC-5)"; Break}
						"Bits40"  	{$Tmp = "40-Bit (RC-5)"; Break}
						"Bits56"  	{$Tmp = "56-Bit (RC-5)"; Break}
						"Bits128" 	{$Tmp = "128-Bit (RC-5)"; Break}
						Default		{$Tmp = "Encryption could not be determined: $($Application.EncryptionLevel)"; Break}
					}
				}
				$rowdata += @(,("Encryption",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$tmp = $Null
				
				If($Application.EncryptionRequired)
				{
					If($Application.EncryptionRequired)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
				}
				$rowdata += @(,("Minimum requirement",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$tmp = $Null
			
				#another weird one, if True then this is Disabled
				If($Application.WaitOnPrinterCreation) 
				{
					$Tmp = "Disabled"
				} 
				Else
				{
					$Tmp = "Enabled"
				}
				$rowdata += @(,("Start app w/o waiting for printer creation",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$tmp = $Null
				
				#appearance properties
				If($Application.WindowType)
				{
					$rowdata += @(,("Session window size",($htmlsilver -bor $htmlbold),$Application.WindowType,$htmlwhite))
				}
				If($Application.ColorDepth)
				{
					Switch ($Application.ColorDepth)
					{
						"Unknown"     	{$Tmp = "Unknown color depth"; Break}
						"Colors8Bit"  	{$Tmp = "256-color (8-bit)"; Break}
						"Colors16Bit" 	{$Tmp = "Better Speed (16-bit)"; Break}
						"Colors32Bit" 	{$Tmp = "Better Appearance (32-bit)"; Break}
						Default			{$Tmp = "Maximum color quality could not be determined: $($Application.ColorDepth)"; Break}
					}
					$rowdata += @(,("Maximum color quality",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					$tmp = $Null
				}
				If($Application.TitleBarHidden)
				{
					If($Application.TitleBarHidden)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
					$rowdata += @(,("Hide application title bar",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					$tmp = $Null
				}
				If($Application.MaximizedOnStartup)
				{
					If($Application.MaximizedOnStartup)
					{
						$Tmp = "Enabled"
					}
					Else
					{
						$Tmp = "Disabled"
					}
					$rowdata += @(,("Maximize application at startup",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					$tmp = $Null
				}
				$AppServerInfo = $Null
				
				$msg = ""
				$columnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 $Application.DisplayName
			}
			ElseIf($Text)
			{
				Line 0 $Application.DisplayName
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $Application.DisplayName
			}
		}
	}
}
#endregion

#region configuration logging functions
Function ProcessConfigLogging
{
	If(!$Summary -and ($Section -eq "All" -or $Section -eq "ConfigLog"))
	{
		Write-Host "$(Get-Date): Setting summary variables" -BackgroundColor Black -ForegroundColor Yellow
		[int]$Script:TotalConfigLogItems = 0

		If($ConfigLog)
		{
			Write-Host "$(Get-Date): Processing Configuration Logging/History Report" -BackgroundColor Black -ForegroundColor Yellow
			#history AKA Configuration Logging report
			#only process if $ConfigLog = $True and XA65ConfigLog.udl file exists
			#build connection string
			#User ID is account that has access permission for the configuration logging database
			#Initial Catalog is the name of the Configuration Logging SQL Database
			#bug fixed by Esther Barthel
			If(Test-Path "$($pwd.path)\XA65ConfigLog.udl")
			{
				Write-Host "$(Get-Date): `tRetrieving logging data for date range $($StartDate) through $($EndDate)" -BackgroundColor Black -ForegroundColor Yellow
				$ConnectionString = Get-Content "$($pwd.path)\XA65ConfigLog.udl"| Select-Object -last 1
				
				If("" -eq $ConnectionString -or $Null -eq $ConnectionString)
				{
					Write-Warning "Configuration Logging connection string for the XA65ConfigLog.udl file was not found"
					Write-Warning "$(Get-Date): Unable to process Configuration Logging/History Report"
				}
				Else
				{
					$ConfigLogReport = @(Get-CtxConfigurationLogReport -connectionstring $ConnectionString -TimePeriodFrom $StartDate -TimePeriodTo $EndDate -EA 0)
					$Script:TotalConfigLogItems = $ConfigLogReport.Count
					
					If($? -and "" -ne $ConfigLogReport)
					{
						OutputConfigLogging $ConfigLogReport
					}
					ElseIf($? -and "" -eq $ConfigLogReport)
					{
						$txt = "There was no configuration logging data returned"
						OutputWarning $txt
					}
					Else
					{
						$txt = "Unable to retrieve configuration logging data"
						OutputWarning $txt
					}
					Write-Host "$(Get-Date): Finished Processing Configuration Logging/History Report" -BackgroundColor Black -ForegroundColor Yellow
				}
			}
			Else 
			{
				Write-Warning "Configuration Logging is enabled but the XA65ConfigLog.udl file was not found"
			}
			$ConnectionString = $Null
			$ConfigLogReport = $Null
		}
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputConfigLogging
{
	Param([object] $ConfigLogReport)
	Write-Host "$(Get-Date): `tProcessing $($ConfigLogReport.Count) history items" -BackgroundColor Black -ForegroundColor Yellow
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "History"
		WriteWordLine 0 0 "For date range $($StartDate) through $($EndDate)"
		[System.Collections.Hashtable[]] $WordTable = @();
		[int] $CurrentServiceIndex = 2;
		ForEach($Item in $ConfigLogReport)
		{
			$WordTableRowHash = @{ 
			Date = $Item.Date;
			Account = $Item.Account;
			Description = $Item.Description;
			TaskType = $Item.TaskType;
			ItemType = $Item.ItemType;
			ItemName = $Item.ItemName;
			}
			$WordTable += $WordTableRowHash;
			$CurrentServiceIndex++;
		}
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns Date, Account, Description, TaskType, ItemType, ItemName `
		-Headers "Date","Account","Change description","Type of change","Type of item","Name of item" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitContent;

		SetWordCellFormat -Collection $Table -Size 9
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "History"
		Line 0 "For date range $($StartDate) through $($EndDate)"
		Line 0 ""
		ForEach($Item in $ConfigLogReport)
		{
			Line 1 "Date`t`t`t: " $Item.Date
			Line 1 "Account`t`t`t: " $Item.Account
			Line 1 "Change description`t: " $Item.Description
			Line 1 "Type of change`t`t: " $Item.TaskType
			Line 1 "Type of item`t`t: " $Item.ItemType
			Line 1 "Name of item`t`t: " $Item.ItemName
			Line 0 ""
		}
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "History"
		WriteHTMLLine 0 0 "For date range $($StartDate) through $($EndDate)"
		$rowdata = @()
		ForEach($Item in $ConfigLogReport)
		{
			$rowdata += @(,(
			$Item.Date,$htmlwhite,
			$Item.Account,$htmlwhite,
			$Item.Description,$htmlwhite,
			$Item.TaskType,$htmlwhite,
			$Item.ItemType,$htmlwhite,
			$Item.ItemName,$htmlwhite))
		}
		$columnHeaders = @(
		'Date',($htmlsilver -bor $htmlbold),
		'Account',($htmlsilver -bor $htmlbold),
		'Change description',($htmlsilver -bor $htmlbold),
		'Type of change',($htmlsilver -bor $htmlbold),
		'Type of item',($htmlsilver -bor $htmlbold),
		'Name of item',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders -fontsize 2
	}
}
#endregion

#region load balancing policy functions
Function ProcessLoadBalancingPolicies
{
	If($Section -eq "All" -or $Section -eq "LBPolicies")
	{
		#load balancing policies
		Write-Host "$(Get-Date): Processing Load Balancing Policies" -BackgroundColor Black -ForegroundColor Yellow
		[int]$Script:TotalLBPolicies = 0

		Write-Host "$(Get-Date): `tRetrieving Load Balancing Policies" -BackgroundColor Black -ForegroundColor Yellow
		$LoadBalancingPolicies = @(Get-XALoadBalancingPolicy -EA 0 | Sort-Object PolicyName)

		If($? -and $Null -ne $LoadBalancingPolicies)
		{
			OutputLoadBalancingPolicies $LoadBalancingPolicies
		}
		Elseif($Null -eq $LoadBalancingPolicies)
		{
			Write-Host "$(Get-Date): There are no Load balancing policies created" -BackgroundColor Black -ForegroundColor Yellow
		}
		Else 
		{
			Write-Warning "Load balancing policy information could not be retrieved"
		}
		$LoadBalancingPolicies = $Null
		Write-Host "$(Get-Date): Finished Processing Load Balancing Policies" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputLoadBalancingPolicies
{
	Param([object] $LoadBalancingPolicies)

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Load Balancing Policies"
	}
	ElseIf($Text)
	{
		Line 0 "Load Balancing Policies"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Load Balancing Policies"
	}

	ForEach($LoadBalancingPolicy in $LoadBalancingPolicies)
	{
		$Script:TotalLBPolicies++
		Write-Host "$(Get-Date): `t`tProcessing Load Balancing Policy $($LoadBalancingPolicy.PolicyName)" -BackgroundColor Black -ForegroundColor Yellow
		$LoadBalancingPolicyConfiguration = Get-XALoadBalancingPolicyConfiguration -PolicyName $LoadBalancingPolicy.PolicyName -EA 0
		$LoadBalancingPolicyFilter = Get-XALoadBalancingPolicyFilter -PolicyName $LoadBalancingPolicy.PolicyName -EA 0
	
		If(!$Summary)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 2 0 $LoadBalancingPolicy.PolicyName
				If(![String]::IsNullOrEmpty($LoadBalancingPolicy.Description))
				{
					WriteWordLine 0 1 "Description`t: " $LoadBalancingPolicy.Description
				}
				WriteWordLine 0 1 "Enabled`t`t: " -nonewline
				If($LoadBalancingPolicy.Enabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				WriteWordLine 0 1 "Priority`t`t: " $LoadBalancingPolicy.Priority
			
				WriteWordLine 0 1 "Filter based on Access Control: " -nonewline
				If($LoadBalancingPolicyFilter.AccessControlEnabled)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
				If($LoadBalancingPolicyFilter.AccessControlEnabled)
				{
					WriteWordLine 0 1 "Apply to connections made through Access Gateway: " -nonewline
					If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
					{
						WriteWordLine 0 1 "Any connection that meets any of the following filters"
						If($LoadBalancingPolicyFilter.AccessSessionConditions)
						{
							ForEach($AccessSessionCondition in $LoadBalancingPolicyFilter.AccessSessionConditions)
							{
								[string]$Tmp = $AccessSessionCondition
								[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
								[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
								WriteWordLine 0 2 "FarmName: " $AGFarm
								WriteWordLine 0 2 "Filter: " $AGFilter
								WriteWordLine 0 0 ""
							}

							$tmp = $Null
							$AGFarm = $Null
							$AGFilter = $Null
						}
					}
					If($LoadBalancingPolicyFilter.AllowOtherConnections)
					{
						WriteWordLine 0 2 "Apply to all other connections"
					} 
				}
			
				If($LoadBalancingPolicyFilter.ClientIPAddressEnabled)
				{
					WriteWordLine 0 1 "Filter based on client IP address"
					If($LoadBalancingPolicyFilter.ApplyToAllClientIPAddresses)
					{
						WriteWordLine 0 2 "Apply to all client IP addresses"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedIPAddresses)
						{
							ForEach($AllowedIPAddress in $LoadBalancingPolicyFilter.AllowedIPAddresses)
							{
								WriteWordLine 0 2 "Client IP Address Matched: " $AllowedIPAddress
							}
						}
						If($LoadBalancingPolicyFilter.DeniedIPAddresses)
						{
							ForEach($DeniedIPAddress in $LoadBalancingPolicyFilter.DeniedIPAddresses)
							{
								WriteWordLine 0 2 "Client IP Address Ignored: " $DeniedIPAddress
							}
						}
					}
				}
				If($LoadBalancingPolicyFilter.ClientNameEnabled)
				{
					WriteWordLine 0 1 "Filter based on client name"
					If($LoadBalancingPolicyFilter.ApplyToAllClientNames)
					{
						WriteWordLine 0 2 "Apply to all client names"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedClientNames)
						{
							ForEach($AllowedClientName in $LoadBalancingPolicyFilter.AllowedClientNames)
							{
								WriteWordLine 0 2 "Client Name Matched: " $AllowedClientName
							}
						}
						If($LoadBalancingPolicyFilter.DeniedClientNames)
						{
							ForEach($DeniedClientName in $LoadBalancingPolicyFilter.DeniedClientNames)
							{
								WriteWordLine 0 2 "Client Name Ignored: " $DeniedClientName
							}
						}
					}
				}
				If($LoadBalancingPolicyFilter.AccountEnabled)
				{
					WriteWordLine 0 1 "Filter based on user"
					WriteWordLine 0 2 "Apply to anonymous users: " -nonewline
					If($LoadBalancingPolicyFilter.ApplyToAnonymousAccounts)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
					If($LoadBalancingPolicyFilter.ApplyToAllExplicitAccounts)
					{
						WriteWordLine 0 2 "Apply to all explicit (non-anonymous) users"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedAccounts)
						{
							ForEach($AllowedAccount in $LoadBalancingPolicyFilter.AllowedAccounts)
							{
								WriteWordLine 0 2 "User Matched: " $AllowedAccount
							}
						}
						If($LoadBalancingPolicyFilter.DeniedAccounts)
						{
							ForEach($DeniedAccount in $LoadBalancingPolicyFilter.DeniedAccounts)
							{
								WriteWordLine 0 2 "User Ignored: " $DeniedAccount
							}
						}
					}
				}
				If($LoadBalancingPolicyConfiguration.WorkerGroupPreferenceAndFailoverState)
				{
					WriteWordLine 0 1 "Configure application connection preference based on worker group"
					If($LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
					{
						ForEach($WorkerGroupPreference in $LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
						{
							[string]$Tmp = $WorkerGroupPreference
							[string]$WGName = $Tmp.substring($Tmp.indexof("=")+1)
							[string]$WGPriority = $Tmp.substring($Tmp.indexof(":")+1, (($Tmp.indexof("=")-1)-$Tmp.indexof(":")))
							WriteWordLine 0 2 "Worker Group: " $WGName
							WriteWordLine 0 2 "Priority: " $WGPriority
							WriteWordLine 0 0 ""
						}

						$tmp = $Null
						$WGName = $Null
						$WGPriority = $Null
					}
				}
				If($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Enabled")
				{
					WriteWordLine 0 1 "Set the delivery protocols for applications streamed to client"
					WriteWordLine 0 2 "" -nonewline
					Switch ($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)
					{
						"Unknown"                {WriteWordLine 0 0 "Unknown"; Break}
						"ForceServerAccess"      {WriteWordLine 0 0 "Do not allow applications to stream to the client"; Break}
						"ForcedStreamedDelivery" {WriteWordLine 0 0 "Force applications to stream to the client"; Break}
						Default {WriteWordLine 0 0 "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"; Break}
					}
				}
				Elseif($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Disabled")
				{
					#In the GUI, if "Set the delivery protocols for applications streamed to client" IS selected AND 
					#"Allow applications to stream to the client or run on a Terminal Server (Default)" IS selected
					#then "Set the delivery protocols for applications streamed to client" is set to Disabled
					WriteWordLine 0 1 "Set the delivery protocols for applications streamed to client"
					WriteWordLine 0 2 "Allow applications to stream to the client or run on a Terminal Server (Default)"
				}
				Else
				{
					WriteWordLine 0 1 "Streamed App Delivery is not configured"
				}
			
				$LoadBalancingPolicyConfiguration = $Null
				$LoadBalancingPolicyFilter = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 0 $LoadBalancingPolicy.PolicyName
				Line 1 "Description`t: " $LoadBalancingPolicy.Description
				Line 1 "Enabled`t`t: " -nonewline
				If($LoadBalancingPolicy.Enabled)
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
				Line 1 "Priority`t: " $LoadBalancingPolicy.Priority
				Line 0 ""
			
				Line 1 "Filter based on Access Control: " -nonewline
				If($LoadBalancingPolicyFilter.AccessControlEnabled)
				{
					Line 0 "Yes"
				}
				Else
				{
					Line 0 "No"
				}
				If($LoadBalancingPolicyFilter.AccessControlEnabled)
				{
					Line 1 "Apply to connections made through Access Gateway: " -nonewline
					If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
					{
						Line 0 "Yes"
					}
					Else
					{
						Line 0 "No"
					}
					If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
					{
						Line 1 "Any connection that meets any of the following filters"
						If($LoadBalancingPolicyFilter.AccessSessionConditions)
						{
							ForEach($AccessSessionCondition in $LoadBalancingPolicyFilter.AccessSessionConditions)
							{
								[string]$Tmp = $AccessSessionCondition
								[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
								[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
								Line 2 "Farm Name: " $AGFarm
								Line 2 "Filter: " $AGFilter
							}

							$tmp = $Null
							$AGFarm = $Null
							$AGFilter = $Null
						}
					}
					If($LoadBalancingPolicyFilter.AllowOtherConnections)
					{
						Line 2 "Apply to all other connections"
					} 
				}
			
				If($LoadBalancingPolicyFilter.ClientIPAddressEnabled)
				{
					Line 1 "Filter based on client IP address"
					If($LoadBalancingPolicyFilter.ApplyToAllClientIPAddresses)
					{
						Line 2 "Apply to all client IP addresses"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedIPAddresses)
						{
							ForEach($AllowedIPAddress in $LoadBalancingPolicyFilter.AllowedIPAddresses)
							{
								Line 2 "Client IP Address Matched: " $AllowedIPAddress
							}
						}
						If($LoadBalancingPolicyFilter.DeniedIPAddresses)
						{
							ForEach($DeniedIPAddress in $LoadBalancingPolicyFilter.DeniedIPAddresses)
							{
								Line 2 "Client IP Address Ignored: " $DeniedIPAddress
							}
						}
					}
				}
				If($LoadBalancingPolicyFilter.ClientNameEnabled)
				{
					Line 1 "Filter based on client name"
					If($LoadBalancingPolicyFilter.ApplyToAllClientNames)
					{
						Line 2 "Apply to all client names"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedClientNames)
						{
							ForEach($AllowedClientName in $LoadBalancingPolicyFilter.AllowedClientNames)
							{
								Line 2 "Client Name Matched: " $AllowedClientName
							}
						}
						If($LoadBalancingPolicyFilter.DeniedClientNames)
						{
							ForEach($DeniedClientName in $LoadBalancingPolicyFilter.DeniedClientNames)
							{
								Line 2 "Client Name Ignored: " $DeniedClientName
							}
						}
					}
				}
				If($LoadBalancingPolicyFilter.AccountEnabled)
				{
					Line 1 "Filter based on user"
					Line 2 "Apply to anonymous users: " -nonewline
					If($LoadBalancingPolicyFilter.ApplyToAnonymousAccounts)
					{
						Line 0 "Yes"
					}
					Else
					{
						Line 0 "No"
					}
					If($LoadBalancingPolicyFilter.ApplyToAllExplicitAccounts)
					{
						Line 2 "Apply to all explicit (non-anonymous) users"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedAccounts)
						{
							ForEach($AllowedAccount in $LoadBalancingPolicyFilter.AllowedAccounts)
							{
								Line 2 "User Matched: " $AllowedAccount
							}
						}
						If($LoadBalancingPolicyFilter.DeniedAccounts)
						{
							ForEach($DeniedAccount in $LoadBalancingPolicyFilter.DeniedAccounts)
							{
								Line 2 "User Ignored: " $DeniedAccount
							}
						}
					}
					Line 0 ""
				}
				If($LoadBalancingPolicyConfiguration.WorkerGroupPreferenceAndFailoverState)
				{
					Line 1 "Configure application connection preference based on worker group"
					If($LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
					{
						ForEach($WorkerGroupPreference in $LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
						{
							[string]$Tmp = $WorkerGroupPreference
							[string]$WGName = $Tmp.substring($Tmp.indexof("=")+1)
							[string]$WGPriority = $Tmp.substring($Tmp.indexof(":")+1, (($Tmp.indexof("=")-1)-$Tmp.indexof(":")))
							Line 2 "Worker Group`t: " $WGName
							Line 2 "Priority`t: " $WGPriority
							Line 0 ""
						}

						$tmp = $Null
						$WGName = $Null
						$WGPriority = $Null
					}
				}
				If($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Enabled")
				{
					Line 1 "Set the delivery protocols for applications streamed to client"
					Line 2 "" -nonewline
					Switch ($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)
					{
						"Unknown"					{Line 0 "Unknown"; Break}
						"ForceServerAccess"			{Line 0 "Do not allow applications to stream to the client"; Break}
						"ForcedStreamedDelivery"	{Line 0 "Force applications to stream to the client"; Break}
						Default						{Line 0 "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"; Break}
					}
				}
				Elseif($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Disabled")
				{
					#In the GUI, if "Set the delivery protocols for applications streamed to client" IS selected AND 
					#"Allow applications to stream to the client or run on a Terminal Server (Default)" IS selected
					#then "Set the delivery protocols for applications streamed to client" is set to Disabled
					Line 1 "Set the delivery protocols for applications streamed to client"
					Line 2 "Allow applications to stream to the client or run on a Terminal Server (Default)"
				}
				Else
				{
					Line 1 "Streamed App Delivery is not configured"
				}
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 $LoadBalancingPolicy.PolicyName
				If(![String]::IsNullOrEmpty($LoadBalancingPolicy.Description))
				{
					WriteHTMLLine 0 1 "Description: " $LoadBalancingPolicy.Description
				}
				If($LoadBalancingPolicy.Enabled)
				{
					$tmp = "Yes"
				}
				Else
				{
					$tmp = "No"
				}
				WriteHTMLLine 0 1 "Enabled: " $tmp
				WriteHTMLLine 0 1 "Priority: " $LoadBalancingPolicy.Priority
			
				If($LoadBalancingPolicyFilter.AccessControlEnabled)
				{
					$tmp = "Yes"
				}
				Else
				{
					$tmp = "No"
				}
				WriteHTMLLine 0 1 "Filter based on Access Control: " $tmp

				If($LoadBalancingPolicyFilter.AccessControlEnabled)
				{
					If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
					{
						$tmp = "Yes"
					}
					Else
					{
						$tmp = "No"
					}
					WriteHTMLLine 0 1 "Apply to connections made through Access Gateway: " $tmp
					If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
					{
						WriteHTMLLine 0 1 "Any connection that meets any of the following filters"
						If($LoadBalancingPolicyFilter.AccessSessionConditions)
						{
							ForEach($AccessSessionCondition in $LoadBalancingPolicyFilter.AccessSessionConditions)
							{
								[string]$Tmp = $AccessSessionCondition
								[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
								[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
								WriteHTMLLine 0 2 "FarmName: " $AGFarm
								WriteHTMLLine 0 2 "Filter: " $AGFilter
								WriteHTMLLine 0 0 ""
							}

							$tmp = $Null
							$AGFarm = $Null
							$AGFilter = $Null
						}
					}
					If($LoadBalancingPolicyFilter.AllowOtherConnections)
					{
						WriteHTMLLine 0 2 "Apply to all other connections"
					} 
				}
			
				If($LoadBalancingPolicyFilter.ClientIPAddressEnabled)
				{
					WriteHTMLLine 0 1 "Filter based on client IP address"
					If($LoadBalancingPolicyFilter.ApplyToAllClientIPAddresses)
					{
						WriteHTMLLine 0 2 "Apply to all client IP addresses"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedIPAddresses)
						{
							ForEach($AllowedIPAddress in $LoadBalancingPolicyFilter.AllowedIPAddresses)
							{
								WriteHTMLLine 0 2 "Client IP Address Matched: " $AllowedIPAddress
							}
						}
						If($LoadBalancingPolicyFilter.DeniedIPAddresses)
						{
							ForEach($DeniedIPAddress in $LoadBalancingPolicyFilter.DeniedIPAddresses)
							{
								WriteHTMLLine 0 2 "Client IP Address Ignored: " $DeniedIPAddress
							}
						}
					}
				}
				If($LoadBalancingPolicyFilter.ClientNameEnabled)
				{
					WriteHTMLLine 0 1 "Filter based on client name"
					If($LoadBalancingPolicyFilter.ApplyToAllClientNames)
					{
						WriteHTMLLine 0 2 "Apply to all client names"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedClientNames)
						{
							ForEach($AllowedClientName in $LoadBalancingPolicyFilter.AllowedClientNames)
							{
								WriteHTMLLine 0 2 "Client Name Matched: " $AllowedClientName
							}
						}
						If($LoadBalancingPolicyFilter.DeniedClientNames)
						{
							ForEach($DeniedClientName in $LoadBalancingPolicyFilter.DeniedClientNames)
							{
								WriteHTMLLine 0 2 "Client Name Ignored: " $DeniedClientName
							}
						}
					}
				}
				If($LoadBalancingPolicyFilter.AccountEnabled)
				{
					WriteHTMLLine 0 1 "Filter based on user"
					If($LoadBalancingPolicyFilter.ApplyToAnonymousAccounts)
					{
						$tmp = "Yes"
					}
					Else
					{
						$tmp = "No"
					}
					WriteHTMLLine 0 2 "Apply to anonymous users: " $tmp
					If($LoadBalancingPolicyFilter.ApplyToAllExplicitAccounts)
					{
						WriteHTMLLine 0 2 "Apply to all explicit (non-anonymous) users"
					} 
					Else
					{
						If($LoadBalancingPolicyFilter.AllowedAccounts)
						{
							ForEach($AllowedAccount in $LoadBalancingPolicyFilter.AllowedAccounts)
							{
								WriteHTMLLine 0 2 "User Matched: " $AllowedAccount
							}
						}
						If($LoadBalancingPolicyFilter.DeniedAccounts)
						{
							ForEach($DeniedAccount in $LoadBalancingPolicyFilter.DeniedAccounts)
							{
								WriteHTMLLine 0 2 "User Ignored: " $DeniedAccount
							}
						}
					}
				}
				If($LoadBalancingPolicyConfiguration.WorkerGroupPreferenceAndFailoverState)
				{
					WriteHTMLLine 0 1 "Configure application connection preference based on worker group"
					If($LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
					{
						ForEach($WorkerGroupPreference in $LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
						{
							[string]$Tmp = $WorkerGroupPreference
							[string]$WGName = $Tmp.substring($Tmp.indexof("=")+1)
							[string]$WGPriority = $Tmp.substring($Tmp.indexof(":")+1, (($Tmp.indexof("=")-1)-$Tmp.indexof(":")))
							WriteHTMLLine 0 2 "Worker Group: " $WGName
							WriteHTMLLine 0 2 "Priority: " $WGPriority
							WriteHTMLLine 0 0 ""
						}

						$tmp = $Null
						$WGName = $Null
						$WGPriority = $Null
					}
				}
				If($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Enabled")
				{
					WriteHTMLLine 0 1 "Set the delivery protocols for applications streamed to client"
					Switch ($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)
					{
						"Unknown"					{$tmp = "Unknown"; Break}
						"ForceServerAccess"			{$tmp = "Do not allow applications to stream to the client"; Break}
						"ForcedStreamedDelivery"	{$tmp = "Force applications to stream to the client"; Break}
						Default						{$tmp = "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"; Break}
					}
					WriteHTMLLine 0 2 "" $tmp
				}
				Elseif($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Disabled")
				{
					#In the GUI, if "Set the delivery protocols for applications streamed to client" IS selected AND 
					#"Allow applications to stream to the client or run on a Terminal Server (Default)" IS selected
					#then "Set the delivery protocols for applications streamed to client" is set to Disabled
					WriteHTMLLine 0 1 "Set the delivery protocols for applications streamed to client"
					WriteHTMLLine 0 2 "Allow applications to stream to the client or run on a Terminal Server (Default)"
				}
				Else
				{
					WriteHTMLLine 0 1 "Streamed App Delivery is not configured"
				}
			
				$LoadBalancingPolicyConfiguration = $Null
				$LoadBalancingPolicyFilter = $Null
				WriteHTMLLine 0 0 ""
			}
			$LoadBalancingPolicyConfiguration = $Null
			$LoadBalancingPolicyFilter = $Null
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 $LoadBalancingPolicy.PolicyName
			}
			ElseIf($Text)
			{
				Line 0 $LoadBalancingPolicy.PolicyName
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $LoadBalancingPolicy.PolicyName
			}
		}
	}
	
}
#endregion

#region load evaluator functions
Function ProcessLoadEvaluators
{
	If($Section -eq "All" -or $Section -eq "LoadEvals")
	{
		#load evaluators
		Write-Host "$(Get-Date): Processing Load Evaluators" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): `tSetting summary variables" -BackgroundColor Black -ForegroundColor Yellow
		[int]$Script:TotalLoadEvaluators = 0

		Write-Host "$(Get-Date): `tRetrieving Load Evaluators" -BackgroundColor Black -ForegroundColor Yellow
		$LoadEvaluators = Get-XALoadEvaluator -EA 0| Sort-Object LoadEvaluatorName

		If($? -and $Null -ne $LoadEvaluators)
		{
			OutputLoadEvaluators $LoadEvaluators
		}
		ElseIf($? -and $Null -eq $LoadEvaluators)
		{
			Write-Warning "No results returned for Load Evaluator information"
		}
		Else
		{
			Write-Warning "Load Evaluator information could not be retrieved"
		}
		Write-Host "$(Get-Date): Finished Processing Load Evaluators" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputLoadEvaluators
{
	Param([object] $LoadEvaluators)

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Load Evaluators"
	}
	ElseIf($Text)
	{
		Line 0 "Load Evaluators"
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Load Evaluators"
	}
	
	ForEach($LoadEvaluator in $LoadEvaluators)
	{
		$Script:TotalLoadEvaluators++
		Write-Host "$(Get-Date): `t`tProcessing Load Evaluator $($LoadEvaluator.LoadEvaluatorName)" -BackgroundColor Black -ForegroundColor Yellow
		If(!$Summary)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 2 0 $LoadEvaluator.LoadEvaluatorName
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				$ScriptInformation += @{ Data = "Description"; Value = $LoadEvaluator.Description; }
				
				If($LoadEvaluator.IsBuiltIn)
				{
					$ScriptInformation += @{ Data = "Built-in Load Evaluator"; Value = ""; }
				} 
				Else 
				{
					$ScriptInformation += @{ Data = "User created load evaluator"; Value = ""; }
				}
			
				If($LoadEvaluator.ApplicationUserLoadEnabled)
				{
					$ScriptInformation += @{ Data = "Application User Load Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Report full load when the # of users for this application equals"; Value = $LoadEvaluator.ApplicationUserLoad; }
					$ScriptInformation += @{ Data = "     Application"; Value = $LoadEvaluator.ApplicationBrowserName; }
				}
			
				If($LoadEvaluator.ContextSwitchesEnabled)
				{
					$ScriptInformation += @{ Data = "Context Switches Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Report full load when the # of context Switches per second is > than"; Value = $LoadEvaluator.ContextSwitches[1]; }
					$ScriptInformation += @{ Data = "     Report no load when the # of context Switches per second is <= to"; Value = $LoadEvaluator.ContextSwitches[0]; }
				}
			
				If($LoadEvaluator.CpuUtilizationEnabled)
				{
					$ScriptInformation += @{ Data = "CPU Utilization Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Report full load when the processor utilization % is > than"; Value = $LoadEvaluator.CpuUtilization[1]; }
					$ScriptInformation += @{ Data = "     Report no load when the processor utilization % is <= to"; Value = $LoadEvaluator.CpuUtilization[0]; }
				}
			
				If($LoadEvaluator.DiskDataIOEnabled)
				{
					$ScriptInformation += @{ Data = "Disk Data I/O Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Report full load when the total disk I/O in kbps is > than"; Value = $LoadEvaluator.DiskDataIO[1]; }
					$ScriptInformation += @{ Data = "     Report no load when the total disk I/O in kbps per second is <= to"; Value = $LoadEvaluator.DiskDataIO[0]; }
				}
			
				If($LoadEvaluator.DiskOperationsEnabled)
				{
					$ScriptInformation += @{ Data = "Disk Operations Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Report full load when the total # of R/W operations per second is > than"; Value = $LoadEvaluator.DiskOperations[1]; }
					$ScriptInformation += @{ Data = "     Report no load when the total # of R/W operations per second is <= to"; Value = $LoadEvaluator.DiskOperations[0]; }
				}
			
				If($LoadEvaluator.IPRangesEnabled)
				{
					$ScriptInformation += @{ Data = "IP Range Settings"; Value = ""; }
					If($LoadEvaluator.IPRangesAllowed)
					{
						$tmp - "Allow client connections from the listed IP Ranges"
					} 
					Else 
					{
						$tmp = "Deny client connections from the listed IP Ranges"
					}
					$ScriptInformation += @{ Data = $tmp; Value = ""; }
					$ScriptInformation += @{ Data = "IP Address Ranges"; Value = $LoadEvaluator.IPRanges[0]; }
					$cnt =-1
					ForEach($IPRange in $LoadEvaluator.IPRanges)
					{
						$cnt++
						If($cnt -gt 0)
						{
							$ScriptInformation += @{ Data = ""; Value = $IPRange; }
						}
					}
					$tmp = $Null
					$cnt = $Null
				}
			
				If($LoadEvaluator.LoadThrottlingEnabled)
				{
					Switch ($LoadEvaluator.LoadThrottling)
					{
						"Unknown"		{$tmp = "Unknown"; Break}
						"Extreme"		{$tmp = "Extreme"; Break}
						"High"			{$tmp = "High (Default)"; Break}
						"MediumHigh"	{$tmp = "Medium High"; Break}
						"Medium"		{$tmp = "Medium"; Break}
						"MediumLow"		{$tmp = "Medium Low"; Break}
						Default			{$tmp = "Impact of logons on load could not be determined: $($LoadEvaluator.LoadThrottling)"; Break}
					}
					$ScriptInformation += @{ Data = "Load Throttling Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Impact of logons on load"; Value = $tmp; }
					$tmp = $Null
				}
			
				If($LoadEvaluator.MemoryUsageEnabled)
				{
					$ScriptInformation += @{ Data = "Memory Usage Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Report full load when the memory usage is > than"; Value = $LoadEvaluator.MemoryUsage[1]; }
					$ScriptInformation += @{ Data = "     Report no load when the memory usage is <= to"; Value = $LoadEvaluator.MemoryUsage[0]; }
				}
			
				If($LoadEvaluator.PageFaultsEnabled)
				{
					$ScriptInformation += @{ Data = "Page Faults Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Report full load when the # of page faults per second is > than"; Value = $LoadEvaluator.PageFaults[1]; }
					$ScriptInformation += @{ Data = "     Report no load when the # of page faults per second is <= to"; Value = $LoadEvaluator.PageFaults[0]; }
				}
			
				If($LoadEvaluator.PageSwapsEnabled)
				{
					$ScriptInformation += @{ Data = "Page Swaps Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Report full load when the # of page swaps per second is > than"; Value = $LoadEvaluator.PageSwaps[1]; }
					$ScriptInformation += @{ Data = "     Report no load when the # of page swaps per second is <= to"; Value = $LoadEvaluator.PageSwaps[0]; }
				}
			
				If($LoadEvaluator.ScheduleEnabled)
				{
					$ScriptInformation += @{ Data = "Scheduling Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Sunday Schedule"; Value = $LoadEvaluator.SundaySchedule; }
					$ScriptInformation += @{ Data = "     Monday Schedule"; Value = $LoadEvaluator.MondaySchedule; }
					$ScriptInformation += @{ Data = "     Tuesday Schedule"; Value = $LoadEvaluator.TuesdaySchedule; }
					$ScriptInformation += @{ Data = "     Wednesday Schedul"; Value = $LoadEvaluator.WednesdaySchedule; }
					$ScriptInformation += @{ Data = "     Thursday Schedule"; Value = $LoadEvaluator.ThursdaySchedule; }
					$ScriptInformation += @{ Data = "     Friday Schedule"; Value = $LoadEvaluator.FridaySchedule; }
					$ScriptInformation += @{ Data = "     Saturday Schedule"; Value = $LoadEvaluator.SaturdaySchedule; }
				}
			
				If($LoadEvaluator.ServerUserLoadEnabled)
				{
					$ScriptInformation += @{ Data = "Server User Load Settings"; Value = ""; }
					$ScriptInformation += @{ Data = "     Report full load when the # of server users equals"; Value = $LoadEvaluator.ServerUserLoad; }
				}
			
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				## IB - Set the header row format
				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 325;
				$Table.Columns.Item(2).Width = 175;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($Text)
			{
				Line 0 $LoadEvaluator.LoadEvaluatorName
				Line 1 "Description: " $LoadEvaluator.Description
				
				If($LoadEvaluator.IsBuiltIn)
				{
					Line 1 "Built-in Load Evaluator"
				} 
				Else 
				{
					Line 1 "User created load evaluator"
				}
			
				If($LoadEvaluator.ApplicationUserLoadEnabled)
				{
					Line 1 "Application User Load Settings"
					Line 2 "Report full load when the # of users for this application equals: " $LoadEvaluator.ApplicationUserLoad
					Line 2 "Application: " $LoadEvaluator.ApplicationBrowserName
				}
			
				If($LoadEvaluator.ContextSwitchesEnabled)
				{
					Line 1 "Context Switches Settings"
					Line 2 "Report full load when the # of context Switches per second is > than: " $LoadEvaluator.ContextSwitches[1]
					Line 2 "Report no load when the # of context Switches per second is <= to: " $LoadEvaluator.ContextSwitches[0]
				}
			
				If($LoadEvaluator.CpuUtilizationEnabled)
				{
					Line 1 "CPU Utilization Settings"
					Line 2 "Report full load when the processor utilization % is > than: " $LoadEvaluator.CpuUtilization[1]
					Line 2 "Report no load when the processor utilization % is <= to: " $LoadEvaluator.CpuUtilization[0]
				}
			
				If($LoadEvaluator.DiskDataIOEnabled)
				{
					Line 1 "Disk Data I/O Settings"
					Line 2 "Report full load when the total disk I/O in kbps is > than: " $LoadEvaluator.DiskDataIO[1]
					Line 2 "Report no load when the total disk I/O in kbps per second is <= to: " $LoadEvaluator.DiskDataIO[0]
				}
			
				If($LoadEvaluator.DiskOperationsEnabled)
				{
					Line 1 "Disk Operations Settings"
					Line 2 "Report full load when the total # of R/W operations per second is > than: " $LoadEvaluator.DiskOperations[1]
					Line 2 "Report no load when the total # of R/W operations per second is <= to: " $LoadEvaluator.DiskOperations[0]
				}
			
				If($LoadEvaluator.IPRangesEnabled)
				{
					Line 1 "IP Range Settings"
					If($LoadEvaluator.IPRangesAllowed)
					{
						Line 2 "Allow " -NoNewLine
					} 
					Else 
					{
						Line 2 "Deny " -NoNewLine
					}
					Line 0 "client connections from the listed IP Ranges"
					ForEach($IPRange in $LoadEvaluator.IPRanges)
					{
						Line 3 "IP Address Ranges: " $IPRange
					}
				}
			
				If($LoadEvaluator.LoadThrottlingEnabled)
				{
					Line 1 "Load Throttling Settings"
					Line 2 "Impact of logons on load: " -nonewline
					Switch ($LoadEvaluator.LoadThrottling)
					{
						"Unknown"		{Line 0 "Unknown"; Break}
						"Extreme"		{Line 0 "Extreme"; Break}
						"High"			{Line 0 "High (Default)"; Break}
						"MediumHigh"	{Line 0 "Medium High"; Break}
						"Medium"		{Line 0 "Medium"; Break}
						"MediumLow"		{Line 0 "Medium Low"; Break}
						Default			{Line 0 "Impact of logons on load could not be determined: $($LoadEvaluator.LoadThrottling)"; Break}
					}
				}
			
				If($LoadEvaluator.MemoryUsageEnabled)
				{
					Line 1 "Memory Usage Settings"
					Line 2 "Report full load when the memory usage is > than: " $LoadEvaluator.MemoryUsage[1]
					Line 2 "Report no load when the memory usage is <= to: " $LoadEvaluator.MemoryUsage[0]
				}
			
				If($LoadEvaluator.PageFaultsEnabled)
				{
					Line 1 "Page Faults Settings"
					Line 2 "Report full load when the # of page faults per second is > than: " $LoadEvaluator.PageFaults[1]
					Line 2 "Report no load when the # of page faults per second is <= to: " $LoadEvaluator.PageFaults[0]
				}
			
				If($LoadEvaluator.PageSwapsEnabled)
				{
					Line 1 "Page Swaps Settings"
					Line 2 "Report full load when the # of page swaps per second is > than: " $LoadEvaluator.PageSwaps[1]
					Line 2 "Report no load when the # of page swaps per second is <= to: " $LoadEvaluator.PageSwaps[0]
				}
			
				If($LoadEvaluator.ScheduleEnabled)
				{
					Line 1 "Scheduling Settings"
					Line 2 "Sunday Schedule`t: " $LoadEvaluator.SundaySchedule
					Line 2 "Monday Schedule`t: " $LoadEvaluator.MondaySchedule
					Line 2 "Tuesday Schedule`t: " $LoadEvaluator.TuesdaySchedule
					Line 2 "Wednesday Schedule`t: " $LoadEvaluator.WednesdaySchedule
					Line 2 "Thursday Schedule`t: " $LoadEvaluator.ThursdaySchedule
					Line 2 "Friday Schedule`t`t: " $LoadEvaluator.FridaySchedule
					Line 2 "Saturday Schedule`t: " $LoadEvaluator.SaturdaySchedule
				}
			
				If($LoadEvaluator.ServerUserLoadEnabled)
				{
					Line 1 "Server User Load Settings"
					Line 2 "Report full load when the # of server users equals: " $LoadEvaluator.ServerUserLoad
				}
			
				Line 0 ""
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 $LoadEvaluator.LoadEvaluatorName
				$columnHeaders = @()
				$rowdata = @()
				$columnHeaders = @("Description",($htmlsilver -bor $htmlbold),$LoadEvaluator.Description,$htmlwhite)
				
				If($LoadEvaluator.IsBuiltIn)
				{
					$rowdata += @(,("Built-in Load Evaluator",($htmlsilver -bor $htmlbold),"",$htmlwhite))
				} 
				Else 
				{
					$rowdata += @(,("User created load evaluator",($htmlsilver -bor $htmlbold),"",$htmlwhite))
				}
			
				If($LoadEvaluator.ApplicationUserLoadEnabled)
				{
					$rowdata += @(,("Application User Load Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Report full load when the # of users for this application equals",($htmlsilver -bor $htmlbold),$LoadEvaluator.ApplicationUserLoad,$htmlwhite))
					$rowdata += @(,("  Application",($htmlsilver -bor $htmlbold),$LoadEvaluator.ApplicationBrowserName,$htmlwhite))
				}
			
				If($LoadEvaluator.ContextSwitchesEnabled)
				{
					$rowdata += @(,("Context Switches Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Report full load when the # of context Switches per second is > than",($htmlsilver -bor $htmlbold),$LoadEvaluator.ContextSwitches[1],$htmlwhite))
					$rowdata += @(,("  Report no load when the # of context Switches per second is <= to",($htmlsilver -bor $htmlbold),$LoadEvaluator.ContextSwitches[0],$htmlwhite))
				}
			
				If($LoadEvaluator.CpuUtilizationEnabled)
				{
					$rowdata += @(,("CPU Utilization Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Report full load when the processor utilization % is > than",($htmlsilver -bor $htmlbold),$LoadEvaluator.CpuUtilization[1],$htmlwhite))
					$rowdata += @(,("  Report no load when the processor utilization % is <= to",($htmlsilver -bor $htmlbold),$LoadEvaluator.CpuUtilization[0],$htmlwhite))
				}
			
				If($LoadEvaluator.DiskDataIOEnabled)
				{
					$rowdata += @(,("Disk Data I/O Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Report full load when the total disk I/O in kbps is > than",($htmlsilver -bor $htmlbold),$LoadEvaluator.DiskDataIO[1],$htmlwhite))
					$rowdata += @(,("  Report no load when the total disk I/O in kbps per second is <= to",($htmlsilver -bor $htmlbold),$LoadEvaluator.DiskDataIO[0],$htmlwhite))
				}
			
				If($LoadEvaluator.DiskOperationsEnabled)
				{
					$rowdata += @(,("Disk Operations Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Report full load when the total # of R/W operations per second is > than",($htmlsilver -bor $htmlbold),$LoadEvaluator.DiskOperations[1],$htmlwhite))
					$rowdata += @(,("  Report no load when the total # of R/W operations per second is <= to",($htmlsilver -bor $htmlbold),$LoadEvaluator.DiskOperations[0],$htmlwhite))
				}
			
				If($LoadEvaluator.IPRangesEnabled)
				{
					$rowdata += @(,("IP Range Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					If($LoadEvaluator.IPRangesAllowed)
					{
						$tmp - "Allow client connections from the listed IP Ranges"
					} 
					Else 
					{
						$tmp = "Deny client connections from the listed IP Ranges"
					}
					$rowdata += @(,($tmp,($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("IP Address Ranges",($htmlsilver -bor $htmlbold),$LoadEvaluator.IPRanges[0],$htmlwhite))
					$cnt =-1
					ForEach($IPRange in $LoadEvaluator.IPRanges)
					{
						$cnt++
						If($cnt -gt 0)
						{
							$rowdata += @(,("",($htmlsilver -bor $htmlbold),$IPRange,$htmlwhite))
						}
					}
					$tmp = $Null
					$cnt = $Null
				}
			
				If($LoadEvaluator.LoadThrottlingEnabled)
				{
					Switch ($LoadEvaluator.LoadThrottling)
					{
						"Unknown"		{$tmp = "Unknown"; Break}
						"Extreme"		{$tmp = "Extreme"; Break}
						"High"			{$tmp = "High (Default)"; Break}
						"MediumHigh"	{$tmp = "Medium High"; Break}
						"Medium"		{$tmp = "Medium"; Break}
						"MediumLow"		{$tmp = "Medium Low"; Break}
						Default			{$tmp = "Impact of logons on load could not be determined: $($LoadEvaluator.LoadThrottling)"; Break}
					}
					$rowdata += @(,("Load Throttling Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Impact of logons on load",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
					$tmp = $Null
				}
			
				If($LoadEvaluator.MemoryUsageEnabled)
				{
					$rowdata += @(,("Memory Usage Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Report full load when the memory usage is > than",($htmlsilver -bor $htmlbold),$LoadEvaluator.MemoryUsage[1],$htmlwhite))
					$rowdata += @(,("  Report no load when the memory usage is <= to",($htmlsilver -bor $htmlbold),$LoadEvaluator.MemoryUsage[0],$htmlwhite))
				}
			
				If($LoadEvaluator.PageFaultsEnabled)
				{
					$rowdata += @(,("Page Faults Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Report full load when the # of page faults per second is > than",($htmlsilver -bor $htmlbold),$LoadEvaluator.PageFaults[1],$htmlwhite))
					$rowdata += @(,("  Report no load when the # of page faults per second is <= to",($htmlsilver -bor $htmlbold),$LoadEvaluator.PageFaults[0],$htmlwhite))
				}
			
				If($LoadEvaluator.PageSwapsEnabled)
				{
					$rowdata += @(,("Page Swaps Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Report full load when the # of page swaps per second is > than",($htmlsilver -bor $htmlbold),$LoadEvaluator.PageSwaps[1],$htmlwhite))
					$rowdata += @(,("  Report no load when the # of page swaps per second is <= to",($htmlsilver -bor $htmlbold),$LoadEvaluator.PageSwaps[0],$htmlwhite))
				}
			
				If($LoadEvaluator.ScheduleEnabled)
				{
					$rowdata += @(,("Scheduling Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Sunday Schedule",($htmlsilver -bor $htmlbold),$LoadEvaluator.SundaySchedule,$htmlwhite))
					$rowdata += @(,("  Monday Schedule",($htmlsilver -bor $htmlbold),$LoadEvaluator.MondaySchedule,$htmlwhite))
					$rowdata += @(,("  Tuesday Schedule",($htmlsilver -bor $htmlbold),$LoadEvaluator.TuesdaySchedule,$htmlwhite))
					$rowdata += @(,("  Wednesday Schedul",($htmlsilver -bor $htmlbold),$LoadEvaluator.WednesdaySchedule,$htmlwhite))
					$rowdata += @(,("  Thursday Schedule",($htmlsilver -bor $htmlbold),$LoadEvaluator.ThursdaySchedule,$htmlwhite))
					$rowdata += @(,("  Friday Schedule",($htmlsilver -bor $htmlbold),$LoadEvaluator.FridaySchedule,$htmlwhite))
					$rowdata += @(,("  Saturday Schedule",($htmlsilver -bor $htmlbold),$LoadEvaluator.SaturdaySchedule,$htmlwhite))
				}
			
				If($LoadEvaluator.ServerUserLoadEnabled)
				{
					$rowdata += @(,("Server User Load Settings",($htmlsilver -bor $htmlbold),"",$htmlwhite))
					$rowdata += @(,("  Report full load when the # of server users equals",($htmlsilver -bor $htmlbold),$LoadEvaluator.ServerUserLoad,$htmlwhite))
				}
			
				$msg = ""
				$columnWidths = @("325","175")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
		Else
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 $LoadEvaluator.LoadEvaluatorName
			}
			ElseIf($Text)
			{
				Line 0 $LoadEvaluator.LoadEvaluatorName
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 $LoadEvaluator.LoadEvaluatorName
			}
		}
	}
}
#endregion

#region server functions
Function ProcessServers
{
	If($Section -eq "All" -or $Section -eq "Servers")
	{
		#servers
		Write-Host "$(Get-Date): Processing Servers" -BackgroundColor Black -ForegroundColor Yellow
		[int]$Script:TotalControllers = 0
		[int]$Script:TotalWorkers = 0
		[int]$Script:TotalServers = 0
		$Script:ServerItems = @()

		Write-Host "$(Get-Date): `tRetrieving Servers" -BackgroundColor Black -ForegroundColor Yellow
		If($Summary)
		{
			$Servers = @(Get-XAServer -EA 0 | Sort-Object ServerName)
		}
		Else
		{
			$Servers = @(Get-XAServer -EA 0 | Sort-Object FolderPath, ServerName)
		}

		If($? -and $Null -ne $Servers)
		{
			OutputServer $Servers
		}
		ElseIf(!$?)
		{
			Write-Warning "Server information could not be retrieved"
		}
		Else
		{
			Write-Warning "No results returned for Server information"
		}
		$servers = $Null
		Write-Host "$(Get-Date): Finished Processing Servers" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputServer
{
	Param([object] $Servers)

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Servers:"
	}
	ElseIf($Text)
	{
		Line 0 "Servers:"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Servers:"
	}

	ForEach($server in $servers)
	{
		Write-Host "$(Get-Date): `t`tProcessing server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow

		If(!$Summary)
		{
			[bool]$SvrOnline = $False
			Write-Host "$(Get-Date): `t`t`tTesting to see if $($server.ServerName) is online and reachable" -BackgroundColor Black -ForegroundColor Yellow
			If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
			{
				$SvrOnline = $True
				If($Hardware -and $Software)
				{
					Write-Host "$(Get-Date): `t`t`t`t$($server.ServerName) is online." -BackgroundColor Black -ForegroundColor Yellow
					Write-Host "$(Get-Date): `t`t`t`tHardware and Software Inventory, Citrix Services and Hotfix areas will be processed." -BackgroundColor Black -ForegroundColor Yellow
				}
				ElseIf($Hardware -and !($Software))
				{
					Write-Host "$(Get-Date): `t`t`t`t$($server.ServerName) is online." -BackgroundColor Black -ForegroundColor Yellow
					Write-Host "$(Get-Date): `t`t`t`tHardware inventory, Citrix Services and Hotfix areas will be processed." -BackgroundColor Black -ForegroundColor Yellow
				}
				ElseIf(!($Hardware) -and $Software)
				{
					Write-Host "$(Get-Date): `t`t`t`t$($server.ServerName) is online." -BackgroundColor Black -ForegroundColor Yellow
					Write-Host "$(Get-Date): `t`t`t`tSoftware Inventory, Citrix Services and Hotfix areas will be processed." -BackgroundColor Black -ForegroundColor Yellow
				}
				Else
				{
					Write-Host "$(Get-Date): `t`t`t`t$($server.ServerName) is online." -BackgroundColor Black -ForegroundColor Yellow
					Write-Host "$(Get-Date): `t`t`t`tCitrix Services and Hotfix areas will be processed." -BackgroundColor Black -ForegroundColor Yellow
				}
			}
			
			#create array for appendix B
			Write-Host "$(Get-Date): `t`t`tGather server info for Appendix B" -BackgroundColor Black -ForegroundColor Yellow
			$obj = New-Object -TypeName PSObject
			$obj | Add-Member -MemberType NoteProperty -Name ServerName -Value $server.ServerName
			$obj | Add-Member -MemberType NoteProperty -Name ZoneName -Value $server.ZoneName
			$obj | Add-Member -MemberType NoteProperty -Name OSVersion -Value $server.OSVersion
			$obj | Add-Member -MemberType NoteProperty -Name CitrixVersion -Value $server.CitrixVersion
			$obj | Add-Member -MemberType NoteProperty -Name ProductEdition -Value $server.CitrixEdition
			$obj | Add-Member -MemberType NoteProperty -Name LicenseServer -Value $Server.LicenseServerName			

			If($SvrOnline)
			{
				$Reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $server.ServerName)
				Try
				{
					$RegKey= $Reg.OpenSubKey("SYSTEM\\CurrentControlSet\\Control\\Citrix\\Wfshell\\TWI")
					$SSDisabled = $RegKey.GetValue("SeamlessFlags")
					
					If($SSDisabled -eq 1)
					{
						$obj | Add-Member -MemberType NoteProperty -Name SessionSharing -Value "Disabled"
					}
					Else
					{
						$obj | Add-Member -MemberType NoteProperty -Name SessionSharing -Value "Enabled"
					}
				}
				Catch
				{
						$obj | Add-Member -MemberType NoteProperty -Name SessionSharing -Value "Not Available"
				}
			}
			Else
			{
				$obj | Add-Member -MemberType NoteProperty -Name SessionSharing -Value "Server Offline"
			}
			
			$Script:ServerItems += $obj

			If($MSWord -or $PDF)
			{
				WriteWordLine 2 0 $server.ServerName
				[System.Collections.Hashtable[]] $ScriptInformation = @()
				If($server.LogOnsEnabled)
				{
					$tmp = "Enabled"
				} 
				Else 
				{
					$tmp = "Disabled"
				}
				Switch ($Server.LogOnMode)
				{
					"Unknown"                       {$tmp2 = "Unknown"; Break}
					"AllowLogOns"                   {$tmp2 = "Allow logons and reconnections"; Break}
					"ProhibitNewLogOnsUntilRestart" {$tmp2 = "Prohibit logons until server restart"; Break}
					"ProhibitNewLogOns "            {$tmp2 = "Prohibit logons only"; Break}
					"ProhibitLogOns "               {$tmp2 = "Prohibit logons and reconnections"; Break}
					Default							{$tmp2 = "Logon control mode could not be determined: $($Server.LogOnMode)"; Break}
				}
				$ScriptInformation += @{ Data = "Product"; Value = $server.CitrixProductName; }
				$ScriptInformation += @{ Data = "Edition"; Value = $server.CitrixEdition; }
				$ScriptInformation += @{ Data = "Version"; Value = $server.CitrixVersion; }
				$ScriptInformation += @{ Data = "Service Pack"; Value = $server.CitrixServicePack; }
				$ScriptInformation += @{ Data = "IP Address"; Value = $server.IPAddresses; }
				$ScriptInformation += @{ Data = "Logons"; Value = $tmp; }
				$ScriptInformation += @{ Data = "Logon Control Mode"; Value = $tmp2; }
				$tmp = $Null
				$tmp2 = $Null

				Switch ($server.ElectionPreference)
				{
					"Unknown"           {$tmp = "Unknown"; Break}
					"MostPreferred"     {$tmp = "Most Preferred"; $Script:TotalControllers++; Break}
					"Preferred"         {$tmp = "Preferred"; $Script:TotalControllers++; Break}
					"DefaultPreference" {$tmp = "Default Preference"; $Script:TotalControllers++; Break}
					"NotPreferred"      {$tmp = "Not Preferred"; $Script:TotalControllers++; Break}
					"WorkerMode"        {$tmp = "Worker Mode"; $Script:TotalWorkers++; Break}
					Default				{$tmp = "Server election preference could not be determined: $($server.ElectionPreference)"; Break}
				}
				If($server.LicenseServerName)
				{
					$ScriptInformation += @{ Data = "License Server Name"; Value = $server.LicenseServerName; }
					$ScriptInformation += @{ Data = "License Server Port"; Value = $server.LicenseServerPortNumber; }
				}
				If($server.ICAPortNumber -gt 0)
				{
					$ScriptInformation += @{ Data = "ICA Port Number"; Value = $server.ICAPortNumber; }
				}
				$ScriptInformation += @{ Data = "Product Installation Date"; Value = $server.CitrixInstallDate; }
				$ScriptInformation += @{ Data = "Operating System Version"; Value = "$($server.OSVersion) $($server.OSServicePack)"; }
				$ScriptInformation += @{ Data = "Zone"; Value = $server.ZoneName; }
				$ScriptInformation += @{ Data = "Election Preference"; Value = $tmp; }
				$ScriptInformation += @{ Data = "Folder"; Value = $server.FolderPath; }
				$ScriptInformation += @{ Data = "Product Installation Path"; Value = $server.CitrixInstallPath; }
				$tmp = $Null
				
				$Table = AddWordTable -Hashtable $ScriptInformation `
				-Columns Data,Value `
				-List `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				## IB - Set the header row format
				SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				$Table.Columns.Item(1).Width = 250;
				$Table.Columns.Item(2).Width = 250;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""

				If($SvrOnline -and $Hardware)
				{
					GetComputerWMIInfo $server.ServerName
				}
				
				#applications published to server
				$Applications = @(Get-XAApplication -ServerName $server.ServerName -EA 0 | Sort-Object FolderPath, DisplayName)
				If($? -and $Null -ne $Applications)
				{
					WriteWordLine 0 1 "Published applications:"
					Write-Host "$(Get-Date): `t`tProcessing published applications for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					
					[int]$Rows = $Applications.count + 1

					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Style = $myHash.Word_TableGrid
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
					[int]$xRow = 1
					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Display name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Folder path"
					ForEach($app in $Applications)
					{
						Write-Host "$(Get-Date): `t`t`tProcessing published application $($app.DisplayName)" -BackgroundColor Black -ForegroundColor Yellow
						$xRow++
						$Table.Cell($xRow,1).Range.Text = $app.DisplayName
						$Table.Cell($xRow,2).Range.Text = $app.FolderPath
					}
					$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
					$Table.AutoFitBehavior($wdAutoFitContent)

					FindWordDocumentEnd
					WriteWordLine 0 0 ""
				}

				#get list of applications installed on server
				# original work by Shaun Ritchie
				# modified by Jeff Wouters
				# modified by Webster
				# fixed, as usual, by Michael B. Smith
				If($SvrOnline -and $Software)
				{
					#section modified on 3-jan-2014 to add displayversion
					$InstalledApps = @()
					$JustApps = @()

					#Define the variable to hold the location of Currently Installed Programs
					$UninstallKey1="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 

					#Create an instance of the Registry Object and open the HKLM base key
					$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Server.ServerName) 

					#Drill down into the Uninstall key using the OpenSubKey Method
					$regkey1=$reg.OpenSubKey($UninstallKey1) 

					#Retrieve an array of string that contain all the subkey names
					If($Null -ne $regkey1)
					{
						$subkeys1=$regkey1.GetSubKeyNames() 

						#Open each Subkey and use GetValue Method to return the required values for each
						ForEach($key in $subkeys1) 
						{
							$thisKey=$UninstallKey1+"\\"+$key 
							$thisSubKey=$reg.OpenSubKey($thisKey) 
							If(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
							{
								$obj = New-Object PSObject
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
								$InstalledApps += $obj
							}
						}
					}			

					$UninstallKey2="SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
					$regkey2=$reg.OpenSubKey($UninstallKey2)
					If($Null -ne $regkey2)
					{
						$subkeys2=$regkey2.GetSubKeyNames()

						ForEach($key in $subkeys2) 
						{
							$thisKey=$UninstallKey2+"\\"+$key 
							$thisSubKey=$reg.OpenSubKey($thisKey) 
							if(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
							{
								$obj = New-Object PSObject
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
								$InstalledApps += $obj
							}
						}
					}

					$InstalledApps = $InstalledApps | Sort-Object DisplayName

					$tmp1 = SWExclusions
					If($Tmp1 -ne "")
					{
						$Func = ConvertTo-ScriptBlock $tmp1
						$tempapps = Invoke-Command {& $Func}
					}
					Else
					{
						$tempapps = $InstalledApps
					}
					
					$JustApps = $TempApps | Select-Object DisplayName, DisplayVersion | Sort-Object DisplayName -unique

					WriteWordLine 0 1 "Installed applications:"
					Write-Host "$(Get-Date): `t`tProcessing installed applications for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					[int]$Rows = $JustApps.count + 1

					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$Table.rows.first.headingformat = $wdHeadingFormatTrue
					$Table.Style = $myHash.Word_TableGrid
					$Table.Borders.InsideLineStyle = $wdLineStyleSingle
					$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
					[int]$xRow = 1
					$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Application name"
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Application version"
					ForEach($app in $JustApps)
					{
						Write-Host "$(Get-Date): `t`t`tProcessing installed application $($app.DisplayName)" -BackgroundColor Black -ForegroundColor Yellow
						$xRow++
						$Table.Cell($xRow,1).Range.Text = $app.DisplayName
						$Table.Cell($xRow,2).Range.Text = $app.DisplayVersion
					}
					$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)
					$Table.AutoFitBehavior($wdAutoFitContent)

					FindWordDocumentEnd
					WriteWordLine 0 0 ""
				}
				
				#list citrix services
				If($SvrOnline)
				{
					Write-Host "$(Get-Date): `t`tProcessing Citrix services for server $($server.ServerName) by calling Get-Service" -BackgroundColor Black -ForegroundColor Yellow

					Try
					{
						#Iain Brighton optimization 5-Jun-2014
						#Replaced with a single call to retrieve services via WMI. The repeated
						## "Get-WMIObject Win32_Service -Filter" calls were the major delays in the script.
						## If we need to retrieve the StartUp type might as well just use WMI.
						$Services = @(Get-WMIObject Win32_Service -ComputerName $server.ServerName -EA 0 | Where-Object {$_.DisplayName -like "*Citrix*"} | Sort-Object DisplayName)
					}
					
					Catch
					{
						$Services = $Null
					}

					WriteWordLine 0 1 "Citrix Services" -NoNewLine
					If($? -and $Null -ne $Services)
					{
						[int]$NumServices = $Services.count
						Write-Host "$(Get-Date): `t`t $NumServices Services found" -BackgroundColor Black -ForegroundColor Yellow
						
						WriteWordLine 0 0 " ($NumServices Services found)"
						## IB - replacement Services table generation utilising AddWordTable function

						## Create an array of hashtables to store our services
						[System.Collections.Hashtable[]] $ServicesWordTable = @();
						## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
						[System.Collections.Hashtable[]] $HighlightedCells = @();
						## Seed the $Services row index from the second row
						[int] $CurrentServiceIndex = 2;
						
						ForEach($Service in $Services) 
						{
							#Write-Host "$(Get-Date): `t`t`t Processing service $($Service.DisplayName)";

							## Add the required key/values to the hashtable
							$WordTableRowHash = @{ DisplayName = $Service.DisplayName; Status = $Service.State; StartMode = $Service.StartMode; }

							## Add the hash to the array
							$ServicesWordTable += $WordTableRowHash;

							## Store "to highlight" cell references
							If($Service.State -like "Stopped" -and $Service.StartMode -like "Auto") 
							{
								$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
							}
							$CurrentServiceIndex++;
						}
						
						## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
						$Table = AddWordTable -Hashtable $ServicesWordTable `
						-Columns DisplayName, Status, StartMode `
						-Headers "Display Name", "Status", "Startup Type" `
						-AutoFit $wdAutoFitContent;

						## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
						## IB - Set the required highlighted cells
						SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

						#indent the entire table 1 tab stop
						$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""
					}
					ElseIf(!$?)
					{
						Write-Warning "No services were retrieved."
						WriteWordLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $False $True
						WriteWordLine 0 1 "If this is a trusted Forest, you may need to rerun the" "" $Null 0 $False $True
						WriteWordLine 0 1 "script with Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					Else
					{
						Write-Warning "Services retrieval was successful but no services were returned."
						WriteWordLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $False $True
					}

					#Citrix hotfixes installed
					Write-Host "$(Get-Date): `t`tGet list of Citrix hotfixes installed using Get-XAServerHotfix" -BackgroundColor Black -ForegroundColor Yellow
					Write-Host "$(Get-Date): `t`tGet list of Citrix hotfixes installed using Get-XAServerHotfix" -BackgroundColor Black -ForegroundColor Yellow
					Try
					{
						$hotfixes = @((Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | Where-Object {$_.Valid -eq $True}) | Sort-Object HotfixName)
					}
					
					Catch
					{
						$hotfixes = $Null
					}
					
					If($? -and $Null -ne $hotfixes)
					{
						$Rows = $Hotfixes.length
						
						Write-Host "$(Get-Date): `t`tNumber of Citrix hotfixes is $($Rows)" -BackgroundColor Black -ForegroundColor Yellow
						$HotfixArray = @()
						[bool]$HRP2Installed = $False
						[bool]$HRP3Installed = $False
						[bool]$HRP4Installed = $False
						[bool]$HRP5Installed = $False
						[bool]$HRP6Installed = $False
						[bool]$HRP7Installed = $False
						
						WriteWordLine 0 0 ""
						WriteWordLine 0 1 "Citrix Installed Hotfixes ($($Rows-1)):"
						## Create an array of hashtables to store our hotfixes
						[System.Collections.Hashtable[]] $hotfixesWordTable = @();
						## Seed the row index from the second row
						[int] $CurrentServiceIndex = 2;

						ForEach($hotfix in $hotfixes)
						{
							$HotfixArray += $hotfix.HotfixName
							Switch ($hotfix.HotfixName)
							{
								"XA650W2K8R2X64R02" {$HRP2Installed = $True; Break}
								"XA650W2K8R2X64R03" {$HRP3Installed = $True; Break}
								"XA650W2K8R2X64R04" {$HRP4Installed = $True; Break}
								"XA650W2K8R2X64R05" {$HRP5Installed = $True; Break}
								"XA650W2K8R2X64R06" {$HRP6Installed = $True; Break}
								"XA650W2K8R2X64R07" {$HRP7Installed = $True; Break}
							}
							$InstallDate = $hotfix.InstalledOn.ToString()
							
							## Add the required key/values to the hashtable
							$WordTableRowHash = @{ HotfixName = $hotfix.HotfixName; InstalledBy = $hotfix.InstalledBy; InstallDate = $InstallDate.SubString(0,$InstallDate.IndexOf(" ")); HotfixType = $hotfix.HotfixType}

							## Add the hash to the array
							$HotfixesWordTable += $WordTableRowHash;

							$CurrentServiceIndex++;
						}
						
						## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
						$Table = AddWordTable -Hashtable $HotfixesWordTable `
						-Columns HotfixName, InstalledBy, InstallDate, HotfixType `
						-Headers "Hotfix", "Installed By", "Install Date", "Type" `
						-AutoFit $wdAutoFitContent;

						SetWordCellFormat -Collection $Table -Size 10
						## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -Size 10 -BackgroundColor $wdColorGray15;

						#indent the entire table 1 tab stop
						$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

						FindWordDocumentEnd
						$Table = $Null
						WriteWordLine 0 0 ""

						#compare Citrix hotfixes to recommended Citrix hotfixes from CTX129229
						#hotfix lists are from CTX129229 dated 27-DEC-2016
						Write-Host "$(Get-Date): `t`tCompare Citrix hotfixes to recommended Citrix hotfixes from CTX129229" -BackgroundColor Black -ForegroundColor Yellow
						Write-Host "$(Get-Date): `t`tProcessing Citrix hotfix list for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
						If($HRP7Installed)
						{
							$RecommendedList = @()
						}
						ElseIf($HRP6Installed)
						{
							$RecommendedList = @("XA650R06W2K8R2X64001", "XA650R06W2K8R2X64022", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP5Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R06", "XA650R05W2K8R2X64020", "XA650R05W2K8R2X64025", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP4Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R05", "XA650W2K8R2X64R06", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP3Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R04", "XA650W2K8R2X64R05", "XA650W2K8R2X64R06", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP2Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R03", "XA650W2K8R2X64R04", "XA650W2K8R2X64R05", "XA650W2K8R2X64R06", "XA650W2K8R2X64R07")
						}
						Else
						{
							$RecommendedList = @("XA650W2K8R2X64001", "XA650W2K8R2X64011", "XA650W2K8R2X64019", "XA650W2K8R2X64025", 
												"XA650R01W2K8R2X64061", "XA650W2K8R2X64R01", "XA650W2K8R2X64R03")
						}
						
						If($RecommendedList.count -gt 0)
						{
							WriteWordLine 0 1 "Citrix Recommended Hotfixes:"
							## Create an array of hashtables to store our hotfixes
							[System.Collections.Hashtable[]] $HotfixesWordTable = @();
							## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
							[System.Collections.Hashtable[]] $HighlightedCells = @();
							## Seed the row index from the second row
							[int] $CurrentServiceIndex = 2;
							
							ForEach($element in $RecommendedList)
							{
								$Tmp = $Null
								If(!($HotfixArray -contains $element))
								{
									#missing a recommended Citrix hotfix
									$Tmp = "Not Installed"
								}
								Else
								{
									$Tmp = "Installed"
								}
								## Add the required key/values to the hashtable
								$WordTableRowHash = @{ CitrixHotfix = $element; Status = $Tmp}

								## Add the hash to the array
								$HotfixesWordTable += $WordTableRowHash;

								If($Tmp -eq "Not Installed")
								{
									## Store "to highlight" cell references
									$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
								}
								$CurrentServiceIndex++;
							}
							
							## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
							$Table = AddWordTable -Hashtable $HotfixesWordTable `
							-Columns CitrixHotfix, Status `
							-Headers "Citrix Hotfix", "Status" `
							-AutoFit $wdAutoFitContent;

							## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
							SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
							## IB - Set the required highlighted cells
							SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

							#indent the entire table 1 tab stop
							$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

							FindWordDocumentEnd
							$Table = $Null
							WriteWordLine 0 0 ""
						}
						#build list of installed Microsoft hotfixes
						Write-Host "$(Get-Date): `t`tProcessing Microsoft hotfixes for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
						[bool]$GotMSHotfixes = $True
						
						Try
						{
							$results = Get-HotFix -computername $Server.ServerName 
							$MSInstalledHotfixes = $results | select-object -Expand HotFixID | Sort-Object HotFixID
							$results = $Null
						}
						
						Catch
						{
							$GotMSHotfixes = $False
						}
						
						If($GotMSHotfixes)
						{
							If($server.OSServicePack.IndexOf('1') -gt 0)
							{
								#Server 2008 R2 SP1 installed
								$RecommendedList = @("KB2620656", "KB2647753", "KB2728738", "KB2748302", 
												"KB2775511", "KB2778831", "KB2871131", "KB2896256", 
												"KB2908190", "KB2920289", "KB917607")
							}
							Else
							{
								#Server 2008 R2 without SP1 installed
								$RecommendedList = @("KB2265716", "KB2383928", "KB2647753", "KB2728738", 
												"KB2748302", "KB2775511", "KB2778831", "KB2871131", 
												"KB2896256", "KB3014783", "KB917607", "KB975777", 
												"KB979530", "KB980663", "KB983460")
							}
							
							If($RecommendedList.count -gt 0)
							{
								WriteWordLine 0 1 "Microsoft Recommended Hotfixes (from CTX129229):"
								## Create an array of hashtables to store our hotfixes
								[System.Collections.Hashtable[]] $HotfixesWordTable = @();
								## Create an array of hashtables to store references of cells that we wish to highlight after the table has been added
								[System.Collections.Hashtable[]] $HighlightedCells = @();
								## Seed the row index from the second row
								[int] $CurrentServiceIndex = 2;

								ForEach($hotfix in $RecommendedList)
								{
									$Tmp = $Null
									If(!($MSInstalledHotfixes -contains $hotfix))
									{
										$Tmp = "Not Installed"
									}
									Else
									{
										$Tmp = "Installed"
									}
									## Add the required key/values to the hashtable
									$WordTableRowHash = @{ MicrosoftHotfix = $hotfix; Status = $Tmp}

									## Add the hash to the array
									$HotfixesWordTable += $WordTableRowHash;

									If($Tmp -eq "Not Installed")
									{
										## Store "to highlight" cell references
										$HighlightedCells += @{ Row = $CurrentServiceIndex; Column = 2; }
									}
									$CurrentServiceIndex++;
								}
								
								## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
								$Table = AddWordTable -Hashtable $HotfixesWordTable `
								-Columns MicrosoftHotfix, Status `
								-Headers "Microsoft Hotfix", "Status" `
								-AutoFit $wdAutoFitFixed;

								## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
								SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;
								## IB - Set the required highlighted cells
								SetWordCellFormat -Coordinates $HighlightedCells -Table $Table -Bold -BackgroundColor $wdColorRed -Solid;

								$Table.Columns.Item(1).Width = 125;
								$Table.Columns.Item(2).Width = 100;

								#indent the entire table 1 tab stop
								$Table.Rows.SetLeftIndent($Indent1TabStops,$wdAdjustProportional)

								FindWordDocumentEnd
								$Table = $Null
								WriteWordLine 0 1 "Not all missing Microsoft hotfixes may be needed for this server `n`tor might already be replaced and not recorded in CTX129229." -FontSize 8 -BoldFace $True
								WriteWordLine 0 0 ""
							}
						}
						Else
						{
							Write-Host "$(Get-Date): Get-HotFix failed for $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
							Write-Warning "Get-HotFix failed for $($server.ServerName)"
							WriteWordLine 0 0 "Get-HotFix failed for $($server.ServerName)" "" $Null 0 $False $True
							WriteWordLine 0 0 "On $($server.ServerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "No Citrix hotfixes were retrieved"
						WriteWordLine 0 0 "Warning: No Citrix hotfixes were retrieved" "" $Null 0 $False $True
					}
					Else
					{
						Write-Warning "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned."
						WriteWordLine 0 0 "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned." "" $Null 0 $False $True
					}
				}
				Else
				{
					Write-Host "$(Get-Date): `t`t$($server.ServerName) is offline or unreachable.  Citrix Services and Hotfix areas skipped." -BackgroundColor Black -ForegroundColor Yellow
					WriteWordLine 0 0 "Server $($server.ServerName) was offline or unreachable at "(Get-date).ToString()
					WriteWordLine 0 0 "The Citrix Services and Hotfix areas were skipped."
				}
				WriteWordLine 0 0 "" 
			}
			ElseIf($Text)
			{
				Line 0 $server.ServerName
				If($server.LogOnsEnabled)
				{
					$tmp = "Enabled"
				} 
				Else 
				{
					$tmp = "Disabled"
				}
				Switch ($Server.LogOnMode)
				{
					"Unknown"                       {$tmp2 = "Unknown"; Break}
					"AllowLogOns"                   {$tmp2 = "Allow logons and reconnections"; Break}
					"ProhibitNewLogOnsUntilRestart" {$tmp2 = "Prohibit logons until server restart"; Break}
					"ProhibitNewLogOns "            {$tmp2 = "Prohibit logons only"; Break}
					"ProhibitLogOns "               {$tmp2 = "Prohibit logons and reconnections"; Break}
					Default							{$tmp2 = "Logon control mode could not be determined: $($Server.LogOnMode)"; Break}
				}
				Line 1 "Product`t`t`t`t: " $server.CitrixProductName
				Line 1 "Edition`t`t`t`t: " $server.CitrixEdition
				Line 1 "Version`t`t`t`t: " $server.CitrixVersion
				Line 1 "Service Pack`t`t`t: " $server.CitrixServicePack
				Line 1 "IP Address`t`t`t: " $server.IPAddresses
				Line 1 "Logons`t`t`t`t: " $tmp
				Line 1 "Logon Control Mode`t`t: " $tmp2
				$tmp = $Null
				$tmp2 = $Null

				Switch ($server.ElectionPreference)
				{
					"Unknown"           {$tmp = "Unknown"; Break}
					"MostPreferred"     {$tmp = "Most Preferred"; $Script:TotalControllers++; Break}
					"Preferred"         {$tmp = "Preferred"; $Script:TotalControllers++; Break}
					"DefaultPreference" {$tmp = "Default Preference"; $Script:TotalControllers++; Break}
					"NotPreferred"      {$tmp = "Not Preferred"; $Script:TotalControllers++; Break}
					"WorkerMode"        {$tmp = "Worker Mode"; $Script:TotalWorkers++; Break}
					Default				{$tmp = "Server election preference could not be determined: $($server.ElectionPreference)"; Break}
				}
				If($server.LicenseServerName)
				{
					Line 1 "License Server Name`t`t: " $server.LicenseServerName
					Line 1 "License Server Port`t`t: " $server.LicenseServerPortNumber
				}
				If($server.ICAPortNumber -gt 0)
				{
					Line 1 "ICA Port Number`t`t`t: " $server.ICAPortNumber
				}
				Line 1 "Product Installation Date`t: " $server.CitrixInstallDate
				Line 1 "Operating System Version`t: $($server.OSVersion) $($server.OSServicePack)"
				Line 1 "Zone`t`t`t`t: " $server.ZoneName
				Line 1 "Election Preference`t`t: " $tmp
				Line 1 "Folder`t`t`t`t: " $server.FolderPath
				Line 1 "Product Installation Path`t: " $server.CitrixInstallPath
				$tmp = $Null
				Line 0 ""

				If($SvrOnline -and $Hardware)
				{
					GetComputerWMIInfo $server.ServerName
				}
				
				#applications published to server
				$Applications = @(Get-XAApplication -ServerName $server.ServerName -EA 0 | Sort-Object FolderPath, DisplayName)
				If($? -and $Null -ne $Applications)
				{
					Line 1 "Published applications:"
					Write-Host "$(Get-Date): `t`tProcessing published applications for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
					ForEach($app in $Applications)
					{
						Write-Host "$(Get-Date): `t`t`tProcessing published application $($app.DisplayName)" -BackgroundColor Black -ForegroundColor Yellow
						Line 2 "Display name`t: " $app.DisplayName
						Line 2 "Folder path`t: " $app.FolderPath
						Line 0 ""
					}
				}

				#get list of applications installed on server
				# original work by Shaun Ritchie
				# modified by Jeff Wouters
				# modified by Webster
				# fixed, as usual, by Michael B. Smith
				If($SvrOnline -and $Software)
				{
					#section modified on 3-jan-2014 to add displayversion
					$InstalledApps = @()
					$JustApps = @()

					#Define the variable to hold the location of Currently Installed Programs
					$UninstallKey1="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 

					#Create an instance of the Registry Object and open the HKLM base key
					$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Server.ServerName) 

					#Drill down into the Uninstall key using the OpenSubKey Method
					$regkey1=$reg.OpenSubKey($UninstallKey1) 

					#Retrieve an array of string that contain all the subkey names
					If($Null -ne $regkey1)
					{
						$subkeys1=$regkey1.GetSubKeyNames() 

						#Open each Subkey and use GetValue Method to return the required values for each
						ForEach($key in $subkeys1) 
						{
							$thisKey=$UninstallKey1+"\\"+$key 
							$thisSubKey=$reg.OpenSubKey($thisKey) 
							If(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
							{
								$obj = New-Object PSObject
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
								$InstalledApps += $obj
							}
						}
					}			

					$UninstallKey2="SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
					$regkey2=$reg.OpenSubKey($UninstallKey2)
					If($Null -ne $regkey2)
					{
						$subkeys2=$regkey2.GetSubKeyNames()

						ForEach($key in $subkeys2) 
						{
							$thisKey=$UninstallKey2+"\\"+$key 
							$thisSubKey=$reg.OpenSubKey($thisKey) 
							if(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
							{
								$obj = New-Object PSObject
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
								$InstalledApps += $obj
							}
						}
					}

					$InstalledApps = $InstalledApps | Sort-Object DisplayName

					$tmp1 = SWExclusions
					If($Tmp1 -ne "")
					{
						$Func = ConvertTo-ScriptBlock $tmp1
						$tempapps = Invoke-Command {& $Func}
					}
					Else
					{
						$tempapps = $InstalledApps
					}
					
					$JustApps = $TempApps | Select-Object DisplayName, DisplayVersion | Sort-Object DisplayName -unique

					Line 1 "Installed applications:"
					Write-Host "$(Get-Date): `t`tProcessing installed applications for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
					ForEach($app in $JustApps)
					{
						Write-Host "$(Get-Date): `t`t`tProcessing installed application $($app.DisplayName)" -BackgroundColor Black -ForegroundColor Yellow
						Line 2 "Application name`t: " $app.DisplayName
						Line 2 "Application version`t: " $app.DisplayVersion
						Line 0 ""
					}
				}
				
				#list citrix services
				If($SvrOnline)
				{
					Write-Host "$(Get-Date): `t`tProcessing Citrix services for server $($server.ServerName) by calling Get-Service" -BackgroundColor Black -ForegroundColor Yellow

					Try
					{
						#Iain Brighton optimization 5-Jun-2014
						#Replaced with a single call to retrieve services via WMI. The repeated
						## "Get-WMIObject Win32_Service -Filter" calls were the major delays in the script.
						## If we need to retrieve the StartUp type might as well just use WMI.
						$Services = @(Get-WMIObject Win32_Service -ComputerName $server.ServerName -EA 0 | Where-Object {$_.DisplayName -like "*Citrix*"} | Sort-Object DisplayName)
					}
					
					Catch
					{
						$Services = $Null
					}

					Line 1 "Citrix Services" -NoNewLine
					If($? -and $Null -ne $Services)
					{
						[int]$NumServices = $Services.count
						Write-Host "$(Get-Date): `t`t $NumServices Services found" -BackgroundColor Black -ForegroundColor Yellow
						
						Line 0 " ($NumServices Services found)"
						
						ForEach($Service in $Services) 
						{
							Line 2 "Display Name`t: " $Service.DisplayName
							Line 2 "Status`t`t: " $Service.State
							Line 2 "Startup Type`t: " $Service.StartMode
							Line 0 ""
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "No services were retrieved."
						Line 0 "Warning: No Services were retrieved"
						Line 1 "If this is a trusted Forest, you may need to rerun the"
						Line 1 "script with Admin credentials from the trusted Forest."
					}
					Else
					{
						Write-Warning "Services retrieval was successful but no services were returned."
						Line 0 "Services retrieval was successful but no services were returned."
					}

					#Citrix hotfixes installed
					Write-Host "$(Get-Date): `t`tGet list of Citrix hotfixes installed using Get-XAServerHotfix" -BackgroundColor Black -ForegroundColor Yellow
					Try
					{
						$hotfixes = @((Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | Where-Object {$_.Valid -eq $True}) | Sort-Object HotfixName)
					}
					
					Catch
					{
						$hotfixes = $Null
					}
					
					If($? -and $Null -ne $hotfixes)
					{
						$Rows = $Hotfixes.length
						
						Write-Host "$(Get-Date): `t`tNumber of Citrix hotfixes is $($Rows)" -BackgroundColor Black -ForegroundColor Yellow
						$HotfixArray = @()
						[bool]$HRP2Installed = $False
						[bool]$HRP3Installed = $False
						[bool]$HRP4Installed = $False
						[bool]$HRP5Installed = $False
						[bool]$HRP6Installed = $False
						[bool]$HRP7Installed = $False
						
						Line 0 ""
						Line 1 "Citrix Installed Hotfixes ($($Rows)):"

						ForEach($hotfix in $hotfixes)
						{
							$HotfixArray += $hotfix.HotfixName
							Switch ($hotfix.HotfixName)
							{
								"XA650W2K8R2X64R02" {$HRP2Installed = $True; Break}
								"XA650W2K8R2X64R03" {$HRP3Installed = $True; Break}
								"XA650W2K8R2X64R04" {$HRP4Installed = $True; Break}
								"XA650W2K8R2X64R05" {$HRP5Installed = $True; Break}
								"XA650W2K8R2X64R06" {$HRP6Installed = $True; Break}
								"XA650W2K8R2X64R07" {$HRP7Installed = $True; Break}
							}
							$InstallDate = $hotfix.InstalledOn.ToString()
							
							Line 2 "Hotfix`t`t: " $hotfix.HotfixName
							Line 2 "Installed By`t: " $hotfix.InstalledBy
							Line 2 "Install Date`t: " $InstallDate.SubString(0,$InstallDate.IndexOf(" "))
							Line 2 "Type`t`t: " $hotfix.HotfixType
							Line 0 ""
						}
						
						#compare Citrix hotfixes to recommended Citrix hotfixes from CTX129229
						#hotfix lists are from CTX129229 dated 27-DEC-2016
						Write-Host "$(Get-Date): `t`tCompare Citrix hotfixes to recommended Citrix hotfixes from CTX129229" -BackgroundColor Black -ForegroundColor Yellow
						Write-Host "$(Get-Date): `t`tProcessing Citrix hotfix list for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
						If($HRP7Installed)
						{
							$RecommendedList = @()
						}
						ElseIf($HRP6Installed)
						{
							$RecommendedList = @("XA650R06W2K8R2X64001", "XA650R06W2K8R2X64022", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP5Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R06", "XA650R05W2K8R2X64020", "XA650R05W2K8R2X64025", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP4Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R05", "XA650W2K8R2X64R06", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP3Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R04", "XA650W2K8R2X64R05", "XA650W2K8R2X64R06", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP2Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R03", "XA650W2K8R2X64R04", "XA650W2K8R2X64R05", "XA650W2K8R2X64R06", "XA650W2K8R2X64R07")
						}
						Else
						{
							$RecommendedList = @("XA650W2K8R2X64001", "XA650W2K8R2X64011", "XA650W2K8R2X64019", "XA650W2K8R2X64025", 
												"XA650R01W2K8R2X64061", "XA650W2K8R2X64R01", "XA650W2K8R2X64R03")
						}
						
						If($RecommendedList.count -gt 0)
						{
							Line 1 "Citrix Recommended Hotfixes:"
							
							ForEach($element in $RecommendedList)
							{
								$Tmp = $Null
								If(!($HotfixArray -contains $element))
								{
									#missing a recommended Citrix hotfix
									$Tmp = "Not Installed"
								}
								Else
								{
									$Tmp = "Installed"
								}
								Line 2 "Citrix Hotfix`t: " $element
								Line 2 "Status`t`t: " $Tmp
								Line 0 ""
							}
							
						}
						#build list of installed Microsoft hotfixes
						Write-Host "$(Get-Date): `t`tProcessing Microsoft hotfixes for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
						[bool]$GotMSHotfixes = $True
						
						Try
						{
							$results = Get-HotFix -computername $Server.ServerName 
							$MSInstalledHotfixes = $results | select-object -Expand HotFixID | Sort-Object HotFixID
							$results = $Null
						}
						
						Catch
						{
							$GotMSHotfixes = $False
						}
						
						If($GotMSHotfixes)
						{
							If($server.OSServicePack.IndexOf('1') -gt 0)
							{
								#Server 2008 R2 SP1 installed
								$RecommendedList = @("KB2620656", "KB2647753", "KB2728738", "KB2748302", 
												"KB2775511", "KB2778831", "KB2871131", "KB2896256", 
												"KB2908190", "KB2920289", "KB917607")
							}
							Else
							{
								#Server 2008 R2 without SP1 installed
								$RecommendedList = @("KB2265716", "KB2383928", "KB2647753", "KB2728738", 
												"KB2748302", "KB2775511", "KB2778831", "KB2871131", 
												"KB2896256", "KB3014783", "KB917607", "KB975777", 
												"KB979530", "KB980663", "KB983460")
							}
							
							If($RecommendedList.count -gt 0)
							{
								Line 1 "Microsoft Recommended Hotfixes (from CTX129229):"

								ForEach($hotfix in $RecommendedList)
								{
									$Tmp = $Null
									If(!($MSInstalledHotfixes -contains $hotfix))
									{
										$Tmp = "Not Installed"
									}
									Else
									{
										$Tmp = "Installed"
									}
									Line 2 "Microsoft Hotfix: " $hotfix
									Line 2 "Status`t`t: " $Tmp
									Line 0 ""
								}
								
								Line 1 "Not all missing Microsoft hotfixes may be needed for this server or might already be replaced and not recorded in CTX129229"
								Line 0 ""
							}
						}
						Else
						{
							Write-Host "$(Get-Date): Get-HotFix failed for $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
							Write-Warning "Get-HotFix failed for $($server.ServerName)"
							Line 0 "Get-HotFix failed for $($server.ServerName)"
							Line 0 "On $($server.ServerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "No Citrix hotfixes were retrieved"
						Line 0 "Warning: No Citrix hotfixes were retrieved"
					}
					Else
					{
						Write-Warning "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned."
						Line 0 "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned."
					}
				}
				Else
				{
					Write-Host "$(Get-Date): `t`t$($server.ServerName) is offline or unreachable.  Citrix Services and Hotfix areas skipped." -BackgroundColor Black -ForegroundColor Yellow
					Line 0 "Server $($server.ServerName) was offline or unreachable at "(Get-date).ToString()
					Line 0 "The Citrix Services and Hotfix areas were skipped."
				}
				Line 0 "" 
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 2 0 $server.ServerName
				$columnHeaders = @()
				$rowdata = @()
				If($server.LogOnsEnabled)
				{
					$tmp = "Enabled"
				} 
				Else 
				{
					$tmp = "Disabled"
				}
				Switch ($Server.LogOnMode)
				{
					"Unknown"                       {$tmp2 = "Unknown"; Break}
					"AllowLogOns"                   {$tmp2 = "Allow logons and reconnections"; Break}
					"ProhibitNewLogOnsUntilRestart" {$tmp2 = "Prohibit logons until server restart"; Break}
					"ProhibitNewLogOns "            {$tmp2 = "Prohibit logons only"; Break}
					"ProhibitLogOns "               {$tmp2 = "Prohibit logons and reconnections"; Break}
					Default							{$tmp2 = "Logon control mode could not be determined: $($Server.LogOnMode)"; Break}
				}
				$columnHeaders = @("Product",($htmlsilver -bor $htmlbold),$server.CitrixProductName,$htmlwhite)
				$rowdata += @(,("Edition",($htmlsilver -bor $htmlbold),$server.CitrixEdition,$htmlwhite))
				$rowdata += @(,("Version",($htmlsilver -bor $htmlbold),$server.CitrixVersion,$htmlwhite))
				$rowdata += @(,("Service Pack",($htmlsilver -bor $htmlbold),$server.CitrixServicePack,$htmlwhite))
				$rowdata += @(,("IP Address",($htmlsilver -bor $htmlbold),$server.IPAddresses[0].ToString(),$htmlwhite))
				$rowdata += @(,("Logons",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$rowdata += @(,("Logon Control Mode",($htmlsilver -bor $htmlbold),$tmp2,$htmlwhite))
				$tmp = $Null
				$tmp2 = $Null

				Switch ($server.ElectionPreference)
				{
					"Unknown"           {$tmp = "Unknown"; Break}
					"MostPreferred"     {$tmp = "Most Preferred"; $Script:TotalControllers++; Break}
					"Preferred"         {$tmp = "Preferred"; $Script:TotalControllers++; Break}
					"DefaultPreference" {$tmp = "Default Preference"; $Script:TotalControllers++; Break}
					"NotPreferred"      {$tmp = "Not Preferred"; $Script:TotalControllers++; Break}
					"WorkerMode"        {$tmp = "Worker Mode"; $Script:TotalWorkers++; Break}
					Default				{$tmp = "Server election preference could not be determined: $($server.ElectionPreference)"; Break}
				}
				If($server.LicenseServerName)
				{
					$rowdata += @(,("License Server Name",($htmlsilver -bor $htmlbold),$server.LicenseServerName,$htmlwhite))
					$rowdata += @(,("License Server Port",($htmlsilver -bor $htmlbold),$server.LicenseServerPortNumber,$htmlwhite))
				}
				If($server.ICAPortNumber -gt 0)
				{
					$rowdata += @(,("ICA Port Number",($htmlsilver -bor $htmlbold),$server.ICAPortNumber,$htmlwhite))
				}
				$rowdata += @(,("Product Installation Date",($htmlsilver -bor $htmlbold),$server.CitrixInstallDate,$htmlwhite))
				$rowdata += @(,("Operating System Version",($htmlsilver -bor $htmlbold),"$($server.OSVersion) $($server.OSServicePack)",$htmlwhite))
				$rowdata += @(,("Zone",($htmlsilver -bor $htmlbold),$server.ZoneName,$htmlwhite))
				$rowdata += @(,("Election Preference",($htmlsilver -bor $htmlbold),$tmp,$htmlwhite))
				$rowdata += @(,("Folder",($htmlsilver -bor $htmlbold),$server.FolderPath,$htmlwhite))
				$rowdata += @(,("Product Installation Path",($htmlsilver -bor $htmlbold),$server.CitrixInstallPath,$htmlwhite))
				$tmp = $Null
				
				$msg = ""
				$columnWidths = @("250","250")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""

				If($SvrOnline -and $Hardware)
				{
					GetComputerWMIInfo $server.ServerName
				}
				
				#applications published to server
				$Applications = @(Get-XAApplication -ServerName $server.ServerName -EA 0 | Sort-Object FolderPath, DisplayName)
				If($? -and $Null -ne $Applications)
				{
					#WriteHTMLLine 0 1 ":"
					Write-Host "$(Get-Date): `t`tProcessing published applications for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
					$rowdata = @()
					$columnHeaders = @(
					'Display name',($htmlsilver -bor $htmlbold),
					'Folder path',($htmlsilver -bor $htmlbold))
					ForEach($app in $Applications)
					{
						Write-Host "$(Get-Date): `t`t`tProcessing published application $($app.DisplayName)" -BackgroundColor Black -ForegroundColor Yellow
						$rowdata += @(,(
						$app.DisplayName,$htmlwhite,
						$app.FolderPath,$htmlwhite))
					}
					$msg = "Published applications"
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 ""
				}

				#get list of applications installed on server
				# original work by Shaun Ritchie
				# modified by Jeff Wouters
				# modified by Webster
				# fixed, as usual, by Michael B. Smith
				If($SvrOnline -and $Software)
				{
					#section modified on 3-jan-2014 to add displayversion
					$InstalledApps = @()
					$JustApps = @()

					#Define the variable to hold the location of Currently Installed Programs
					$UninstallKey1="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 

					#Create an instance of the Registry Object and open the HKLM base key
					$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Server.ServerName) 

					#Drill down into the Uninstall key using the OpenSubKey Method
					$regkey1=$reg.OpenSubKey($UninstallKey1) 

					#Retrieve an array of string that contain all the subkey names
					If($Null -ne $regkey1)
					{
						$subkeys1=$regkey1.GetSubKeyNames() 

						#Open each Subkey and use GetValue Method to return the required values for each
						ForEach($key in $subkeys1) 
						{
							$thisKey=$UninstallKey1+"\\"+$key 
							$thisSubKey=$reg.OpenSubKey($thisKey) 
							If(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
							{
								$obj = New-Object PSObject
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
								$InstalledApps += $obj
							}
						}
					}			

					$UninstallKey2="SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
					$regkey2=$reg.OpenSubKey($UninstallKey2)
					If($Null -ne $regkey2)
					{
						$subkeys2=$regkey2.GetSubKeyNames()

						ForEach($key in $subkeys2) 
						{
							$thisKey=$UninstallKey2+"\\"+$key 
							$thisSubKey=$reg.OpenSubKey($thisKey) 
							if(![String]::IsNullOrEmpty($($thisSubKey.GetValue("DisplayName")))) 
							{
								$obj = New-Object PSObject
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayName" -Value $($thisSubKey.GetValue("DisplayName"))
								$obj | Add-Member -MemberType NoteProperty -Name "DisplayVersion" -Value $($thisSubKey.GetValue("DisplayVersion"))
								$InstalledApps += $obj
							}
						}
					}

					$InstalledApps = $InstalledApps | Sort-Object DisplayName

					$tmp1 = SWExclusions
					If($Tmp1 -ne "")
					{
						$Func = ConvertTo-ScriptBlock $tmp1
						$tempapps = Invoke-Command {& $Func}
					}
					Else
					{
						$tempapps = $InstalledApps
					}
					
					$JustApps = $TempApps | Select-Object DisplayName, DisplayVersion | Sort-Object -ObjectDisplayName -unique

					#WriteHTMLLine 0 1 ":"
					Write-Host "$(Get-Date): `t`tProcessing installed applications for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
					$rowdata = @()
					$columnHeaders = @(
					'Application name',($htmlsilver -bor $htmlbold),
					'Application version',($htmlsilver -bor $htmlbold))
					ForEach($app in $JustApps)
					{
						Write-Host "$(Get-Date): `t`t`tProcessing installed application $($app.DisplayName)" -BackgroundColor Black -ForegroundColor Yellow
						$rowdata += @(,(
						$app.DisplayName,$htmlwhite,
						$app.DisplayVersion,$htmlwhite))
					}
					$msg = "Installed applications"
					FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
					WriteHTMLLine 0 0 ""
				}
				
				#list citrix services
				If($SvrOnline)
				{
					Write-Host "$(Get-Date): `t`tProcessing Citrix services for server $($server.ServerName) by calling Get-Service" -BackgroundColor Black -ForegroundColor Yellow

					Try
					{
						#Iain Brighton optimization 5-Jun-2014
						#Replaced with a single call to retrieve services via WMI. The repeated
						## "Get-WMIObject Win32_Service -Filter" calls were the major delays in the script.
						## If we need to retrieve the StartUp type might as well just use WMI.
						$Services = @(Get-WMIObject Win32_Service -ComputerName $server.ServerName -EA 0 | Where-Object {$_.DisplayName -like "*Citrix*"} | Sort-Object DisplayName)
					}
					
					Catch
					{
						$Services = $Null
					}

					If($? -and $Null -ne $Services)
					{
						[int]$NumServices = $Services.count
						Write-Host "$(Get-Date): `t`t $NumServices Services found" -BackgroundColor Black -ForegroundColor Yellow
						
						#WriteHTMLLine 0 1 ""
						
						$rowdata = @()
						$columnHeaders = @(
						'Display name',($htmlsilver -bor $htmlbold),
						'Status',($htmlsilver -bor $htmlbold),
						'Startup Type',($htmlsilver -bor $htmlbold))
						ForEach($Service in $Services) 
						{
							$rowdata += @(,(
							$Service.DisplayName,$htmlwhite,
							$Service.State,$htmlwhite,
							$Service.StartMode,$htmlwhite))
						}
						
						$msg = "Citrix Services ($NumServices Services found)"
						FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
						WriteHTMLLine 0 0 ""
					}
					ElseIf(!$?)
					{
						Write-Warning "No services were retrieved."
						WriteHTMLLine 0 1 "Citrix Services" -NoNewLine
						WriteHTMLLine 0 0 "Warning: No Services were retrieved" "" $Null 0 $False $True
						WriteHTMLLine 0 1 "If this is a trusted Forest, you may need to rerun the" "" $Null 0 $False $True
						WriteHTMLLine 0 1 "script with Admin credentials from the trusted Forest." "" $Null 0 $False $True
					}
					Else
					{
						Write-Warning "Services retrieval was successful but no services were returned."
						WriteHTMLLine 0 1 "Citrix Services" -NoNewLine
						WriteHTMLLine 0 0 "Services retrieval was successful but no services were returned." "" $Null 0 $False $True
					}

					#Citrix hotfixes installed
					Write-Host "$(Get-Date): `t`tGet list of Citrix hotfixes installed using Get-XAServerHotfix" -BackgroundColor Black -ForegroundColor Yellow
					Try
					{
						$hotfixes = @((Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | Where-Object {$_.Valid -eq $True}) | Sort-Object HotfixName)
					}
					
					Catch
					{
						$hotfixes = $Null
					}
					
					If($? -and $Null -ne $hotfixes)
					{
						$Rows = $Hotfixes.length
						
						Write-Host "$(Get-Date): `t`tNumber of Citrix hotfixes is $($Rows)" -BackgroundColor Black -ForegroundColor Yellow
						$HotfixArray = @()
						[bool]$HRP2Installed = $False
						[bool]$HRP3Installed = $False
						[bool]$HRP4Installed = $False
						[bool]$HRP5Installed = $False
						[bool]$HRP6Installed = $False
						[bool]$HRP7Installed = $False
						
						WriteHTMLLine 0 0 ""
						#WriteHTMLLine 0 1 ":"
						$rowdata = @()
						$columnHeaders = @(
						'Hotfix',($htmlsilver -bor $htmlbold),
						'Installed By',($htmlsilver -bor $htmlbold),
						'Install Date',($htmlsilver -bor $htmlbold),
						'Type',($htmlsilver -bor $htmlbold))

						ForEach($hotfix in $hotfixes)
						{
							$HotfixArray += $hotfix.HotfixName
							Switch ($hotfix.HotfixName)
							{
								"XA650W2K8R2X64R02" {$HRP2Installed = $True; Break}
								"XA650W2K8R2X64R03" {$HRP3Installed = $True; Break}
								"XA650W2K8R2X64R04" {$HRP4Installed = $True; Break}
								"XA650W2K8R2X64R05" {$HRP5Installed = $True; Break}
								"XA650W2K8R2X64R06" {$HRP6Installed = $True; Break}
								"XA650W2K8R2X64R07" {$HRP7Installed = $True; Break}
							}
							$InstallDate = $hotfix.InstalledOn.ToString()
							
							$rowdata += @(,(
							$hotfix.HotfixName,$htmlwhite,
							$hotfix.InstalledBy,$htmlwhite,
							$InstallDate.SubString(0,$InstallDate.IndexOf(" ")),$htmlwhite,
							$hotfix.HotfixType,$htmlwhite))
						}
						
						$msg = "Citrix Installed Hotfixes ($($Rows))"
						FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
						WriteHTMLLine 0 0 ""

						#compare Citrix hotfixes to recommended Citrix hotfixes from CTX129229
						#hotfix lists are from CTX129229 dated 27-DEC-2016
						Write-Host "$(Get-Date): `t`tCompare Citrix hotfixes to recommended Citrix hotfixes from CTX129229" -BackgroundColor Black -ForegroundColor Yellow
						Write-Host "$(Get-Date): `t`tProcessing Citrix hotfix list for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
						If($HRP7Installed)
						{
							$RecommendedList = @()
						}
						ElseIf($HRP6Installed)
						{
							$RecommendedList = @("XA650R06W2K8R2X64001", "XA650R06W2K8R2X64022", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP5Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R06", "XA650R05W2K8R2X64020", "XA650R05W2K8R2X64025", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP4Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R05", "XA650W2K8R2X64R06", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP3Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R04", "XA650W2K8R2X64R05", "XA650W2K8R2X64R06", "XA650W2K8R2X64R07")
						}
						ElseIf($HRP2Installed)
						{
							$RecommendedList = @("XA650W2K8R2X64R03", "XA650W2K8R2X64R04", "XA650W2K8R2X64R05", "XA650W2K8R2X64R06", "XA650W2K8R2X64R07")
						}
						Else
						{
							$RecommendedList = @("XA650W2K8R2X64001", "XA650W2K8R2X64011", "XA650W2K8R2X64019", "XA650W2K8R2X64025", 
												"XA650R01W2K8R2X64061", "XA650W2K8R2X64R01", "XA650W2K8R2X64R03")
						}
						
						If($RecommendedList.count -gt 0)
						{
							#WriteHTMLLine 0 1 ":"
							$rowdata = @()
							$columnHeaders = @(
							'Citrix Hotfix',($htmlsilver -bor $htmlbold),
							'Status',($htmlsilver -bor $htmlbold))
							
							ForEach($element in $RecommendedList)
							{
								$Tmp = $Null
								If(!($HotfixArray -contains $element))
								{
									#missing a recommended Citrix hotfix
									$Tmp = "Not Installed"
								}
								Else
								{
									$Tmp = "Installed"
								}
								$rowdata += @(,(
								$element,$htmlwhite,
								$tmp,$htmlwhite))
							}
							
							$msg = "Citrix Recommended Hotfixes"
							FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
							WriteHTMLLine 0 0 ""
						}
						#build list of installed Microsoft hotfixes
						Write-Host "$(Get-Date): `t`tProcessing Microsoft hotfixes for server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
						[bool]$GotMSHotfixes = $True
						
						Try
						{
							$results = Get-HotFix -computername $Server.ServerName 
							$MSInstalledHotfixes = $results | select-object -Expand HotFixID | Sort-Object HotFixID
							$results = $Null
						}
						
						Catch
						{
							$GotMSHotfixes = $False
						}
						
						If($GotMSHotfixes)
						{
							If($server.OSServicePack.IndexOf('1') -gt 0)
							{
								#Server 2008 R2 SP1 installed
								$RecommendedList = @("KB2620656", "KB2647753", "KB2728738", "KB2748302", 
												"KB2775511", "KB2778831", "KB2871131", "KB2896256", 
												"KB2908190", "KB2920289", "KB917607")
							}
							Else
							{
								#Server 2008 R2 without SP1 installed
								$RecommendedList = @("KB2265716", "KB2383928", "KB2647753", "KB2728738", 
												"KB2748302", "KB2775511", "KB2778831", "KB2871131", 
												"KB2896256", "KB3014783", "KB917607", "KB975777", 
												"KB979530", "KB980663", "KB983460")
							}
							
							If($RecommendedList.count -gt 0)
							{
								#WriteHTMLLine 0 1 ":"
								$rowdata = @()
								$columnHeaders = @(
								'Microsoft Hotfix',($htmlsilver -bor $htmlbold),
								'Status',($htmlsilver -bor $htmlbold))

								ForEach($hotfix in $RecommendedList)
								{
									$Tmp = $Null
									If(!($MSInstalledHotfixes -contains $hotfix))
									{
										$Tmp = "Not Installed"
									}
									Else
									{
										$Tmp = "Installed"
									}
									$rowdata += @(,(
									$hotfix,$htmlwhite,
									$Tmp,$htmlwhite))
								}
								
								$msg = "Microsoft Recommended Hotfixes (from CTX129229)"
								FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
								WriteHTMLLine 0 0 "Not all missing Microsoft hotfixes may be needed for this server `n`tor might already be replaced and not recorded in CTX129229." -FontSize 8 -BoldFace $True
								WriteHTMLLine 0 0 ""
							}
						}
						Else
						{
							Write-Host "$(Get-Date): Get-HotFix failed for $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
							Write-Warning "Get-HotFix failed for $($server.ServerName)"
							WriteHTMLLine 0 0 "Get-HotFix failed for $($server.ServerName)" "" $Null 0 $False $True
							WriteHTMLLine 0 0 "On $($server.ServerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
						}
					}
					ElseIf(!$?)
					{
						Write-Warning "No Citrix hotfixes were retrieved"
						WriteHTMLLine 0 0 "Warning: No Citrix hotfixes were retrieved" "" $Null 0 $False $True
					}
					Else
					{
						Write-Warning "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned."
						WriteHTMLLine 0 0 "Citrix hotfix retrieval was successful but no Citrix hotfixes were returned." "" $Null 0 $False $True
					}
				}
				Else
				{
					Write-Host "$(Get-Date): `t`t$($server.ServerName) is offline or unreachable.  Citrix Services and Hotfix areas skipped." -BackgroundColor Black -ForegroundColor Yellow
					WriteHTMLLine 0 0 "Server $($server.ServerName) was offline or unreachable at "(Get-date).ToString()
					WriteHTMLLine 0 0 "The Citrix Services and Hotfix areas were skipped."
				}
				WriteHTMLLine 0 0 "" 
			}

			Write-Host "$(Get-Date): `tFinished Processing server $($server.ServerName)" -BackgroundColor Black -ForegroundColor Yellow
			Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
		}
		Else
		{
			WriteWordLine 0 0 $server.ServerName
			$Script:TotalServers++
		}
	}
}
#endregion

#region worker group functions
Function ProcessWorkerGroups
{
	If($Section -eq "All" -or $Section -eq "WGs")
	{
		#worker groups
		Write-Host "$(Get-Date): Processing Worker Groups" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): `tSetting summary variables" -BackgroundColor Black -ForegroundColor Yellow
		[int]$Script:TotalWGByServerName = 0
		[int]$Script:TotalWGByServerGroup = 0
		[int]$Script:TotalWGByOU = 0
		[int]$Script:TotalWGs = 0

		Write-Host "$(Get-Date): `tRetrieving Worker Groups" -BackgroundColor Black -ForegroundColor Yellow
		$WorkerGroups = Get-XAWorkerGroup -EA 0| Sort-Object WorkerGroupName

		If($? -and $Null -ne $WorkerGroups)
		{
			If($Summary)
			{
				OutputSummaryWorkerGroups $WorkerGroups
			}
			Else
			{
				OutputWorkerGroups $WorkerGroups
			}
		}
		ElseIf($? -and $Null -eq  $WorkerGroups)
		{
			$txt = "There are no Worker Groups created"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Worker Groups"
			OutputWarning $txt
		}
		$WorkerGroups = $Null
		Write-Host "$(Get-Date): Finished Processing Worker Groups" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputSummaryWorkerGroups
{
	Param([object] $WorkerGroups)
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Worker Groups"
		[System.Collections.Hashtable[]] $WordTable = @();
		[int] $CurrentServiceIndex = 2;
	}
	ElseIf($Text)
	{
		Line 0 "Worker Groups"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Worker Groups"
		$rowdata = @()
	}
	
	ForEach($WorkerGroup in $WorkerGroups)
	{
		Write-Host "$(Get-Date): `t`tProcessing Worker Group $($WorkerGroup.WorkerGroupName)" -BackgroundColor Black -ForegroundColor Yellow
		$Script:TotalWGs++
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			WGName = $WorkerGroup.WorkerGroupName;
			}
			$WordTable += $WordTableRowHash;
			$CurrentServiceIndex++;
		}
		ElseIf($Text)
		{
			Line 0 $WorkerGroup.WorkerGroupName
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$WorkerGroup.WorkerGroupName,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns WGName `
		-Headers "Name" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Name',($htmlsilver -bor $htmlbold))

		$msg = ""
		$columnWidths = @("100")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""
	}
}

Function OutputWorkerGroups
{
	Param([object] $WorkGroups)
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Worker Groups"
	}
	ElseIf($Text)
	{
		Line 0 "Worker Groups"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Worker Groups"
	}
	
	ForEach($WorkerGroup in $WorkerGroups)
	{
		Write-Host "$(Get-Date): `t`tProcessing Worker Group $($WorkerGroup.WorkerGroupName)" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 $WorkerGroup.WorkerGroupName
			[System.Collections.Hashtable[]] $ScriptInformation = @()
			$ScriptInformation += @{ Data = "Description"; Value = $WorkerGroup.Description; }
			$ScriptInformation += @{ Data = "Folder Path"; Value = $WorkerGroup.FolderPath; }
			If($WorkerGroup.ServerNames)
			{
				$Script:TotalWGByServerName++
				Write-Host "$(Get-Date): `t`tProcessing Worker Group by Farm Servers" -BackgroundColor Black -ForegroundColor Yellow
				$TempArray = @($WorkerGroup.ServerNames | Sort-Object)
				$ScriptInformation += @{ Data = "Farm Servers"; Value = $TempArray[0]; }
				$cnt = -1
				ForEach($Item in $TempARray)
				{
					$cnt++
					
					If($cnt -gt 0)
					{
						$ScriptInformation += @{ Data = ""; Value = $Item; }
					}
				}
				$TempArray = $Null
			}
			If($WorkerGroup.ServerGroups)
			{
				$Script:TotalWGByServerGroup++
				Write-Host "$(Get-Date): `t`tProcessing Worker Group by Server Groups" -BackgroundColor Black -ForegroundColor Yellow
				$TempArray = @($WorkerGroup.ServerGroups | Sort-Object)
				$ScriptInformation += @{ Data = "Server Group Accounts"; Value = $TempArray[0]; }
				$cnt = -1
				ForEach($Item in $TempARray)
				{
					$cnt++
					
					If($cnt -gt 0)
					{
						$ScriptInformation += @{ Data = ""; Value = $Item; }
					}
				}
				$TempArray = $Null
			}
			If($WorkerGroup.OUs)
			{
				$Script:TotalWGByOU++
				Write-Host "$(Get-Date): `t`tProcessing Worker Group by OUs" -BackgroundColor Black -ForegroundColor Yellow
				$TempArray = @($WorkerGroup.OUs | Sort-Object {$_.Length})
				$ScriptInformation += @{ Data = "Container"; Value = $TempArray[0]; }
				$cnt = -1
				ForEach($Item in $TempARray)
				{
					$cnt++
					
					If($cnt -gt 0)
					{
						$ScriptInformation += @{ Data = ""; Value = $Item; }
					}
				}
				$TempArray = $Null
			}
			$Table = AddWordTable -Hashtable $ScriptInformation `
			-Columns Data,Value `
			-List `
			-Format $wdTableGrid `
			-AutoFit $wdAutoFitFixed;

			## IB - Set the header row format
			SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Columns.Item(1).Width = 150;
			$Table.Columns.Item(2).Width = 250;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
			
			#applications published to worker group
			$Applications = @(Get-XAApplication -WorkerGroup $WorkerGroup.WorkerGroupName -EA 0| Sort-Object FolderPath, DisplayName)
			If($? -and $Applications.Count -gt 0)
			{
				WriteWordLine 0 0 ""
				WriteWordLine 0 0 "Published applications for Worker Group $($WorkerGroup.WorkerGroupName)"
				Write-Host "$(Get-Date): `t`tProcessing published applications for Worker Group $($WorkerGroup.WorkerGroupName)" -BackgroundColor Black -ForegroundColor Yellow
				$TableRange = $doc.Application.Selection.Range
				[int]$Columns = 2
				
				[int]$Rows = $Applications.count + 1

				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$Table.rows.first.headingformat = $wdHeadingFormatTrue
				$Table.Style = $myHash.Word_TableGrid
				$Table.Borders.InsideLineStyle = $wdLineStyleSingle
				$Table.Borders.OutsideLineStyle = $wdLineStyleSingle
				[int]$xRow = 1
				$Table.Rows.First.Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Display name"
				$Table.Cell($xRow,2).Range.Font.Bold = $True
				$Table.Cell($xRow,2).Range.Text = "Folder path"
				ForEach($app in $Applications)
				{
					Write-Host "$(Get-Date): `t`t`tProcessing published application $($app.DisplayName)" -BackgroundColor Black -ForegroundColor Yellow
					$xRow++
					$Table.Cell($xRow,1).Range.Text = $app.DisplayName
					$Table.Cell($xRow,2).Range.Text = $app.FolderPath
				}
				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
				$Table.AutoFitBehavior($wdAutoFitContent)

				FindWordDocumentEnd
			}
		}
		ElseIf($Text)
		{
			Line 0 $WorkerGroup.WorkerGroupName
			Line 1 "Description`t: " $WorkerGroup.Description
			Line 1 "Folder Path`t: " $WorkerGroup.FolderPath
			If($WorkerGroup.ServerNames)
			{
				$Script:TotalWGByServerName++
				Write-Host "$(Get-Date): `t`tProcessing Worker Group by Farm Servers" -BackgroundColor Black -ForegroundColor Yellow
				$TempArray = @($WorkerGroup.ServerNames | Sort-Object)
				Line 1 "Farm Servers`t: " $TempArray[0]
				$cnt = -1
				ForEach($Item in $TempArray)
				{
					$cnt++
					
					If($cnt -gt 0)
					{
						Line 4 $Item
					}
				}
				$TempArray = $Null
			}
			If($WorkerGroup.ServerGroups)
			{
				$Script:TotalWGByServerGroup++
				Write-Host "$(Get-Date): `t`tProcessing Worker Group by Server Groups" -BackgroundColor Black -ForegroundColor Yellow
				$TempArray = @($WorkerGroup.ServerGroups | Sort-Object)
				Line 1 "Server Group`t: " $TempArray[0]
				$cnt = -1
				ForEach($Item in $TempArray)
				{
					$cnt++
					
					If($cnt -gt 0)
					{
						Line 4 $Item
					}
				}
				$TempArray = $Null
			}
			If($WorkerGroup.OUs)
			{
				$Script:TotalWGByOU++
				Write-Host "$(Get-Date): `t`tProcessing Worker Group by OUs" -BackgroundColor Black -ForegroundColor Yellow
				$TempArray = @($WorkerGroup.OUs | Sort-Object {$_.Length})
				Line 1 "Container`t: " $TempArray[0]
				$cnt = -1
				ForEach($Item in $TempArray)
				{
					$cnt++
					
					If($cnt -gt 0)
					{
						Line 4 $Item
					}
				}
				$TempArray = $Null
			}
			Line 0 ""
			
			#applications published to worker group
			$Applications = @(Get-XAApplication -WorkerGroup $WorkerGroup.WorkerGroupName -EA 0| Sort-Object FolderPath, DisplayName)
			If($? -and $Applications.Count -gt 0)
			{
				Line 1 "Published applications for Worker Group $($WorkerGroup.WorkerGroupName)"
				Write-Host "$(Get-Date): `t`tProcessing published applications for Worker Group $($WorkerGroup.WorkerGroupName)" -BackgroundColor Black -ForegroundColor Yellow
				ForEach($app in $Applications)
				{
					Write-Host "$(Get-Date): `t`t`tProcessing published application $($app.DisplayName)" -BackgroundColor Black -ForegroundColor Yellow
					Line 1 "Display name`t: " $app.DisplayName
					Line 1 "Folder path`t: " $app.FolderPath
					Line 0 ""
				}
			}
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 2 0 $WorkerGroup.WorkerGroupName
			$columnHeaders = @()
			$rowdata = @()
			$columnHeaders = @("Description",($htmlsilver -bor $htmlbold),$WorkerGroup.Description,$htmlwhite)
			$rowdata += @(,("Folder Path",($htmlsilver -bor $htmlbold),$WorkerGroup.FolderPath,$htmlwhite))
			If($WorkerGroup.ServerNames)
			{
				$Script:TotalWGByServerName++
				Write-Host "$(Get-Date): `t`tProcessing Worker Group by Farm Servers" -BackgroundColor Black -ForegroundColor Yellow
				$TempArray = @($WorkerGroup.ServerNames | Sort-Object)
				$rowdata += @(,("Farm Servers",($htmlsilver -bor $htmlbold),$TempArray[0],$htmlwhite))
				$cnt = -1
				ForEach($Item in $TempARray)
				{
					$cnt++
					
					If($cnt -gt 0)
					{
						$rowdata += @(,("",($htmlsilver -bor $htmlbold),$Item,$htmlwhite))
					}
				}
				$TempArray = $Null
			}
			If($WorkerGroup.ServerGroups)
			{
				$Script:TotalWGByServerGroup++
				Write-Host "$(Get-Date): `t`tProcessing Worker Group by Server Groups" -BackgroundColor Black -ForegroundColor Yellow
				$TempArray = @($WorkerGroup.ServerGroups | Sort-Object)
				$rowdata += @(,("Server Group Accounts",($htmlsilver -bor $htmlbold),$TempArray[0],$htmlwhite))
				$cnt = -1
				ForEach($Item in $TempARray)
				{
					$cnt++
					
					If($cnt -gt 0)
					{
						$rowdata += @(,("",($htmlsilver -bor $htmlbold),$Item,$htmlwhite))
					}
				}
				$TempArray = $Null
			}
			If($WorkerGroup.OUs)
			{
				$Script:TotalWGByOU++
				Write-Host "$(Get-Date): `t`tProcessing Worker Group by OUs" -BackgroundColor Black -ForegroundColor Yellow
				$TempArray = @($WorkerGroup.OUs | Sort-Object {$_.Length})
				$rowdata += @(,("Container",($htmlsilver -bor $htmlbold),$TempArray[0],$htmlwhite))
				$cnt = -1
				ForEach($Item in $TempARray)
				{
					$cnt++
					
					If($cnt -gt 0)
					{
						$rowdata += @(,("",($htmlsilver -bor $htmlbold),$Item,$htmlwhite))
					}
				}
				$TempArray = $Null
			}
			$msg = ""
			$columnWidths = @("150","250")
			FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
			
			#applications published to worker group
			$Applications = @(Get-XAApplication -WorkerGroup $WorkerGroup.WorkerGroupName -EA 0| Sort-Object FolderPath, DisplayName)
			If($? -and $Applications.Count -gt 0)
			{
				WriteHTMLLine 0 0 ""
				Write-Host "$(Get-Date): `t`tProcessing published applications for Worker Group $($WorkerGroup.WorkerGroupName)" -BackgroundColor Black -ForegroundColor Yellow
				$rowdata = @()
				$columnHeaders = @(
				'Display name',($htmlsilver -bor $htmlbold),
				'Folder path',($htmlsilver -bor $htmlbold))
				ForEach($app in $Applications)
				{
					Write-Host "$(Get-Date): `t`t`tProcessing published application $($app.DisplayName)" -BackgroundColor Black -ForegroundColor Yellow
					$rowdata += @(,(
					$app.DisplayName,$htmlwhite,
					$app.FolderPath,$htmlwhite))
				}
				$msg = "Published applications for Worker Group $($WorkerGroup.WorkerGroupName)"
				FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			}
		}
	}
}
#endregion

#region zone functions
Function ProcessZones
{
	If($Section -eq "All" -or $Section -eq "Zones")
	{
		#zones
		Write-Host "$(Get-Date): Processing Zones" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): `tSetting summary variables" -BackgroundColor Black -ForegroundColor Yellow
		[int]$Script:TotalZones = 0

		Write-Host "$(Get-Date): `tRetrieving Zones" -BackgroundColor Black -ForegroundColor Yellow
		$Zones = Get-XAZone -EA 0| Sort-Object ZoneName
		If($? -and $Null -ne $Zones)
		{
			If($Summary)
			{
				OutputSummaryZones $Zones
			}
			Else
			{
				OutputZones $Zones
			}
		}
		ElseIf($? -and $Null -eq $Zones)
		{
			$txt = "There are no Zones"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve Zones"
			OutputWarning $txt
		}
		Write-Host "$(Get-Date): Finished Processing Zones" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputSummaryZones
{
	Param([object] $Zones)

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Zones"
		[System.Collections.Hashtable[]] $WordTable = @();
		[int] $CurrentServiceIndex = 2;	}
	ElseIf($Text)
	{
		Line 0 "Zones"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Zones"
		$rowdata = @()
	}

	ForEach($Zone in $Zones)
	{
		$Script:TotalZones++
		Write-Host "$(Get-Date): `t`tProcessing Zone $($Zone.ZoneName)" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWord -or $PDF)
		{
			$WordTableRowHash = @{ 
			ZoneName = $Zone.ZoneName;
			}
			$WordTable += $WordTableRowHash;
			$CurrentServiceIndex++;
		}
		ElseIf($Text)
		{
			Line 0 $Zone.ZoneName
		}
		ElseIf($HTML)
		{
			$rowdata += @(,(
			$Zone.ZoneName,$htmlwhite))
		}
	}

	If($MSWord -or $PDF)
	{
		## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
		$Table = AddWordTable -Hashtable $WordTable `
		-Columns ZoneName `
		-Headers "Zone Name" `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		## IB - set column widths without recursion
		$Table.Columns.Item(1).Width = 100;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 ""
	}
	ElseIf($HTML)
	{
		$columnHeaders = @(
		'Zone Name',($htmlsilver -bor $htmlbold))

		$msg = ""
		FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
		WriteHTMLLine 0 0 ""
	}
}

Function OutputZones
{
	Param([object] $Zones)

	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Zones"
		[System.Collections.Hashtable[]] $WordTable = @();
		[int] $CurrentServiceIndex = 2;
	}
	ElseIf($Text)
	{
		Line 0 "Zones"
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Zones"
		$rowdata = @()
	}
	
	ForEach($Zone in $Zones)
	{
		$Script:TotalZones++
		Write-Host "$(Get-Date): `t`tProcessing Zone $($Zone.ZoneName)" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWord -or $PDF)
		{
			WriteWordLine 2 0 $Zone.ZoneName
			WriteWordLine 0 0 "Current Data Collector: " $Zone.DataCollector
			WriteWordLine 0 0 ""
		}
		ElseIf($Text)
		{
			Line 0 $Zone.ZoneName
			Line 1 "Current Data Collector: " $Zone.DataCollector
			Line 0 ""
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 2 0 $Zone.ZoneName
			WriteHTMLLine 0 0 "Current Data Collector: " $Zone.DataCollector
		}
		
		$Servers = @(Get-XAServer -ZoneName $Zone.ZoneName -EA 0| Sort-Object ElectionPreference, ServerName)
		If($? -and $Null -ne $Servers)
		{		
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 "Servers in Zone"
			}
			ElseIf($Text)
			{
				Line 1 "Servers in Zone"
			}
			ElseIf($HTML)
			{
			}
	
			ForEach($Server in $Servers)
			{
				$ElectionPref = ""
				Switch ($server.ElectionPreference)
				{
					"Unknown"           {$ElectionPref = "Unknown"; Break}
					"MostPreferred"     {$ElectionPref = "Most Preferred"; Break}
					"Preferred"         {$ElectionPref = "Preferred"; Break}
					"DefaultPreference" {$ElectionPref = "Default Preference"; Break}
					"NotPreferred"      {$ElectionPref = "Not Preferred"; Break}
					"WorkerMode"        {$ElectionPref = "Worker Mode"; Break}
					Default {$ElectionPref = "Zone preference could not be determined: $($server.ElectionPreference)"; Break}
				}
				If($MSWord -or $PDF)
				{
					$WordTableRowHash = @{ 
					ServerName = $server.ServerName;
					Pref = $ElectionPref;
					}
					$WordTable += $WordTableRowHash;
					$CurrentServiceIndex++;
				}
				ElseIf($Text)
				{
					Line 2 "Server Name`t: " $server.ServerName
					Line 2 "Preference`t: " $ElectionPref
					Line 0 ""
				}
				ElseIf($HTML)
				{
					$rowdata += @(,(
					$server.ServerName,$htmlwhite,
					$ElectionPref,$htmlwhite))
				}
			}
			If($MSWord -or $PDF)
			{
				## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
				$Table = AddWordTable -Hashtable $WordTable `
				-Columns ServerName, Pref `
				-Headers "Server Name", "Preference" `
				-Format $wdTableGrid `
				-AutoFit $wdAutoFitFixed;

				## IB - Set the header row format
				SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

				## IB - set column widths without recursion
				$Table.Columns.Item(1).Width = 150;
				$Table.Columns.Item(2).Width = 150;

				$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

				FindWordDocumentEnd
				$Table = $Null
				WriteWordLine 0 0 ""
			}
			ElseIf($HTML)
			{
				$columnHeaders = @(
				'Server Name',($htmlsilver -bor $htmlbold),
				'Preference',($htmlsilver -bor $htmlbold))

				$msg = "Servers in Zone"
				$columnWidths = @("150","150")
				FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
				WriteHTMLLine 0 0 ""
			}
		}
		ElseIf($? -and $Null -eq $Servers)
		{
			$txt = "There are no servers in the Zone"
			OutputWarning $txt
		}
		Else
		{
			$txt = "Unable to retrieve servers for this Zone"
			OutputWarning $txt
		}
		$Servers = $Null
	}
	
}
#endregion

#region policy functions
Function validStateProp( [object] $object, [string] $topLevel, [string] $secondLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			If((Get-Member -Name $secondLevel -InputObject $object.$topLevel))
			{
				Return $True
			}
		}
	}
	Return $False
}

Function validObject( [object] $object, [string] $topLevel )
{
	#function created 8-jan-2014 by Michael B. Smith
	If( $object )
	{
		If((Get-Member -Name $topLevel -InputObject $object))
		{
			Return $True
		}
	}
	Return $False
}

Function ConvertNumberToTime
{
	Param([int]$val = 0)
	
	#this is stored as a number between 0 (00:00 AM) and 1439 (23:59 PM)
	#180 = 3AM
	#900 = 3PM
	#1027 = 5:07 PM
	#[int] (1027/60) = 17 or 5PM
	#1027 % 60 leaves 7 or 7 minutes
	
	#thanks to MBS for the next line
	[int]$hour = [System.Math]::Floor(([int] $val) / ([int] 60))
	[int]$minute = $val % 60
	[string]$Strminute = $minute.ToString()
	[string]$tempminute = ""
	If($Strminute.length -lt 2)
	{
		$tempMinute = "0" + $Strminute
	}
	Else
	{
		$tempminute = $strminute
	}
	[string]$AMorPM = "AM"
	If($Hour -ge 0 -and $Hour -le 11)
	{
		$AMorPM = "AM"
	}
	Else
	{
		$AMorPM = "PM"
		If($Hour -ge 12)
		{
			$Hour = $Hour - 12
		}
	}
	Return "$($hour):$($tempminute) $($AMorPM)"
}

Function ConvertIntegerToDate
{
	#thanks to MBS for helping me on this Function
	Param([int]$DateAsInteger = 0)
	
	#this is stored as an integer but is actually a bitmask
	#01/01/2013 = 131924225 = 11111011101 00000001 00000001
	#01/17/2013 = 131924241 = 11111011101 00000001 00010001
	#
	# last 8 bits are the day
	# previous 8 bits are the month
	# the rest (up to 16) are the year
	
	[int]$year     = [Math]::Floor($DateAsInteger / 65536)
	[int]$month    = [Math]::Floor($DateAsInteger / 256) % 256
	[int]$day      = $DateAsInteger % 256

	Return "$Month/$Day/$Year"
}
	
Function ProcessPolicies
{
	Write-Host "$(Get-Date): Processing Policies" -BackgroundColor Black -ForegroundColor Yellow
	
	[int]$Script:TotalComputerPolicies = 0
	[int]$Script:TotalUserPolicies = 0
	[int]$Script:TotalIMAPolicies = 0
	[int]$Script:TotalADPolicies = 0
	[int]$Script:TotalADPoliciesNotProcessed = 0
	[int]$Script:TotalPolicies = 0

	If($Policies)
	{
		ProcessCitrixPolicies "localfarmgpo" "Computer"
		Write-Host "$(Get-Date): Finished Processing Citrix Site Computer Policies" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow

		ProcessCitrixPolicies "localfarmgpo" "User"
		Write-Host "$(Get-Date): Finished Processing Citrix Site User Policies" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
		
		If($NoADPolicies)
		{
			#don't process AD policies
		}
		Else
		{
			#thanks to the Citrix Engineering Team for helping me solve processing Citrix AD based Policies
			Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
			Write-Host "$(Get-Date): `tSee if there are any Citrix AD based policies to process" -BackgroundColor Black -ForegroundColor Yellow
			$CtxGPOArray = @()
			$CtxGPOArray = GetCtxGPOsInAD
			If($CtxGPOArray -is [Array] -and $CtxGPOArray.Count -gt 0)
			{
				Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
				Write-Host "$(Get-Date): `tThere are $($CtxGPOArray.Count) Citrix AD based policies to process" -BackgroundColor Black -ForegroundColor Yellow
				Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow

				[array]$CtxGPOArray = $CtxGPOArray | Sort-Object -unique
				
				ForEach($CtxGPO in $CtxGPOArray)
				{
					Write-Host "$(Get-Date): `tCreating ADGpoDrv PSDrive for Computer Policies" -BackgroundColor Black -ForegroundColor Yellow
					New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope Global *>$Null
		
					If(Get-PSDrive ADGpoDrv -EA 0)
					{
						Write-Host "$(Get-Date): `tProcessing Citrix AD Policy $($CtxGPO)" -BackgroundColor Black -ForegroundColor Yellow
					
						Write-Host "$(Get-Date): `tRetrieving AD Policy $($CtxGPO)" -BackgroundColor Black -ForegroundColor Yellow
						ProcessCitrixPolicies "ADGpoDrv" "Computer"
						Write-Host "$(Get-Date): Finished Processing Citrix AD Computer Policy $($CtxGPO)" -BackgroundColor Black -ForegroundColor Yellow
						Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
					}
					Else
					{
						$Script:TotalADPoliciesNotProcessed++
						Write-Warning "$($CtxGPO) is not readable by this XenApp Collector" -BackgroundColor Black -ForegroundColor Yellow
						Write-Warning "$($CtxGPO) was probably created by an updated Citrix Group Policy Provider" -BackgroundColor Black -ForegroundColor Yellow
					}

					Write-Host "$(Get-Date): `tCreating ADGpoDrv PSDrive for UserPolicies" -BackgroundColor Black -ForegroundColor Yellow
					New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope Global *>$Null
		
					If(Get-PSDrive ADGpoDrv -EA 0)
					{
						Write-Host "$(Get-Date): `tProcessing Citrix AD Policy $($CtxGPO)" -BackgroundColor Black -ForegroundColor Yellow
					
						Write-Host "$(Get-Date): `tRetrieving AD Policy $($CtxGPO)" -BackgroundColor Black -ForegroundColor Yellow
						ProcessCitrixPolicies "ADGpoDrv" "User"
						Write-Host "$(Get-Date): Finished Processing Citrix AD User Policy $($CtxGPO)" -BackgroundColor Black -ForegroundColor Yellow
						Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
					}
					Else
					{
						$Script:TotalADPoliciesNotProcessed++
						Write-Warning "$($CtxGPO) is not readable by this XenApp Collector"
						Write-Warning "$($CtxGPO) was probably created by an updated Citrix Group Policy Provider"
					}
				}
				Write-Host "$(Get-Date): Finished Processing Citrix AD Policies" -BackgroundColor Black -ForegroundColor Yellow
				Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
			}
			Else
			{
				Write-Host "$(Get-Date): There are no Citrix AD based policies to process" -BackgroundColor Black -ForegroundColor Yellow
				Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
			}
		}
	}
	Write-Host "$(Get-Date): Finished Processing Citrix Policies" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
}

Function ProcessCitrixPolicies
{
	Param([string]$xDriveName, [string]$xPolicyType)

	If($xDriveName -eq "localfarmgpo")
	{
		If($Summary)
		{
			WriteWordLine 0 0 ""
			WriteWordLine 0 0 "IMA Policies"
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 ""
				WriteWordLine 0 0 "IMA Policies"
			}
			ElseIf($Text)
			{
				Line 0 ""
				Line 0 "IMA Policies"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 ""
				WriteHTMLLine 0 0 "IMA Policies"
			}
		}
		Else
		{
		}
	}
	Else
	{
		If($Summary)
		{
			If($MSWord -or $PDF)
			{
				WriteWordLine 0 0 ""
				WriteWordLine 0 0 "Active Directory Policies"
			}
			ElseIf($Text)
			{
				Line 0 ""
				Line 0 "Active Directory Policies"
			}
			ElseIf($HTML)
			{
				WriteHTMLLine 0 0 ""
				WriteHTMLLine 0 0 "Active Directory Policies"
			}
		}
		Else
		{
		}
	}
	
	$Policies = Get-CtxGroupPolicy -Type $xPolicyType `
	-DriveName $xDriveName -EA 0 `
	| Select-Object PolicyName, Type, Description, Enabled, Priority `
	| Sort-Object Priority

	If($? -and $Null -ne $Policies)
	{
		ForEach($Policy in $Policies)
		{
			Write-Host "$(Get-Date): `tStarted $($Policy.PolicyName)" -BackgroundColor Black -ForegroundColor Yellow
			If(!$Summary)
			{
				If($xDriveName -eq "localfarmgpo")
				{
					$Script:TotalIMAPolicies++
				}
				Else
				{
					$Script:TotalADPolicies++
				}

				If($Policy.Type -eq "Computer")
				{
					$Script:TotalComputerPolicies++
				}
				Else
				{
					$Script:TotalUserPolicies++
				}
					
				If($MSWord -or $PDF)
				{
					$selection.InsertNewPage()
					If($xDriveName -eq "localfarmgpo")
					{
						WriteWordLine 2 0 $Policy.PolicyName
						WriteWordLine 0 0 "IMA Farm based policy"
					}
					Else
					{
						WriteWordLine 2 0 "$($Policy.PolicyName) in $($CtxGPO)"
						WriteWordLine 0 0 "Active Directory based policy"
					}
					[System.Collections.Hashtable[]] $ScriptInformation = @()
				
					$ScriptInformation += @{Data = "Description"; Value = $Policy.Description; }
					$ScriptInformation += @{Data = "Enabled"; Value = $Policy.Enabled; }
					$ScriptInformation += @{Data = "Type"; Value = $Policy.Type; }
					$ScriptInformation += @{Data = "Priority"; Value = $Policy.Priority; }
					
					$Table = AddWordTable -Hashtable $ScriptInformation `
					-Columns Data,Value `
					-List `
					-Format $wdTableGrid `
					-AutoFit $wdAutoFitFixed;

					SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

					$Table.Columns.Item(1).Width = 90;
					$Table.Columns.Item(2).Width = 200;

					$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

					FindWordDocumentEnd
					$Table = $Null
				}
				ElseIf($Text)
				{
					If($xDriveName -eq "localfarmgpo")
					{
						Line 0 $Policy.PolicyName
						Line 1 "IMA Farm based policy"
					}
					Else
					{
						Line 0 "$($Policy.PolicyName) in $($CtxGPO)"
						Line 1 "Active Directory based policy"
					}
					Line 1 "Description`t: " $Policy.Description
					Line 1 "Enabled`t`t: " $Policy.Enabled
					Line 1 "Type`t`t: " $Policy.Type
					Line 1 "Priority`t: " $Policy.Priority
				}
				ElseIf($HTML)
				{
					If($xDriveName -eq "localfarmgpo")
					{
						WriteHTMLLine 2 0 $Policy.PolicyName
						WriteHTMLLine 0 0"IMA Farm based policy"
					}
					Else
					{
						WriteHTMLLine 2 0 "$($Policy.PolicyName) in $($CtxGPO)"
						WriteHTMLLine 0 0 "Active Directory based policy"
					}
					$rowdata = @()
					$columnHeaders = @("Description",($htmlsilver -bor $htmlbold),$Policy.Description,$htmlwhite)
					$rowdata += @(,('Enabled',($htmlsilver -bor $htmlbold),$Policy.Enabled,$htmlwhite))
					$rowdata += @(,('Type',($htmlsilver -bor $htmlbold),$Policy.Type,$htmlwhite))
					$rowdata += @(,('Priority',($htmlsilver -bor $htmlbold),$Policy.Priority,$htmlwhite))

					$msg = ""
					$columnWidths = @("90","200")
					FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "290"
					WriteHTMLLine 0 0 " "
				}

				$filters = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName -DriveName $xDriveName -EA 0 | Sort-Object FilterType, FilterName

				If($? -and $Null -ne $Filters)
				{
					If(![String]::IsNullOrEmpty($filters))
					{
						Write-Host "$(Get-Date): `t`tProcessing all filters" -BackgroundColor Black -ForegroundColor Yellow
						$txt = "Assigned to"
						If($MSWord -or $PDF)
						{
							WriteWordLine 3 0 $txt
						}
						ElseIf($Text)
						{
							Line 0 $txt
						}
						ElseIf($HTML)
						{
							WriteHTMLLine 3 0 $txt
						}
						
						If($MSWord -or $PDF)
						{
							[System.Collections.Hashtable[]] $FiltersWordTable = @();
						}
						ElseIf($HTML)
						{
							$rowdata = @()
						}
						
						ForEach($Filter in $Filters)
						{
							$tmp = ""
							#5-May-2017 add back the WorkerGroup filter for xenapp 6.x
							Switch($filter.FilterType)
							{
								"AccessControl"  {$tmp = "Access Control"; Break}
								"BranchRepeater" {$tmp = "Citrix CloudBridge"; Break}
								"ClientIP"       {$tmp = "Client IP Address"; Break}
								"ClientName"     {$tmp = "Client Name"; Break}
								"DesktopGroup"   {$tmp = "Delivery Group"; Break}
								"DesktopKind"    {$tmp = "Delivery GroupType"; Break}
								"DesktopTag"     {$tmp = "Tag"; Break}
								"OU"             {$tmp = "Organizational Unit (OU)"; Break}
								"User"           {$tmp = "User or group"; Break}
								"WorkerGroup"    {$tmp = "Worker Group"; Break}
								Default {$tmp = "Policy Filter Type could not be determined: $($filter.FilterType)"; Break}
							}
							
							If($MSWord -or $PDF)
							{
								$FiltersWordTable += @{
								Name = $filter.FilterName;
								Type= $tmp;
								Enabled = $filter.Enabled;
								Mode = $filter.Mode;
								Value = $filter.FilterValue;
								}
							}
							ElseIf($Text)
							{
								Line 2 "Name`t: " $filter.FilterName
								Line 2 "Type`t: " $tmp
								Line 2 "Enabled`t: " $filter.Enabled
								Line 2 "Mode`t: " $filter.Mode
								Line 2 "Value`t: " $filter.FilterValue
								Line 2 ""
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$filter.FilterName,$htmlwhite,
								$tmp,$htmlwhite,
								$filter.Enabled,$htmlwhite,
								$filter.Mode,$htmlwhite,
								$filter.FilterValue,$htmlwhite))
							}
						}
						$tmp = $Null
						If($MSWord -or $PDF)
						{
							$Table = AddWordTable -Hashtable $FiltersWordTable `
							-Columns  Name,Type,Enabled,Mode,Value `
							-Headers  "Name","Type","Enabled","Mode","Value" `
							-Format $wdTableGrid `
							-AutoFit $wdAutoFitFixed;

							SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

							$Table.Columns.Item(1).Width = 115;
							$Table.Columns.Item(2).Width = 125;
							$Table.Columns.Item(3).Width = 50;
							$Table.Columns.Item(4).Width = 40;
							$Table.Columns.Item(5).Width = 170;

							$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

							FindWordDocumentEnd
							$Table = $Null
						}
						ElseIf($HTML)
						{
							$columnHeaders = @(
							'Name',($htmlsilver -bor $htmlbold),
							'Type',($htmlsilver -bor $htmlbold),
							'Enabled',($htmlsilver -bor $htmlbold),
							'Mode',($htmlsilver -bor $htmlbold),
							'Value',($htmlsilver -bor $htmlbold))

							$msg = ""
							$columnWidths = @("115","125","50","40","170")
							FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "500"
							WriteHTMLLine 0 0 " "
						}
					}
					Else
					{
						If($MSWord -or $PDF)
						{
							WriteWordLine 0 0 "Assigned to: None"
						}
						ElseIf($Text)
						{
							Line 0 "Assigned to`t`t: None"
						}
						ElseIf($HTML)
						{
							WriteHTMLLine 0 0 "Assigned to: None"
						}
					}
				}
				ElseIf($? -and $Null -eq $Filters)
				{
					$txt = "$($Policy.PolicyName) policy applies to all objects in the Farm"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 "Assigned to"
						WriteWordLine 0 0 $txt
					}
					ElseIf($Text)
					{
						Line 0 "Assigned to: " $txt
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 3 0 "Assigned to"
						WriteHTMLLine 0 0 $txt
					}
				}
				ElseIf($? -and $Policy.PolicyName -eq "Unfiltered")
				{
					$txt = "Unfiltered policy applies to all objects in the Farm"
					If($MSWord -or $PDF)
					{
						WriteWordLine 3 0 "Assigned to"
						WriteWordLine 0 0 $txt
					}
					ElseIf($Text)
					{
						Line 0 "Assigned to: " $txt
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 3 0 "Assigned to"
						WriteHTMLLine 0 0 $txt
					}
				}
				Else
				{
					$txt = "Unable to retrieve Filter settings"
					If($MSWord -or $PDF)
					{
						WriteWordLine 0 0 $txt
					}
					ElseIf($Text)
					{
						Line 0 $txt
					}
					ElseIf($HTML)
					{
						WriteHTMLLine 0 0 $txt
					}
				}

				$Settings = @(Get-CtxGroupPolicyConfiguration -PolicyName $Policy.PolicyName -DriveName $xDriveName -EA 0)
					
				If($? -and $Null -ne $Settings)
				{
					If($MSWord -or $PDF)
					{
						[System.Collections.Hashtable[]] $SettingsWordTable = @();
					}
					ElseIf($HTML)
					{
						$rowdata = @()
					}
				}
				
				$First = $True
				ForEach($Setting in $Settings)
				{
					If($First)
					{
						$txt = "Policy settings"
						If($MSWord -or $PDF)
						{
							WriteWordLine 3 0 $txt
						}
						ElseIf($Text)
						{
							Line 1 $txt
						}
						ElseIf($HTML)
						{
							WriteHTMLLine 3 0 $txt
						}
					}
					$First = $False
					
					Write-Host "$(Get-Date): `t`tPolicy settings" -BackgroundColor Black -ForegroundColor Yellow
					Write-Host "$(Get-Date): `t`t`tConnector for Configuration Manager 2012" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting AdvanceWarningFrequency State ) -and ($Setting.AdvanceWarningFrequency.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning frequency interval"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AdvanceWarningFrequency.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AdvanceWarningFrequency.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AdvanceWarningFrequency.Value
						}
					}
					If((validStateProp $Setting AdvanceWarningMessageBody State ) -and ($Setting.AdvanceWarningMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning message box body text"
						$tmpArray = $Setting.AdvanceWarningMessageBody.Value.Split("`n")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "`t`t`t`t`t`t`t`t`t      " $tmp
								}
							}
							$txt = ""
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting AdvanceWarningMessageTitle State ) -and ($Setting.AdvanceWarningMessageTitle.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning message box title"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AdvanceWarningMessageTitle.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AdvanceWarningMessageTitle.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AdvanceWarningMessageTitle.Value
						}
					}
					If((validStateProp $Setting AdvanceWarningPeriod State ) -and ($Setting.AdvanceWarningPeriod.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Advance warning time period"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AdvanceWarningPeriod.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AdvanceWarningPeriod.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AdvanceWarningPeriod.Value 
						}
					}
					If((validStateProp $Setting PvsImageUpdateDeadlinePeriod State ) -and ($Setting.PvsImageUpdateDeadlinePeriod.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Deadline calculation time for newly available PVS images: "
						$tmp = $Setting.PvsImageUpdateDeadlinePeriod.Value
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FinalForceLogoffMessageBody State ) -and ($Setting.FinalForceLogoffMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Final force logoff message box body text"
						$tmpArray = $Setting.FinalForceLogoffMessageBody.Value.Split("`n")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "`t`t`t`t`t`t`t`t`t" $tmp
								}
							}
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting FinalForceLogoffMessageTitle State ) -and ($Setting.FinalForceLogoffMessageTitle.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Final force logoff message box title"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FinalForceLogoffMessageTitle.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FinalForceLogoffMessageTitle.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FinalForceLogoffMessageTitle.Value 
						}
					}
					If((validStateProp $Setting ForceLogoffGracePeriod State ) -and ($Setting.ForceLogoffGracePeriod.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Force logoff grace period"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ForceLogoffGracePeriod.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ForceLogoffGracePeriod.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ForceLogoffGracePeriod.Value 
						}
					}
					If((validStateProp $Setting ForceLogoffMessageBody State ) -and ($Setting.ForceLogoffMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Force logoff message box body text"
						$tmpArray = $Setting.ForceLogoffMessageBody.Value.Split("`n")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $TmpArray)
						{
							If($Null -eq $Thing)
							{
								$Thing = ''
							}
							$cnt++
							$tmp = "$($Thing) "
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "`t`t`t`t`t`t`t`t`t   " $tmp
								}
							}
							$txt = ""
						}
						$TmpArray = $Null
						$tmp = $Null
					}
					If((validStateProp $Setting ForceLogoffMessageTitle State ) -and ($Setting.ForceLogoffMessageTitle.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Force logoff message box title"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ForceLogoffMessageTitle.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ForceLogoffMessageTitle.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ForceLogoffMessageTitle.Value 
						}
					}
					If((validStateProp $Setting PvsIntegrationEnabled State ) -and ($Setting.PvsIntegrationEnabled.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\PVS Integration enabled"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.PvsIntegrationEnabled.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PvsIntegrationEnabled.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PvsIntegrationEnabled.State 
						}
					}
					If((validStateProp $Setting RebootMessageBody State ) -and ($Setting.RebootMessageBody.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Reboot message box body text"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RebootMessageBody.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RebootMessageBody.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RebootMessageBody.Value 
						}
					}
					If((validStateProp $Setting AgentTaskInterval State ) -and ($Setting.AgentTaskInterval.State -ne "NotConfigured"))
					{
						$txt = "Connector for Configuration Manager 2012\Regular time interval at which the agent task is to run"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AgentTaskInterval.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AgentTaskInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AgentTaskInterval.Value 
						}
					}
					
					Write-Host "$(Get-Date): `t`t`tICA" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting ClipboardRedirection State ) -and ($Setting.ClipboardRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Client clipboard redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClipboardRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClipboardRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClipboardRedirection.State 
						}
					}
					If((validStateProp $Setting DesktopLaunchForNonAdmins State ) -and ($Setting.DesktopLaunchForNonAdmins.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop launches"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DesktopLaunchForNonAdmins.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DesktopLaunchForNonAdmins.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DesktopLaunchForNonAdmins.State 
						}
					}
					If((validStateProp $Setting IcaListenerTimeout State ) -and ($Setting.IcaListenerTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\ICA listener connection timeout (milliseconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaListenerTimeout.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaListenerTimeout.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaListenerTimeout.Value 
						}
					}
					If((validStateProp $Setting IcaListenerPortNumber State ) -and ($Setting.IcaListenerPortNumber.State -ne "NotConfigured"))
					{
						$txt = "ICA\ICA listener port number"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaListenerPortNumber.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaListenerPortNumber.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaListenerPortNumber.Value 
						}
					}
					If((validStateProp $Setting NonPublishedProgramLaunching State ) -and ($Setting.NonPublishedProgramLaunching.State -ne "NotConfigured"))
					{
						$txt = "ICA\Launching of non-published programs during client connection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.NonPublishedProgramLaunching.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.NonPublishedProgramLaunching.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.NonPublishedProgramLaunching.State
						}
					}
					
					Write-Host "$(Get-Date): `t`t`tICA\Adobe Flash Delivery\Flash Redirection" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting FlashAcceleration State ) -and ($Setting.FlashAcceleration.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash acceleration"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FlashAcceleration.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FlashAcceleration.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FlashAcceleration.State 
						}
					}
					If((validStateProp $Setting FlashUrlColorList State ) -and ($Setting.FlashUrlColorList.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash background color list"
						If(validStateProp $Setting FlashUrlColorList Values )
						{
							$Values = $Setting.FlashUrlColorList.Values
							$tmp = ""
							$cnt = 0
							ForEach($Value in $Values)
							{
								If($Null -eq $Value)
								{
									$Value = ''
								}
								$cnt++
								$tmp = "$($Value)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = $txt;
										Value = $tmp;
										}
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp 
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmp = $Null
							$Values = $Null
						}
						Else
						{
							$tmp = "No Flash background color list were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp 
							}
						}
					}
					If((validStateProp $Setting FlashBackwardsCompatibility State ) -and ($Setting.FlashBackwardsCompatibility.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash backwards compatibility"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FlashBackwardsCompatibility.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FlashBackwardsCompatibility.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FlashBackwardsCompatibility.State 
						}
					}
					If((validStateProp $Setting FlashDefaultBehavior State ) -and ($Setting.FlashDefaultBehavior.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash default behavior"
						$tmp = ""
						Switch ($Setting.FlashDefaultBehavior.Value)
						{
							"Block"		{$tmp = "Block Flash player"; Break}
							"Disable"	{$tmp = "Disable Flash acceleration"; Break}
							"Enable"	{$tmp = "Enable Flash acceleration"; Break}
							Default		{$tmp = "Flash Default behavior could not be determined: $($Setting.FlashDefaultBehavior.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting FlashEventLogging State ) -and ($Setting.FlashEventLogging.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash event logging"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FlashEventLogging.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FlashEventLogging.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FlashEventLogging.State 
						}
					}
					If((validStateProp $Setting FlashIntelligentFallback State ) -and ($Setting.FlashIntelligentFallback.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash intelligent fallback"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FlashIntelligentFallback.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FlashIntelligentFallback.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FlashIntelligentFallback.State 
						}
					}
					If((validStateProp $Setting FlashLatencyThreshold State ) -and ($Setting.FlashLatencyThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash latency threshold (milliseconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FlashLatencyThreshold.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FlashLatencyThreshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FlashLatencyThreshold.Value 
						}
					}
					If((validStateProp $Setting FlashServerSideContentFetchingWhitelist State ) -and ($Setting.FlashServerSideContentFetchingWhitelist.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash server-side content fetching URL list"
						If(validStateProp $Setting FlashServerSideContentFetchingWhitelist Values )
						{
							$Values = $Setting.FlashServerSideContentFetchingWhitelist.Values
							$tmp = ""
							$cnt = 0
							ForEach($Value in $Values)
							{
								If($Null -eq $Value)
								{
									$Value = ''
								}
								$cnt++
								$tmp = "$($Value)"
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = $txt;
										Value = $tmp;
										}
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp 
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$tmp = $Null
							$Values = $Null
						}
						Else
						{
							$tmp = "No Flash server-side content fetching URL list were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp 
							}
						}
					}
					If((validStateProp $Setting FlashUrlCompatibilityList State ) -and ($Setting.FlashUrlCompatibilityList.State -ne "NotConfigured"))
					{
						$txt = "ICA\Adobe Flash Delivery\Flash Redirection\Flash URL compatibility list"
						If(validStateProp $Setting FlashUrlCompatibilityList Values )
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = "";
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								"",$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt
							}
							$Values = $Setting.FlashUrlCompatibilityList.Values
							$tmp = ""
							ForEach($Value in $Values)
							{
								$Items = $Value.Split(' ')
								$Action = $Items[0]
								If($Action -eq "CLIENT")
								{
									$Action = "Render On Client"
								}
								ElseIf($Action -eq "SERVER")
								{
									$Action = "Render On Server"
								}
								ElseIf($Action -eq "BLOCK")
								{
									$Action = "BLOCK           "
								}
								$Url = $Items[1]
								If($Items.Count -eq 3)
								{
									$FlashInstance = $Items[2]
								}
								Else
								{
									$FlashInstance = "Any"
								}
								$tmp = "Action: $($Action)"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
								$tmp = "URL Pattern: $($Url)"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
								$tmp = "Flash Instance: $($FlashInstance)"
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
							$Values = $Null
							$Action = $Null
							$Url = $Null
							$FlashInstance = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Flash URL compatibility list were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Audio" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting AllowRtpAudio State ) -and ($Setting.AllowRtpAudio.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Audio over UDP real-time transport"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowRtpAudio.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowRtpAudio.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowRtpAudio.State 
						}
					}
					If((validStateProp $Setting AudioPlugNPlay State ) -and ($Setting.AudioPlugNPlay.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Audio Plug N Play"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AudioPlugNPlay.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AudioPlugNPlay.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AudioPlugNPlay.State 
						}
					}
					If((validStateProp $Setting AudioQuality State ) -and ($Setting.AudioQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Audio quality"
						$tmp = ""
						Switch ($Setting.AudioQuality.Value)
						{
							"Low"		{$tmp = "Low - for low-speed connections"; Break}
							"Medium"	{$tmp = "Medium - optimized for speech"; Break}
							"High"		{$tmp = "High - high definition audio"; Break}
							Default		{$tmp = "Audio quality could not be determined: $($Setting.AudioQuality.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ClientAudioRedirection State ) -and ($Setting.ClientAudioRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Client audio redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientAudioRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientAudioRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientAudioRedirection.State 
						}
					}
					If((validStateProp $Setting MicrophoneRedirection State ) -and ($Setting.MicrophoneRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Audio\Client microphone redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MicrophoneRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MicrophoneRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MicrophoneRedirection.State 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Auto Client Reconnect" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting AutoClientReconnect State ) -and ($Setting.AutoClientReconnect.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AutoClientReconnect.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoClientReconnect.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AutoClientReconnect.State 
						}
					}
					If((validStateProp $Setting AutoClientReconnectAuthenticationRequired  State ) -and ($Setting.AutoClientReconnectAuthenticationRequired.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect authentication"
						$tmp = ""
						Switch ($Setting.AutoClientReconnectAuthenticationRequired.Value)
						{
							"DoNotRequireAuthentication" {$tmp = "Do not require authentication"; Break}
							"RequireAuthentication"      {$tmp = "Require authentication"; Break}
							Default {$tmp = "Auto client reconnect authentication could not be determined: $($Setting.AutoClientReconnectAuthenticationRequired.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting AutoClientReconnectLogging State ) -and ($Setting.AutoClientReconnectLogging.State -ne "NotConfigured"))
					{
						$txt = "ICA\Auto Client Reconnect\Auto client reconnect logging"
						$tmp = ""
						Switch ($Setting.AutoClientReconnectLogging.Value)
						{
							"DoNotLogAutoReconnectEvents" {$tmp = "Do Not Log auto-reconnect events"; Break}
							"LogAutoReconnectEvents"      {$tmp = "Log auto-reconnect events"; Break}
							Default {$tmp = "Auto client reconnect logging could not be determined: $($Setting.AutoClientReconnectLogging.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					
					Write-Host "$(Get-Date): `t`t`tICA\Bandwidth" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting AudioBandwidthLimit State ) -and ($Setting.AudioBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Audio redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AudioBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AudioBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AudioBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting AudioBandwidthPercent State ) -and ($Setting.AudioBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Audio redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AudioBandwidthPercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AudioBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AudioBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting USBBandwidthLimit State ) -and ($Setting.USBBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Client USB device redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.USBBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.USBBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.USBBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting USBBandwidthPercent State ) -and ($Setting.USBBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Client USB device redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.USBBandwidthPercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.USBBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.USBBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting ClipboardBandwidthLimit State ) -and ($Setting.ClipboardBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Clipboard redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClipboardBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClipboardBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClipboardBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting ClipboardBandwidthPercent State ) -and ($Setting.ClipboardBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Clipboard redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClipboardBandwidthPercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClipboardBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClipboardBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting ComPortBandwidthLimit State ) -and ($Setting.ComPortBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\COM port redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ComPortBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ComPortBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ComPortBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting ComPortBandwidthPercent State ) -and ($Setting.ComPortBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\COM port redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ComPortBandwidthPercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ComPortBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ComPortBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting FileRedirectionBandwidthLimit State ) -and ($Setting.FileRedirectionBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\File redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FileRedirectionBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FileRedirectionBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FileRedirectionBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting FileRedirectionBandwidthPercent State ) -and ($Setting.FileRedirectionBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\File redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FileRedirectionBandwidthPercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FileRedirectionBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FileRedirectionBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting HDXMultimediaBandwidthLimit State ) -and ($Setting.HDXMultimediaBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.HDXMultimediaBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HDXMultimediaBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.HDXMultimediaBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting HDXMultimediaBandwidthPercent State ) -and ($Setting.HDXMultimediaBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.HDXMultimediaBandwidthPercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HDXMultimediaBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.HDXMultimediaBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting LptBandwidthLimit State ) -and ($Setting.LptBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\LPT port redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LptBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LptBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LptBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting LptBandwidthLimitPercent State ) -and ($Setting.LptBandwidthLimitPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\LPT port redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LptBandwidthLimitPercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LptBandwidthLimitPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LptBandwidthLimitPercent.Value 
						}
					}
					If((validStateProp $Setting OverallBandwidthLimit State ) -and ($Setting.OverallBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Overall session bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.OverallBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OverallBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.OverallBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting PrinterBandwidthLimit State ) -and ($Setting.PrinterBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Printer redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.PrinterBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PrinterBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PrinterBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting PrinterBandwidthPercent State ) -and ($Setting.PrinterBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\Printer redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.PrinterBandwidthPercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PrinterBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PrinterBandwidthPercent.Value 
						}
					}
					If((validStateProp $Setting TwainBandwidthLimit State ) -and ($Setting.TwainBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\TWAIN device redirection bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TwainBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TwainBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TwainBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting TwainBandwidthPercent State ) -and ($Setting.TwainBandwidthPercent.State -ne "NotConfigured"))
					{
						$txt = "ICA\Bandwidth\TWAIN device redirection bandwidth limit %"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TwainBandwidthPercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TwainBandwidthPercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TwainBandwidthPercent.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Client Sensors\Location" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting AllowLocationServices State ) -and ($Setting.AllowLocationServices.State -ne "NotConfigured"))
					{
						$txt = "ICA\Client Sensors\Location\Allow applications to use the physical location of the client device"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AllowLocationServices.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AllowLocationServices.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AllowLocationServices.State 
						}
					}
					
					Write-Host "$(Get-Date): `t`t`tICA\Desktop UI" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting AeroRedirection State ) -and ($Setting.AeroRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Aero Redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AeroRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AeroRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AeroRedirection.State 
						}
					}
					If((validStateProp $Setting GraphicsQuality State ) -and ($Setting.GraphicsQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Desktop Composition graphics quality"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.GraphicsQuality.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.GraphicsQuality.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.GraphicsQuality.Value 
						}
					}
					If((validStateProp $Setting AeroRedirection State ) -and ($Setting.AeroRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Desktop Composition Redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AeroRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AeroRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AeroRedirection.State 
						}
					}
					If((validStateProp $Setting DesktopWallpaper State ) -and ($Setting.DesktopWallpaper.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Desktop wallpaper"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DesktopWallpaper.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DesktopWallpaper.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DesktopWallpaper.State 
						}
					}
					If((validStateProp $Setting MenuAnimation State ) -and ($Setting.MenuAnimation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\Menu animation"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MenuAnimation.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MenuAnimation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MenuAnimation.State 
						}
					}
					If((validStateProp $Setting WindowContentsVisibleWhileDragging State ) -and ($Setting.WindowContentsVisibleWhileDragging.State -ne "NotConfigured"))
					{
						$txt = "ICA\Desktop UI\View window contents while dragging"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WindowContentsVisibleWhileDragging.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WindowContentsVisibleWhileDragging.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.WindowContentsVisibleWhileDragging.State 
						}
					}
			
					Write-Host "$(Get-Date): `t`t`tICA\End User Monitoring" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting IcaRoundTripCalculation State ) -and ($Setting.IcaRoundTripCalculation.State -ne "NotConfigured"))
					{
						$txt = "ICA\End User Monitoring\ICA round trip calculation"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaRoundTripCalculation.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaRoundTripCalculation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaRoundTripCalculation.State 
						}
					}
					If((validStateProp $Setting IcaRoundTripCalculationInterval State ) -and ($Setting.IcaRoundTripCalculationInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\End User Monitoring\ICA round trip calculation interval (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaRoundTripCalculationInterval.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaRoundTripCalculationInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaRoundTripCalculationInterval.Value 
						}	
					}
					If((validStateProp $Setting IcaRoundTripCalculationWhenIdle State ) -and ($Setting.IcaRoundTripCalculationWhenIdle.State -ne "NotConfigured"))
					{
						$txt = "ICA\End User Monitoring\ICA round trip calculations for idle connections"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaRoundTripCalculationWhenIdle.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaRoundTripCalculationWhenIdle.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaRoundTripCalculationWhenIdle.State 
						}	
					}

					Write-Host "$(Get-Date): `t`t`tICA\File Redirection" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting AutoConnectDrives State ) -and ($Setting.AutoConnectDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Auto connect client drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AutoConnectDrives.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoConnectDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AutoConnectDrives.State 
						}
					}
					If((validStateProp $Setting ClientDriveRedirection State ) -and ($Setting.ClientDriveRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client drive redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientDriveRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientDriveRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientDriveRedirection.State 
						}
					}
					If((validStateProp $Setting ClientFixedDrives State ) -and ($Setting.ClientFixedDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client fixed drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientFixedDrives.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientFixedDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientFixedDrives.State 
						}
					}
					If((validStateProp $Setting ClientFloppyDrives State ) -and ($Setting.ClientFloppyDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client floppy drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientFloppyDrives.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientFloppyDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientFloppyDrives.State 
						}
					}
					If((validStateProp $Setting ClientNetworkDrives State ) -and ($Setting.ClientNetworkDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client network drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientNetworkDrives.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientNetworkDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientNetworkDrives.State 
						}
					}
					If((validStateProp $Setting ClientOpticalDrives State ) -and ($Setting.ClientOpticalDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client optical drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientOpticalDrives.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientOpticalDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientOpticalDrives.State 
						}
					}
					If((validStateProp $Setting ClientRemoveableDrives State ) -and ($Setting.ClientRemoveableDrives.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Client removable drives"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientRemoveableDrives.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientRemoveableDrives.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientRemoveableDrives.State 
						}
					}
					If((validStateProp $Setting HostToClientRedirection State ) -and ($Setting.HostToClientRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Host to client redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.HostToClientRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HostToClientRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.HostToClientRedirection.State 
						}
					}
					If((validStateProp $Setting ClientDriveLetterPreservation State ) -and ($Setting.ClientDriveLetterPreservation.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Preserve client drive letters"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientDriveLetterPreservation.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientDriveLetterPreservation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientDriveLetterPreservation.State 
						}
					}
					If((validStateProp $Setting ReadOnlyMappedDrive State ) -and ($Setting.ReadOnlyMappedDrive.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Read-only client drive access"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ReadOnlyMappedDrive.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ReadOnlyMappedDrive.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ReadOnlyMappedDrive.State 
						}
					}
					If((validStateProp $Setting SpecialFolderRedirection State ) -and ($Setting.SpecialFolderRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Special folder redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SpecialFolderRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SpecialFolderRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SpecialFolderRedirection.State 
						}
					}
					If((validStateProp $Setting AsynchronousWrites State ) -and ($Setting.AsynchronousWrites.State -ne "NotConfigured"))
					{
						$txt = "ICA\File Redirection\Use asynchronous writes"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AsynchronousWrites.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AsynchronousWrites.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AsynchronousWrites.State 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Graphics" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting DisplayMemoryLimit State ) -and ($Setting.DisplayMemoryLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Display memory limit (KB)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DisplayMemoryLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DisplayMemoryLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DisplayMemoryLimit.Value 
						}	
					}
					If((validStateProp $Setting DisplayDegradePreference State ) -and ($Setting.DisplayDegradePreference.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Display mode degrade preference"
						$tmp = ""
						Switch ($Setting.DisplayDegradePreference.Value)
						{
							"ColorDepth"	{$tmp = "Degrade color depth first"; Break}
							"Resolution"	{$tmp = "Degrade resolution first"; Break}
							Default			{$tmp = "Display mode degrade preference could not be determined: $($Setting.DisplayDegradePreference.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}	
						$tmp = $Null
					}
					If((validStateProp $Setting DynamicPreview State ) -and ($Setting.DynamicPreview.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Dynamic windows preview"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DynamicPreview.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DynamicPreview.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DynamicPreview.State 
						}	
					}
					If((validStateProp $Setting ImageCaching State ) -and ($Setting.ImageCaching.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Image caching"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ImageCaching.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ImageCaching.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ImageCaching.State 
						}	
					}
					If((validStateProp $Setting MaximumColorDepth State ) -and ($Setting.MaximumColorDepth.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Maximum allowed color depth"
						$tmp = ""
						Switch ($Setting.MaximumColorDepth.Value)
						{
							"BitsPerPixel24"	{$tmp = "24 Bits Per Pixel"; Break}
							"BitsPerPixel32"	{$tmp = "32 Bits Per Pixel"; Break}
							"BitsPerPixel16"	{$tmp = "16 Bits Per Pixel"; Break}
							"BitsPerPixel15"	{$tmp = "15 Bits Per Pixel"; Break}
							"BitsPerPixel8"		{$tmp = "8 Bits Per Pixel"; Break}
							Default				{$tmp = "Maximum allowed color depth could not be determined: $($Setting.MaximumColorDepth.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}	
						$tmp = $Null
					}
					If((validStateProp $Setting DisplayDegradeUserNotification State ) -and ($Setting.DisplayDegradeUserNotification.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Notify user when display mode is degraded"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DisplayDegradeUserNotification.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DisplayDegradeUserNotification.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DisplayDegradeUserNotification.State 
						}	
					}
					If((validStateProp $Setting QueueingAndTossing State ) -and ($Setting.QueueingAndTossing.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Queueing and tossing"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.QueueingAndTossing.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.QueueingAndTossing.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.QueueingAndTossing.State 
						}	
					}

					Write-Host "$(Get-Date): `t`t`tICA\Graphics\Caching" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting PersistentCache State ) -and ($Setting.PersistentCache.State -ne "NotConfigured"))
					{
						$txt = "ICA\Graphics\Caching\Persistent cache threshold (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.PersistentCache.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.PersistentCache.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.PersistentCache.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Keep Alive" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting IcaKeepAliveTimeout State ) -and ($Setting.IcaKeepAliveTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\Keep Alive\ICA keep alive timeout (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IcaKeepAliveTimeout.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IcaKeepAliveTimeout.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IcaKeepAliveTimeout.Value 
						}
					}
					If((validStateProp $Setting IcaKeepAlives State ) -and ($Setting.IcaKeepAlives.State -ne "NotConfigured"))
					{
						$txt = "ICA\Keep Alive\ICA keep alives"
						$tmp = ""
						Switch ($Setting.IcaKeepAlives.Value)
						{
							"DoNotSendKeepAlives" {$tmp = "Do not send ICA keep alive messages"; Break}
							"SendKeepAlives"      {$tmp = "Send ICA keep alive messages"; Break}
							Default {$tmp = "ICA keep alives could not be determined: $($Setting.IcaKeepAlives.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Host "$(Get-Date): `t`t`tICA\Mobile Experience" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting AutoKeyboardPopUp State ) -and ($Setting.AutoKeyboardPopUp.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Automatic keyboard display"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AutoKeyboardPopUp.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AutoKeyboardPopUp.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AutoKeyboardPopUp.State 
						}
					}
					If((validStateProp $Setting MobileDesktop State ) -and ($Setting.MobileDesktop.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Launch touch-optimized desktop"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MobileDesktop.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MobileDesktop.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MobileDesktop.State 
						}
					}
					If((validStateProp $Setting ComboboxRemoting State ) -and ($Setting.ComboboxRemoting.State -ne "NotConfigured"))
					{
						$txt = "ICA\Mobile Experience\Remote the combo box"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ComboboxRemoting.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ComboboxRemoting.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ComboboxRemoting.State 
						}
					}
					
					Write-Host "$(Get-Date): `t`t`tICA\Multimedia" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting MultimediaConferencing State ) -and ($Setting.MultimediaConferencing.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Multimedia conferencing"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultimediaConferencing.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaConferencing.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaConferencing.State 
						}
					}
					If((validStateProp $Setting MultimediaAcceleration State ) -and ($Setting.MultimediaAcceleration.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows Media redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultimediaAcceleration.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaAcceleration.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaAcceleration.State 
						}
					}
					If((validStateProp $Setting MultimediaAccelerationDefaultBufferSize State ) -and ($Setting.MultimediaAccelerationDefaultBufferSize.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows Media redirection buffer size (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultimediaAccelerationDefaultBufferSize.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaAccelerationDefaultBufferSize.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaAccelerationDefaultBufferSize.Value 
						}
					}
					If((validStateProp $Setting MultimediaAccelerationUseDefaultBufferSize State ) -and ($Setting.MultimediaAccelerationUseDefaultBufferSize.State -ne "NotConfigured"))
					{
						$txt = "ICA\Multimedia\Windows Media redirection buffer size use"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultimediaAccelerationUseDefaultBufferSize.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultimediaAccelerationUseDefaultBufferSize.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultimediaAccelerationUseDefaultBufferSize.State 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Multi-Stream Connections" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting UDPAudioOnServer State ) -and ($Setting.UDPAudioOnServer.State -ne "NotConfigured"))
					{
						$txt = "ICA\MultiStream Connections\Audio over UDP"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UDPAudioOnServer.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UDPAudioOnServer.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UDPAudioOnServer.State
						}
					}
					If((validStateProp $Setting MultiPortPolicy State ) -and ($Setting.MultiPortPolicy.State -ne "NotConfigured"))
					{
						$txt1 = "ICA\MultiStream Connections\Multi-Port Policy\CGP default port"
						$txt2 = "ICA\MultiStream Connections\Multi-Port Policy\CGP default port priority"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt1;
							Value = "Default Port";
							}

							$SettingsWordTable += @{
							Text = $txt2;
							Value = "High";
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt1,$htmlbold,
							"Default Port",$htmlwhite))

							$rowdata += @(,(
							$txt2,$htmlbold,
							"High",$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt1 "Default Port"
							OutputPolicySetting $txt2 "High"
						}
						$txt1 = $Null
						$txt2 = $Null
						[string]$Tmp = $Setting.MultiPortPolicy.Value
						If($Tmp.Length -gt 0)
						{
							$Port1Priority = ""
							$Port2Priority = ""
							$Port3Priority = ""
							[string]$cgpport1 = $Tmp.substring(0, $Tmp.indexof(";"))
							[string]$cgpport2 = $Tmp.substring($cgpport1.length + 1 , ($Tmp.indexof(";")+1))
							[string]$cgpport3 = $Tmp.substring((($cgpport1.length + 1)+($cgpport2.length + 1)) , ($Tmp.indexof(";")+1))
							[string]$cgpport1priority = $cgpport1.substring($cgpport1.length -1, 1)
							[string]$cgpport2priority = $cgpport2.substring($cgpport2.length -1, 1)
							[string]$cgpport3priority = $cgpport3.substring($cgpport3.length -1, 1)
							$cgpport1 = $cgpport1.substring(0, $cgpport1.indexof(","))
							$cgpport2 = $cgpport2.substring(0, $cgpport2.indexof(","))
							$cgpport3 = $cgpport3.substring(0, $cgpport3.indexof(","))
							Switch ($cgpport1priority)
							{
								"0"	{$Port1Priority = "Very High"; Break}
								"2"	{$Port1Priority = "Medium"; Break}
								"3"	{$Port1Priority = "Low"; Break}
								Default	{$Port1Priority = "Unknown"; Break}
							}
							Switch ($cgpport2priority)
							{
								"0"	{$Port2Priority = "Very High"; Break}
								"2"	{$Port2Priority = "Medium"; Break}
								"3"	{$Port2Priority = "Low"; Break}
								Default	{$Port2Priority = "Unknown"; Break}
							}
							Switch ($cgpport3priority)
							{
								"0"	{$Port3Priority = "Very High"; Break}
								"2"	{$Port3Priority = "Medium"; Break}
								"3"	{$Port3Priority = "Low"; Break}
								Default	{$Port3Priority = "Unknown"; Break}
							}
							$txt1 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port1"
							$txt2 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port1 priority"
							$txt3 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port2"
							$txt4 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port2 priority"
							$txt5 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port3"
							$txt6 = "ICA\MultiStream Connections\Multi-Port Policy\CGP port3 priority"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt1;
								Value = $cgpport1;
								}

								$SettingsWordTable += @{
								Text = $txt2;
								Value = $port1priority;
								}

								$SettingsWordTable += @{
								Text = $txt3;
								Value = $cgpport2;
								}

								$SettingsWordTable += @{
								Text = $txt4;
								Value = $port2priority;
								}

								$SettingsWordTable += @{
								Text = $txt5;
								Value = $cgpport3;
								}

								$SettingsWordTable += @{
								Text = $txt6;
								Value = $port3priority;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt1,$htmlbold,
								$cgpport1,$htmlwhite))
								
								$rowdata += @(,(
								$txt2,$htmlbold,
								$port1priority,$htmlwhite))
								
								$rowdata += @(,(
								$txt3,$htmlbold,
								$cgpport2,$htmlwhite))
								
								$rowdata += @(,(
								$txt4,$htmlbold,
								$port2priority,$htmlwhite))
								
								$rowdata += @(,(
								$txt5,$htmlbold,
								$cgpport3,$htmlwhite))
								
								$rowdata += @(,(
								$txt6,$htmlbold,
								$port3priority,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt1 $cgpport1
								OutputPolicySetting $txt2 $port1priority
								OutputPolicySetting $txt3 $cgpport2
								OutputPolicySetting $txt4 $port2priority
								OutputPolicySetting $txt5 $cgpport3
								OutputPolicySetting $txt6 $port3priority
							}	
						}
						$Tmp = $Null
						$cgpport1 = $Null
						$cgpport2 = $Null
						$cgpport3 = $Null
						$cgpport1priority = $Null
						$cgpport2priority = $Null
						$cgpport3priority = $Null
						$Port1Priority = $Null
						$Port2Priority = $Null
						$Port3Priority = $Null
						$txt1 = $Null
						$txt2 = $Null
						$txt3 = $Null
						$txt4 = $Null
						$txt5 = $Null
						$txt6 = $Null
					}
					If((validStateProp $Setting MultiStreamPolicy State ) -and ($Setting.MultiStreamPolicy.State -ne "NotConfigured"))
					{
						$txt = "ICA\MultiStream Connections\Multi-Stream computer setting"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultiStreamPolicy.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultiStreamPolicy.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultiStreamPolicy.State 
						}
					}
					If((validStateProp $Setting MultiStream State ) -and ($Setting.MultiStream.State -ne "NotConfigured"))
					{
						$txt = "ICA\MultiStream Connections\Multi-Stream user setting"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MultiStream.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MultiStream.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MultiStream.State 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Port Redirection" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting ClientComPortsAutoConnection State ) -and ($Setting.ClientComPortsAutoConnection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Auto connect client COM ports"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientComPortsAutoConnection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientComPortsAutoConnection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientComPortsAutoConnection.State 
						}
					}
					If((validStateProp $Setting ClientLptPortsAutoConnection State ) -and ($Setting.ClientLptPortsAutoConnection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Auto connect client LPT ports"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientLptPortsAutoConnection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientLptPortsAutoConnection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientLptPortsAutoConnection.State 
						}
					}
					If((validStateProp $Setting ClientComPortRedirection State ) -and ($Setting.ClientComPortRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Client COM port redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientComPortRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientComPortRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientComPortRedirection.State 
						}
					}
					If((validStateProp $Setting ClientLptPortRedirection State ) -and ($Setting.ClientLptPortRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Port Redirection\Client LPT port redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientLptPortRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientLptPortRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientLptPortRedirection.State 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Printing" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting ClientPrinterRedirection State ) -and ($Setting.ClientPrinterRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client printer redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ClientPrinterRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ClientPrinterRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ClientPrinterRedirection.State 
						}
					}
					If((validStateProp $Setting DefaultClientPrinter State ) -and ($Setting.DefaultClientPrinter.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Default printer - Choose client's Default printer"
						$tmp = ""
						Switch ($Setting.DefaultClientPrinter.Value)
						{
							"ClientDefault" {$tmp = "Set Default printer to the client's main printer"; Break}
							"DoNotAdjust"   {$tmp = "Do not adjust the user's Default printer"; Break}
							Default {$tmp = "Default printer could not be determined: $($Setting.DefaultClientPrinter.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting PrinterAssignments State ) -and ($Setting.PrinterAssignments.State -ne "NotConfigured"))
					{
						If($Setting.PrinterAssignments.State -eq "Enabled")
						{
							$txt = "ICA\Printing\Printer assignments"
							$PrinterAssign = Get-ChildItem -path "$($xDriveName):\User\$($Policy.PolicyName)\Settings\ICA\Printing\PrinterAssignments"
							If($? -and $Null -ne $PrinterAssign)
							{
								$PrinterAssignments = $PrinterAssign.Contents
								ForEach($PrinterAssignment in $PrinterAssignments)
								{
									$Client = @()
									$DefaultPrinter = ""
									$SessionPrinters = @()
									$tmp1 = ""
									$tmp2 = ""
									$tmp3 = ""
									
									ForEach($Filter in $PrinterAssignment.Filters)
									{
										$Client += "$($Filter); "
									}
									
									Switch ($PrinterAssignment.DefaultPrinterOption)
									{
										"ClientDefault"		{$DefaultPrinter = "Client main printer"; Break}
										"NotConfigured"		{$DefaultPrinter = "<Not set>"; Break}
										"DoNotAdjust"		{$DefaultPrinter = "Do not adjust"; Break}
										"SpecificPrinter"	{$DefaultPrinter = $PrinterAssignment.SpecificDefaultPrinter; Break}
										Default				{$DefaultPrinter = "<Not set>"; Break}
									}
									
									ForEach($SessionPrinter in $PrinterAssignment.SessionPrinters)
									{
										$SessionPrinters += $SessionPrinter
									}
									
									$tmp1 = "Client Names/IP's: $($Client)"
									$tmp2 = "Default Printer  : $($DefaultPrinter)"
									$tmp3 = "Session Printers : $($SessionPrinters)"
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = $txt;
										Value = $tmp1;
										}
										
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp2;
										}
										
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp3;
										}
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp1,$htmlwhite))
										
										$rowdata += @(,(
										"",$htmlbold,
										$tmp2,$htmlwhite))
										
										$rowdata += @(,(
										"",$htmlbold,
										$tmp3,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp1
										OutputPolicySetting "`t`t`t`t" $tmp2
										OutputPolicySetting "`t`t`t`t" $tmp3
									}
									$tmp1 = $Null
									$tmp2 = $Null
									$tmp3 = $Null
								}
							}
						}
						Else
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $Setting.PrinterAssignments.State;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$Setting.PrinterAssignments.State,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $Setting.PrinterAssignments.State 
							}
						}
					}
					If((validStateProp $Setting AutoCreationEventLogPreference State ) -and ($Setting.AutoCreationEventLogPreference.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Printer auto-creation event log preference"
						$tmp = ""
						Switch ($Setting.AutoCreationEventLogPreference.Value)
						{
							"LogErrorsOnly"        {$tmp = "Log errors only"; Break}
							"LogErrorsAndWarnings" {$tmp = "Log errors and warnings"; Break}
							"DoNotLog"             {$tmp = "Do not log errors or warnings"; Break}
							Default {$tmp = "Printer auto-creation event log preference could not be determined: $($Setting.AutoCreationEventLogPreference.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting SessionPrinters State ) -and ($Setting.SessionPrinters.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Session printers"
						If(validStateProp $Setting SessionPrinters Values )
						{
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = "";
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								"",$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt ""
							}
							$valArray = $Setting.SessionPrinters.Values
							$tmp = ""
							ForEach($printer in $valArray)
							{
								$prArray = $printer.Split(',')
								ForEach($element in $prArray)
								{
									If($element.SubString(0, 2) -eq "\\")
									{
										$index = $element.SubString(2).IndexOf('\')
										If($index -ge 0)
										{
											$server = $element.SubString(0, $index + 2)
											$share  = $element.SubString($index + 3)
											$tmp = "Server: $($server)"
											If($MSWord -or $PDF)
											{
												$SettingsWordTable += @{
												Text = "";
												Value = $tmp;
												}
											}
											ElseIf($HTML)
											{
												$rowdata += @(,(
												"",$htmlbold,
												$tmp,$htmlwhite))
											}
											ElseIf($Text)
											{
												OutputPolicySetting "" $tmp
											}
											$tmp = "Shared Name: $($share)"
											If($MSWord -or $PDF)
											{
												$SettingsWordTable += @{
												Text = "";
												Value = $tmp;
												}
											}
											ElseIf($HTML)
											{
												$rowdata += @(,(
												"",$htmlbold,
												$tmp,$htmlwhite))
											}
											ElseIf($Text)
											{
												OutputPolicySetting "" $tmp
											}
										}
										$index = $Null
									}
									Else
									{
										$tmp1 = $element.SubString(0, 4)
										$tmp = Get-PrinterModifiedSettings $tmp1 $element
										If(![String]::IsNullOrEmpty($tmp))
										{
											If($MSWord -or $PDF)
											{
												$SettingsWordTable += @{
												Text = "";
												Value = $tmp;
												}
											}
											ElseIf($HTML)
											{
												$rowdata += @(,(
												"",$htmlbold,
												$tmp,$htmlwhite))
											}
											ElseIf($Text)
											{
												OutputPolicySetting "" $tmp
											}
										}
										$tmp1 = $Null
										$tmp = $Null
									}
								}
							}

							$valArray = $Null
							$prArray = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Session printers were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}
					If((validStateProp $Setting WaitForPrintersToBeCreated State ) -and ($Setting.WaitForPrintersToBeCreated.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Wait for printers to be created (server desktop)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WaitForPrintersToBeCreated.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WaitForPrintersToBeCreated.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.WaitForPrintersToBeCreated.State 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Printing\Client Printers" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting ClientPrinterAutoCreation State ) -and ($Setting.ClientPrinterAutoCreation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Auto-create client printers"
						$tmp = ""
						Switch ($Setting.ClientPrinterAutoCreation.Value)
						{
							"DoNotAutoCreate"    {$tmp = "Do not auto-create client printers"; Break}
							"DefaultPrinterOnly" {$tmp = "Auto-create the client's Default printer only"; Break}
							"LocalPrintersOnly"  {$tmp = "Auto-create local (non-network) client printers only"; Break}
							"AllPrinters"        {$tmp = "Auto-create all client printers"; Break}
							Default {$tmp = "Auto-create client printers could not be determined: $($Setting.ClientPrinterAutoCreation.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp
						}
						$tmp = $Null
					}
					If((validStateProp $Setting GenericUniversalPrinterAutoCreation State ) -and ($Setting.GenericUniversalPrinterAutoCreation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Auto-create generic universal printer"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.GenericUniversalPrinterAutoCreation.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.GenericUniversalPrinterAutoCreation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.GenericUniversalPrinterAutoCreation.State 
						}
					}
					If((validStateProp $Setting ClientPrinterNames State ) -and ($Setting.ClientPrinterNames.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Client printer names"
						$tmp = ""
						Switch ($Setting.ClientPrinterNames.Value)
						{
							"StandardPrinterNames" {$tmp = "Standard printer names"; Break}
							"LegacyPrinterNames"   {$tmp = "Legacy printer names"; Break}
							Default {$tmp = "Client printer names could not be determined: $($Setting.ClientPrinterNames.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting DirectConnectionsToPrintServers State ) -and ($Setting.DirectConnectionsToPrintServers.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Direct connections to print servers"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DirectConnectionsToPrintServers.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DirectConnectionsToPrintServers.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DirectConnectionsToPrintServers.State 
						}
					}
					If((validStateProp $Setting PrinterDriverMappings State ) -and ($Setting.PrinterDriverMappings.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Printer driver mapping and compatibility"
						If(validStateProp $Setting PrinterDriverMappings Values )
						{
							$array = $Setting.PrinterDriverMappings.Values
							$tmp = $array[0]
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp
							}
							
							$cnt = -1
							ForEach($element in $array)
							{
								$cnt++
								
								If($cnt -ne 0)
								{
									$Items = $element.Split(',')
									$DriverName = $Items[0]
									$Action = $Items[1]
									If($Action -match 'Replace=')
									{
										$ServerDriver = $Action.substring($Action.indexof("=")+1)
										$Action = "Replace "
									}
									Else
									{
										$ServerDriver = ""
										If($Action -eq "Allow")
										{
											$Action = "Allow "
										}
										ElseIf($Action -eq "Deny")
										{
											$Action = "Do not create "
										}
										ElseIf($Action -eq "UPD_Only")
										{
											$Action = "Create with universal driver "
										}
									}
									$tmp = "Driver Name: $($DriverName)"
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
									}
									$tmp = "Action     : $($Action)"
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
									}
									$tmp = "Settings   : "
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
									}
									If($Items.count -gt 2)
									{
										[int]$BeginAt = 2
										[int]$EndAt = $Items.count
										for ($i=$BeginAt;$i -lt $EndAt; $i++) 
										{
											$tmp2 = $Items[$i].SubString(0, 4)
											$tmp = Get-PrinterModifiedSettings $tmp2 $Items[$i]
											If(![String]::IsNullOrEmpty($tmp))
											{
												If($MSWord -or $PDF)
												{
													$SettingsWordTable += @{
													Text = "";
													Value = $tmp;
													}
												}
												ElseIf($HTML)
												{
													$rowdata += @(,(
													"",$htmlbold,
													$tmp,$htmlwhite))
												}
												ElseIf($Text)
												{
													OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
												}
											}
										}
									}
									Else
									{
										$tmp = "Unmodified "
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										ElseIf($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										ElseIf($Text)
										{
											OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
										}
									}

									If(![String]::IsNullOrEmpty($ServerDriver))
									{
										$tmp = "Server Driver: $($ServerDriver)"
										If($MSWord -or $PDF)
										{
											$SettingsWordTable += @{
											Text = "";
											Value = $tmp;
											}
										}
										ElseIf($HTML)
										{
											$rowdata += @(,(
											"",$htmlbold,
											$tmp,$htmlwhite))
										}
										ElseIf($Text)
										{
											OutputPolicySetting "`t`t`t`t`t`t`t`t     " $tmp
										}
									}
									$tmp = $Null
								}
							}
						}
						Else
						{
							$tmp = "No Printer driver mapping and compatibility were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}
					If((validStateProp $Setting PrinterPropertiesRetention State ) -and ($Setting.PrinterPropertiesRetention.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Printer properties retention"
						$tmp = ""
						Switch ($Setting.PrinterPropertiesRetention.Value)
						{
							"SavedOnClientDevice"   {$tmp = "Saved on the client device only"; Break}
							"RetainedInUserProfile" {$tmp = "Retained in user profile only"; Break}
							"FallbackToProfile"     {$tmp = "Held in profile only if not saved on client"; Break}
							"DoNotRetain"           {$tmp = "Do not retain printer properties"; Break}
							Default {$tmp = "Printer properties retention could not be determined: $($Setting.PrinterPropertiesRetention.Value)"; Break}
						}

						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting RetainedAndRestoredClientPrinters State ) -and ($Setting.RetainedAndRestoredClientPrinters.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Client Printers\Retained and restored client printers"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RetainedAndRestoredClientPrinters.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RetainedAndRestoredClientPrinters.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RetainedAndRestoredClientPrinters.State 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Printing\Drivers" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting InboxDriverAutoInstallation State ) -and ($Setting.InboxDriverAutoInstallation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Drivers\Automatic installation of in-box printer drivers"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.InboxDriverAutoInstallation.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.InboxDriverAutoInstallation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.InboxDriverAutoInstallation.State 
						}
					}
					If((validStateProp $Setting UniversalDriverPriority State ) -and ($Setting.UniversalDriverPriority.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Drivers\Universal driver preference"
						$Values = $Setting.UniversalDriverPriority.Value.Split(';')
						$tmp = ""
						$cnt = 0
						ForEach($Value in $Values)
						{
							If($Null -eq $Value)
							{
								$Value = ''
							}
							$cnt++
							$tmp = "$($Value)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp 
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "`t`t`t`t`t`t" $tmp
								}
							}
						}
						$tmp = $Null
						$Values = $Null
					}
					If((validStateProp $Setting UniversalPrintDriverUsage State ) -and ($Setting.UniversalPrintDriverUsage.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Drivers\Universal print driver usage"
						$tmp = ""
						Switch ($Setting.UniversalPrintDriverUsage.Value)
						{
							"SpecificOnly"       {$tmp = "Use only printer model specific drivers"; Break}
							"UpdOnly"            {$tmp = "Use universal printing only"; Break}
							"FallbackToUpd"      {$tmp = "Use universal printing only if requested driver is unavailable"; Break}
							"FallbackToSpecific" {$tmp = "Use printer model specific drivers only if universal printing is unavailable"; Break}
							Default {$tmp = "Universal print driver usage could not be determined: $($Setting.UniversalPrintDriverUsage.Value)"; Break}
						}

						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Host "$(Get-Date): `t`t`tICA\Printing\Universal Print Server" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting UpsEnable State ) -and ($Setting.UpsEnable.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server enable"
						If($Setting.UpsEnable.State)
						{
							$tmp = ""
						}
						Else
						{
							$tmp = "Disabled"
						}
						Switch ($Setting.UpsEnable.Value)
						{
							"UpsEnabledWithFallback"	{$tmp = "Enabled with fallback to Windows' native remote printing"; Break}
							"UpsOnlyEnabled"			{$tmp = "Enabled with no fallback to Windows' native remote printing"; Break}
							"UpsDisabled"				{$tmp = "Disabled"; Break}
							Default	{$tmp = "Universal Print Server enable value could not be determined: $($Setting.UpsEnable.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting UpsCgpPort State ) -and ($Setting.UpsCgpPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server print data stream (CGP) port"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpsCgpPort.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpsCgpPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UpsCgpPort.Value 
						}
					}
					If((validStateProp $Setting UpsPrintStreamInputBandwidthLimit State ) -and ($Setting.UpsPrintStreamInputBandwidthLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server print stream input bandwidth limit (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpsPrintStreamInputBandwidthLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpsPrintStreamInputBandwidthLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UpsPrintStreamInputBandwidthLimit.Value 
						}
					}
					If((validStateProp $Setting UpsHttpPort State ) -and ($Setting.UpsHttpPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Print Server\Universal Print Server web service (HTTP/SOAP) port"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UpsHttpPort.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UpsHttpPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UpsHttpPort.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Printing\Universal Printing" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting EMFProcessingMode State ) -and ($Setting.EMFProcessingMode.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing EMF processing mode"
						$tmp = ""
						Switch ($Setting.EMFProcessingMode.Value)
						{
							"ReprocessEMFsForPrinter" {$tmp = "Reprocess EMFs for printer"; Break}
							"SpoolDirectlyToPrinter"  {$tmp = "Spool directly to printer"; Break}
							Default {$tmp = "Universal printing EMF processing mode could not be determined: $($Setting.EMFProcessingMode.Value)"; Break}
						}
						 
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ImageCompressionLimit State ) -and ($Setting.ImageCompressionLimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing image compression limit"
						$tmp = ""
						Switch ($Setting.ImageCompressionLimit.Value)
						{
							"NoCompression"       {$tmp = "No compression"; Break}
							"LosslessCompression" {$tmp = "Best quality (lossless compression)"; Break}
							"MinimumCompression"  {$tmp = "High quality"; Break}
							"MediumCompression"   {$tmp = "Standard quality"; Break}
							"MaximumCompression"  {$tmp = "Reduced quality (maximum compression)"; Break}
							Default {$tmp = "Universal printing image compression limit could not be determined: $($Setting.ImageCompressionLimit.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting UPDCompressionDefaults State ) -and ($Setting.UPDCompressionDefaults.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing optimization defaults"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = "";
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt "" 
						}
						
						$TmpArray = $Setting.UPDCompressionDefaults.Value.Split(',')
						$tmp = ""
						ForEach($Thing in $TmpArray)
						{
							$TestLabel = $Thing.substring(0, $Thing.indexof("="))
							$TestSetting = $Thing.substring($Thing.indexof("=")+1)
							$TxtLabel = ""
							$TxtSetting = ""
							Switch($TestLabel)
							{
								"ImageCompression"
								{
									$TxtLabel = "Desired image quality:"
									Switch($TestSetting)
									{
										"StandardQuality"	{$TxtSetting = "Standard quality"; Break}
										"BestQuality"		{$TxtSetting = "Best quality (lossless compression)"; Break}
										"HighQuality"		{$TxtSetting = "High quality"; Break}
										"ReducedQuality"	{$TxtSetting = "Reduced quality (maximum compression)"; Break}
									}
								}
								"HeavyweightCompression"
								{
									$TxtLabel = "Enable heavyweight compression:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
								"ImageCaching"
								{
									$TxtLabel = "Allow caching of embedded images:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
								"FontCaching"
								{
									$TxtLabel = "Allow caching of embedded fonts:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
								"AllowNonAdminsToModify"
								{
									$TxtLabel = "Allow non-administrators to modify these settings:"
									If($TestSetting -eq "True")
									{
										$TxtSetting = "Yes"
									}
									Else
									{
										$TxtSetting = "No"
									}
								}
							}
							$tmp = "$($TxtLabel) $TxtSetting "
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = "";
								Value = $tmp;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting "`t`t`t`t`t`t`t`t`t" $tmp
							}
						}
						$TmpArray = $Null
						$tmp = $Null
						$TestLabel = $Null
						$TestSetting = $Null
						$TxtLabel = $Null
						$TxtSetting = $Null
					}
					If((validStateProp $Setting UniversalPrintingPreviewPreference State ) -and ($Setting.UniversalPrintingPreviewPreference.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing preview preference"
						$tmp = ""
						Switch ($Setting.UniversalPrintingPreviewPreference.Value)
						{
							"NoPrintPreview"        {$tmp = "Do not use print preview for auto-created or generic universal printers"; Break}
							"AutoCreatedOnly"       {$tmp = "Use print preview for auto-created printers only"; Break}
							"GenericOnly"           {$tmp = "Use print preview for generic universal printers only"; Break}
							"AutoCreatedAndGeneric" {$tmp = "Use print preview for both auto-created and generic universal printers"; Break}
							Default {$tmp = "Universal printing preview preference could not be determined: $($Setting.UniversalPrintingPreviewPreference.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting DPILimit State ) -and ($Setting.DPILimit.State -ne "NotConfigured"))
					{
						$txt = "ICA\Printing\Universal Printing\Universal printing print quality limit"
						$tmp = ""
						Switch ($Setting.DPILimit.Value)
						{
							"Draft"				{$tmp = "Draft (150 DPI)"; Break}
							"LowResolution"		{$tmp = "Low Resolution (300 DPI)"; Break}
							"MediumResolution"	{$tmp = "Medium Resolution (600 DPI)"; Break}
							"HighResolution"	{$tmp = "High Resolution (1200 DPI)"; Break}
							"Unlimited"			{$tmp = "No Limit"; Break}
							Default {$tmp = "Universal printing print quality limit could not be determined: $($Setting.DPILimit.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Host "$(Get-Date): `t`t`tICA\Security" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting MinimumEncryptionLevel State ) -and ($Setting.MinimumEncryptionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Security\SecureICA minimum encryption level" 
						$tmp = ""
						Switch ($Setting.MinimumEncryptionLevel.Value)
						{
							"Unknown"	{$tmp = "Unknown encryption"; Break}
							"Basic"		{$tmp = "Basic"; Break}
							"LogOn"		{$tmp = "RC5 (128 bit) logon only"; Break}
							"Bits40"	{$tmp = "RC5 (40 bit)"; Break}
							"Bits56"	{$tmp = "RC5 (56 bit)"; Break}
							"Bits128"	{$tmp = "RC5 (128 bit)"; Break}
							Default		{$tmp = "SecureICA minimum encryption level could not be determined: $($Setting.MinimumEncryptionLevel.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Server Limits" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting IdleTimerInterval State ) -and ($Setting.IdleTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Server Limits\Server idle timer interval (milliseconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.IdleTimerInterval.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.IdleTimerInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.IdleTimerInterval.Value 
						}
					}
					
					Write-Host "$(Get-Date): `t`t`tICA\Session Limits" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting SessionDisconnectTimer State ) -and ($Setting.SessionDisconnectTimer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Disconnected session timer"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionDisconnectTimer.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionDisconnectTimer.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionDisconnectTimer.State 
						}
					}
					If((validStateProp $Setting SessionDisconnectTimerInterval State ) -and ($Setting.SessionDisconnectTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Disconnected session timer interval (minutes)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionDisconnectTimerInterval.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionDisconnectTimerInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionDisconnectTimerInterval.Value 
						}
					}
					If((validStateProp $Setting SessionConnectionTimer State ) -and ($Setting.SessionConnectionTimer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session connection timer"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionConnectionTimer.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionConnectionTimer.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionConnectionTimer.State 
						}
					}
					If((validStateProp $Setting SessionConnectionTimerInterval State ) -and ($Setting.SessionConnectionTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session connection timer interval (minutes)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionConnectionTimerInterval.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionConnectionTimerInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionConnectionTimerInterval.Value 
						}
					}
					If((validStateProp $Setting SessionIdleTimer State ) -and ($Setting.SessionIdleTimer.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session idle timer"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionIdleTimer.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionIdleTimer.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionIdleTimer.State 
						}
					}
					If((validStateProp $Setting SessionIdleTimerInterval State ) -and ($Setting.SessionIdleTimerInterval.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Limits\Session idle timer interval (minutes)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionIdleTimerInterval.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionIdleTimerInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionIdleTimerInterval.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Session Reliability" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting SessionReliabilityConnections State ) -and ($Setting.SessionReliabilityConnections.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Reliability\Session reliability connections"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionReliabilityConnections.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionReliabilityConnections.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionReliabilityConnections.State 
						}
					}
					If((validStateProp $Setting SessionReliabilityPort State ) -and ($Setting.SessionReliabilityPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Reliability\Session reliability port number"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionReliabilityPort.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionReliabilityPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionReliabilityPort.Value 
						}
					}
					If((validStateProp $Setting SessionReliabilityTimeout State ) -and ($Setting.SessionReliabilityTimeout.State -ne "NotConfigured"))
					{
						$txt = "ICA\Session Reliability\Session reliability timeout (seconds)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.SessionReliabilityTimeout.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SessionReliabilityTimeout.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SessionReliabilityTimeout.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Time Zone Control" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting LocalTimeEstimation State ) -and ($Setting.LocalTimeEstimation.State -ne "NotConfigured"))
					{
						$txt = "ICA\Time Zone Control\Estimate local time for legacy clients"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LocalTimeEstimation.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LocalTimeEstimation.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LocalTimeEstimation.State 
						}
					}
					If((validStateProp $Setting SessionTimeZone State ) -and ($Setting.SessionTimeZone.State -ne "NotConfigured"))
					{
						$txt = "ICA\Time Zone Control\Use local time of client"
						$tmp = ""
						Switch ($Setting.SessionTimeZone.Value)
						{
							"UseServerTimeZone" {$tmp = "Use server time zone"; Break}
							"UseClientTimeZone" {$tmp = "Use client time zone"; Break}
							Default {$tmp = "Use local time of client could not be determined: $($Setting.SessionTimeZone.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Host "$(Get-Date): `t`t`tICA\TWAIN Devices" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting TwainRedirection State ) -and ($Setting.TwainRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\TWAIN devices\Client TWAIN device redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TwainRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TwainRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TwainRedirection.State 
						}
					}
					If((validStateProp $Setting TwainCompressionLevel State ) -and ($Setting.TwainCompressionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\TWAIN devices\TWAIN compression level"
						Switch ($Setting.TwainCompressionLevel.Value)
						{
							"None"		{$tmp = "None"; Break}
							"Low"		{$tmp = "Low"; Break}
							"Medium"	{$tmp = "Medium"; Break}
							"High"		{$tmp = "High"; Break}
							Default		{$tmp = "TWAIN compression level could not be determined: $($Setting.TwainCompressionLevel.Value)"; Break}
						}

						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Host "$(Get-Date): `t`t`tICA\USB Devices" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting UsbDeviceRedirection State ) -and ($Setting.UsbDeviceRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB device redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UsbDeviceRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UsbDeviceRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UsbDeviceRedirection.State 
						}
					}
					If((validStateProp $Setting UsbDeviceRedirectionRules State ) -and ($Setting.UsbDeviceRedirectionRules.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB device redirection rules"
						If(validStateProp $Setting UsbDeviceRedirectionRules Values )
						{
							$array = $Setting.UsbDeviceRedirectionRules.Values
							$tmp = $array[0]
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp 
							}

							$txt = ""
							$cnt = -1
							ForEach($element in $array)
							{
								$cnt++
								
								If($cnt -ne 0)
								{
									$tmp = "$($element) "
									If($MSWord -or $PDF)
									{
										$SettingsWordTable += @{
										Text = "";
										Value = $tmp;
										}
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$array = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Client USB device redirections rules were found"
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = $txt;
								Value = $tmp;
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp 
							}
						}
					}
					If((validStateProp $Setting UsbPlugAndPlayRedirection State ) -and ($Setting.UsbPlugAndPlayRedirection.State -ne "NotConfigured"))
					{
						$txt = "ICA\USB devices\Client USB Plug and Play device redirection"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UsbPlugAndPlayRedirection.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UsbPlugAndPlayRedirection.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UsbPlugAndPlayRedirection.State 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Visual Display" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting PreferredColorDepthForSimpleGraphics State ) -and ($Setting.PreferredColorDepthForSimpleGraphics.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Preferred color depth for simple graphics"
						$tmp = ""
						Switch ($Setting.PreferredColorDepthForSimpleGraphics.Value)
						{
							"ColorDepth24Bit"	{$tmp = "24 bits per pixel"; Break}
							"ColorDepth16Bit"	{$tmp = "16 bits per pixel"; Break}
							"ColorDepth8Bit"	{$tmp = "8 bits per pixel"; Break}
							"Default" {$tmp = "Preferred color depth for simple graphics could not be determined: $($Setting.PreferredColorDepthForSimpleGraphics.Value)"; Break}
						}
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
					}
					If((validStateProp $Setting FramesPerSecond State ) -and ($Setting.FramesPerSecond.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Target frame rate (fps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FramesPerSecond.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FramesPerSecond.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FramesPerSecond.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Visual Display\Moving Images" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting MinimumAdaptiveDisplayJpegQuality State ) -and ($Setting.MinimumAdaptiveDisplayJpegQuality.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Minimum image quality"
						$tmp = ""
						Switch ($Setting.MinimumAdaptiveDisplayJpegQuality.Value)
						{
							"UltraHigh" {$tmp = "Ultra high"; Break}
							"VeryHigh"  {$tmp = "Very high"; Break}
							"High"      {$tmp = "High"; Break}
							"Normal"    {$tmp = "Normal"; Break}
							"Low"       {$tmp = "Low"; Break}
							Default {$tmp = "Minimum image quality could not be determined: $($Setting.MinimumAdaptiveDisplayJpegQuality.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting MovingImageCompressionConfiguration State ) -and ($Setting.MovingImageCompressionConfiguration.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Moving image compression"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MovingImageCompressionConfiguration.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MovingImageCompressionConfiguration.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MovingImageCompressionConfiguration.State 
						}
					}
					If((validStateProp $Setting ProgressiveCompressionLevel State ) -and ($Setting.ProgressiveCompressionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Progressive compression level"
						$tmp = ""
						Switch ($Setting.ProgressiveCompressionLevel.Value)
						{
							"UltraHigh" {$tmp = "Ultra high"; Break}
							"VeryHigh"  {$tmp = "Very high"; Break}
							"High"      {$tmp = "High"; Break}
							"Normal"    {$tmp = "Normal"; Break}
							"Low"       {$tmp = "Low"; Break}
							"None"      {$tmp = "None"; Break}
							Default {$tmp = "Progressive compression level could not be determined: $($Setting.ProgressiveCompressionLevel.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting ProgressiveCompressionThreshold State ) -and ($Setting.ProgressiveCompressionThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Progressive compression threshold value (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ProgressiveCompressionThreshold.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProgressiveCompressionThreshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProgressiveCompressionThreshold.Value 
						}
					}
					If((validStateProp $Setting TargetedMinimumFramesPerSecond State ) -and ($Setting.TargetedMinimumFramesPerSecond.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Moving Images\Target Minimum Frame Rate (fps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.TargetedMinimumFramesPerSecond.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.TargetedMinimumFramesPerSecond.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.TargetedMinimumFramesPerSecond.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\Visual Display\Still Images" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting ExtraColorCompression State ) -and ($Setting.ExtraColorCompression.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Extra Color Compression"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExtraColorCompression.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExtraColorCompression.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExtraColorCompression.State 
						}
					}
					If((validStateProp $Setting ExtraColorCompressionThreshold State ) -and ($Setting.ExtraColorCompressionThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Extra Color Compression Threshold (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ExtraColorCompressionThreshold.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ExtraColorCompressionThreshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ExtraColorCompressionThreshold.Value 
						}
					}
					If((validStateProp $Setting ProgressiveHeavyweightCompression State ) -and ($Setting.ProgressiveHeavyweightCompression.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Heavyweight compression"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ProgressiveHeavyweightCompression.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProgressiveHeavyweightCompression.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProgressiveHeavyweightCompression.State 
						}
					}
					If((validStateProp $Setting LossyCompressionLevel State ) -and ($Setting.LossyCompressionLevel.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Lossy compression level"
						$tmp = ""
						Switch ($Setting.LossyCompressionLevel.Value)
						{
							"None"		{$tmp = "None"; Break}
							"Low"		{$tmp = "Low"; Break}
							"Medium"	{$tmp = "Medium"; Break}
							"High"		{$tmp = "High"; Break}
							Default		{$tmp = "Lossy compression level could not be determined: $($Setting.LossyCompressionLevel.Value)"; Break}
						}
						
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If((validStateProp $Setting LossyCompressionThreshold State ) -and ($Setting.LossyCompressionThreshold.State -ne "NotConfigured"))
					{
						$txt = "ICA\Visual Display\Still Images\Lossy compression threshold value (Kbps)"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LossyCompressionThreshold.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LossyCompressionThreshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LossyCompressionThreshold.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tICA\WebSockets" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting AcceptWebSocketsConnections State ) -and ($Setting.AcceptWebSocketsConnections.State -ne "NotConfigured"))
					{
						$txt = "ICA\WebSockets\WebSocket connections"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.AcceptWebSocketsConnections.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.AcceptWebSocketsConnections.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.AcceptWebSocketsConnections.State 
						}
					}
					If((validStateProp $Setting WebSocketsPort State ) -and ($Setting.WebSocketsPort.State -ne "NotConfigured"))
					{
						$txt = "ICA\WebSockets\WebSockets port number"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.WebSocketsPort.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.WebSocketsPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.WebSocketsPort.Value 
						}
					}
					If((validStateProp $Setting WSTrustedOriginServerList State ) -and ($Setting.WSTrustedOriginServerList.State -ne "NotConfigured"))
					{
						$txt = "ICA\WebSockets\WebSockets trusted origin server list"
						$tmpArray = $Setting.WSTrustedOriginServerList.Value.Split(",")
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $tmpArray)
						{
							$cnt++
							$tmp = "$($Thing)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$tmpArray = $Null
						$tmp = $Null
					}
					
					Write-Host "$(Get-Date): `t`t`tServer Settings" -BackgroundColor Black -ForegroundColor Yellow
					If($Setting.ConnectionAccessControl.State -ne "NotConfigured")
					{
						Switch ($Setting.ConnectionAccessControl.Value)
						{
							"AllowAny"                     {$tmp = "Any connections"}
							"AllowTicketedConnectionsOnly" {$tmp = "Citrix Access Gateway, Citrix Receiver, and Web Interface connections only"}
							"AllowAccessGatewayOnly"       {$tmp = "Citrix Access Gateway connections only"}
							Default {$tmp = "Connection access control could not be determined: $($Setting.ConnectionAccessControl.Value)"}
						}
						$txt = "Server Settings\Connection access control"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tp 
						}
						$tmp = $Null
					}
					If($Setting.DnsAddressResolution.State -ne "NotConfigured")
					{
						$txt = "Server Settings\DNS address resolution"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.DnsAddressResolution.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.DnsAddressResolution.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.DnsAddressResolution.State 
						}
					}
					If($Setting.FullIconCaching.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Full icon caching"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.FullIconCaching.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.FullIconCaching.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.FullIconCaching.State
						}
					}
					#the next setting is only available for AD based policies
					If($Setting.InitialZone.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Initial Zone Name"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.InitialZone.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.InitialZone.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.InitialZone.State 
						}
					}
					If($Setting.LoadEvaluator.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Load Evaluator Name - Load evaluator"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.LoadEvaluator.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.LoadEvaluator.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.LoadEvaluator.Value 
						}
					}
					If($Setting.ProductEdition.State -ne "NotConfigured")
					{
						$txt = "Server Settings\XenApp product edition"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ProductEdition.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProductEdition.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProductEdition.Value 
						}
					}
					If($Setting.ProductModel.State -ne "NotConfigured")
					{
						Switch ($Setting.ProductModel.Value)
						{
							"XenAppCCU"                  {$tmp = "XenApp"}
							"XenDesktopConcurrentServer" {$tmp = "XenDesktop Concurrent"}
							"XenDesktopUserDevice"       {$tmp = "XenDesktop User Device"}
							Default {$tmp = "XenApp product model could not be determined: $($Setting.ProductModel.Value)"}
						}
						$txt = "Server Settings\XenApp product model"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tp 
						}
						$tmp = $Null
					}
					If($Setting.UserSessionLimit.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Connection Limits\Limit user sessions"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UserSessionLimit.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UserSessionLimit.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UserSessionLimit.Value 
						}
					}
					If($Setting.UserSessionLimitAffectsAdministrators.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Connection Limits\Limits on administrator sessions"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UserSessionLimitAffectsAdministrators.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UserSessionLimitAffectsAdministrators.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UserSessionLimitAffectsAdministrators.State 
						}
					}
					If($Setting.UserSessionLimitLogging.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Connection Limits\Logging of logon limit events"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.UserSessionLimitLogging.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.UserSessionLimitLogging.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.UserSessionLimitLogging.State 
						}
					}
					#the next 3 settings are available only for AD based policies
					If($Setting.InitialDatabaseName.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Database Settings\Initial Database Name"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.InitialDatabaseName.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.InitialDatabaseName.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.InitialDatabaseName.Value 
						}
					}
					If($Setting.InitialDatabaseServerName.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Database Settings\Initial Database Server Name"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.InitialDatabaseServerName.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.InitialDatabaseServerName.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.InitialDatabaseServerName.Value 
						}
					}
					If($Setting.InitialFailoverPartner.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Database Settings\Initial Failover Partner"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.InitialFailoverPartner.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.InitialFailoverPartner.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.InitialFailoverPartner.Value 
						}
					}
					#the previous 3 settings are available only for AD based policies
					If($Setting.HealthMonitoring.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Health Monitoring and Recovery\Health monitoring"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.HealthMonitoring.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.HealthMonitoring.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.HealthMonitoring.State
						}
					}
					If($Setting.HealthMonitoringTests.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Health Monitoring and Recovery\Health monitoring tests"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = "";
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							"",$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt "" 
						}
						[xml]$XML = $Setting.HealthMonitoringTests.Value
						ForEach($Test in $xml.hmrtests.tests.test)
						{
							Switch ($test.RecoveryAction)
							{
								"AlertOnly"                     {$tmp = "Alert Only"}
								"RemoveServerFromLoadBalancing" {$tmp = "Remove Server from load balancing"}
								"RestartIma"                    {$tmp = "Restart IMA"}
								"ShutdownIma"                   {$tmp = "Shutdown IMA"}
								"RebootServer"                  {$tmp = "Reboot Server"}
								Default {$tmp = "Recovery Action could not be determined: $($test.RecoveryAction)"}
							}
							$tmparray = @()
							$tmparray += "Name: $($test.name)"
							$tmparray += "File Location: $($test.file)"
							$tmparray += "Arguments: $($test.arguments)"
							$tmparray += "Description: $($test.description)"
							$tmparray += "Interval: $($test.interval)"
							$tmparray += "Time-out: $($test.timeout)"
							$tmparray += "Threshold: $($test.threshold)"
							$tmparray += "Recovery Action: $($tmp)"
							ForEach($item in $tmparray)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $item;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$item,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "`t`t`t`t`t`t`t`t`t" $item
								}
							}
							#insert a blank line for spacing
							If($MSWord -or $PDF)
							{
								$SettingsWordTable += @{
								Text = "";
								Value = "";
								}
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								"",$htmlbold,
								"",$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting "" ""
							}
						}
						$XML = $Null
					}

					If($Setting.MaximumServersOfflinePercent.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Health Monitoring and Recovery\Max % of servers with logon control"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MaximumServersOfflinePercent.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MaximumServersOfflinePercent.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MaximumServersOfflinePercent.Value 
						}
					}
					If($Setting.CpuManagementServerLevel.State -ne "NotConfigured")
					{
						Switch ($Setting.CpuManagementServerLevel.Value)
						{
							"NoManagement" {$tmp = "No CPU utilization management"}
							"Fair"         {$tmp = "Fair sharing of CPU between sessions"}
							"Preferential" {$tmp = "Preferential Load Balancing"}
							Default {$tmp = "CPU management server level could not be determined: $($Setting.CpuManagementServerLevel.Value)"}
						}
						$txt = "Server Settings\Memory/CPU\CPU management server level"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If($Setting.MemoryOptimization.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Memory/CPU\Memory optimization"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MemoryOptimization.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MemoryOptimization.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MemoryOptimization.State 
						}
					}
					If($Setting.MemoryOptimizationExcludedPrograms.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Memory/CPU\Memory optimization application exclusion lis"
						$tmpArray = $Setting.MemoryOptimizationExcludedPrograms.Values
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $tmpArray)
						{
							$cnt++
							$tmp = "$($Thing)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$tmpArray = $Null
						$tmp = $Null
					}
					If($Setting.MemoryOptimizationIntervalType.State -ne "NotConfigured")
					{
						Switch ($Setting.MemoryOptimizationIntervalType.Value)
						{
							"AtStartup" {$tmp = "Only at startup time"}
							"Daily"     {$tmp = "Daily"}
							"Weekly"    {$tmp = "Weekly"}
							"Monthly"   {$tmp = "Monthly"}
							Default {$tmp = " could not be determined: $($Setting.MemoryOptimizationIntervalType.Value)"}
						}
						$txt = "Server Settings\Memory/CPU\Memory optimization interval"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If($Setting.MemoryOptimizationDayOfMonth.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Memory/CPU\Memory optimization schedule\day of month"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MemoryOptimizationDayOfMonth.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MemoryOptimizationDayOfMonth.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MemoryOptimizationDayOfMonth.Value 
						}
					}
					If($Setting.MemoryOptimizationDayOfWeek.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Memory/CPU\Memory optimization schedule\day of week"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.MemoryOptimizationDayOfWeek.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.MemoryOptimizationDayOfWeek.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.MemoryOptimizationDayOfWeek.Value 
						}
					}
					If($Setting.MemoryOptimizationTime.State -ne "NotConfigured")
					{
						$tmp = ConvertNumberToTime $Setting.MemoryOptimizationTime.Value
						$txt = "Server Settings\Memory/CPU\Memory optimization schedule time"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If($Setting.OfflineClientTrust.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Offline Applications\Offline app client trust"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.OfflineClientTrust.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OfflineClientTrust.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.OfflineClientTrust.State 
						}
					}
					If($Setting.OfflineEventLogging.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Offline Applications\Offline app event logging"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.OfflineEventLogging.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OfflineEventLogging.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.OfflineEventLogging.State 
						}
					}
					If($Setting.OfflineLicensePeriod.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Offline Applications\Offline app license period - Days"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.OfflineLicensePeriod.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.OfflineLicensePeriod.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.OfflineLicensePeriod.Value 
						}
					}
					If($Setting.OfflineUsers.State -ne "NotConfigured")
					{
						$array = $Null
						$txt = "Server Settings\Offline Applications\Offline app users"
						$tmpArray = $Setting.OfflineUsers.Values
						$tmp = ""
						$cnt = 0
						ForEach($Thing in $tmpArray)
						{
							$cnt++
							$tmp = "$($Thing)"
							If($cnt -eq 1)
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = $txt;
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									$txt,$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting $txt $tmp
								}
							}
							Else
							{
								If($MSWord -or $PDF)
								{
									$SettingsWordTable += @{
									Text = "";
									Value = $tmp;
									}
								}
								ElseIf($HTML)
								{
									$rowdata += @(,(
									"",$htmlbold,
									$tmp,$htmlwhite))
								}
								ElseIf($Text)
								{
									OutputPolicySetting "" $tmp
								}
							}
						}
						$tmpArray = $Null
						$tmp = $Null
					}
					If($Setting.RebootCustomMessage.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Reboot Behavior\Reboot custom warning"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RebootCustomMessage.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RebootCustomMessage.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RebootCustomMessage.State 
						}
					}
					If($Setting.RebootCustomMessageText.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Reboot Behavior\Reboot custom warning text"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RebootCustomMessageText.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RebootCustomMessageText.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RebootCustomMessageText.Value 
						}
					}
					If($Setting.RebootDisableLogOnTime.State -ne "NotConfigured")
					{
						Switch ($Setting.RebootDisableLogOnTime.Value)
						{
							"DoNotDisableLogOnsBeforeReboot" {$tmp = "Do not disable logons before reboot"}
							"Disable5MinutesBeforeReboot"    {$tmp = "Disable 5 minutes before reboot"}
							"Disable10MinutesBeforeReboot"   {$tmp = "Disable 10 minutes before reboot"}
							"Disable15MinutesBeforeReboot"   {$tmp = "Disable 15 minutes before reboot"}
							"Disable30MinutesBeforeReboot"   {$tmp = "Disable 30 minutes before reboot"}
							"Disable60MinutesBeforeReboot"   {$tmp = "Disable 60 minutes before reboot"}
							Default {$tmp = "Reboot logon disable time could not be determined: $($Setting.RebootDisableLogOnTime.Value)"}
						}
						$txt = "Server Settings\Reboot Behavior\Reboot logon disable time"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}
					If($Setting.RebootScheduleFrequency.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Reboot Behavior\Reboot schedule frequency - Days"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RebootScheduleFrequency.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RebootScheduleFrequency.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RebootScheduleFrequency.Value 
						}
					}
					If($Setting.RebootScheduleRandomizationInterval.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Reboot Behavior\Reboot schedule randomization interval\Minutes"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RebootScheduleRandomizationInterval.Value;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RebootScheduleRandomizationInterval.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RebootScheduleRandomizationInterval.Value 
						}
					}
					If($Setting.RebootScheduleStartDate.State -ne "NotConfigured")
					{
						$Tmp = ConvertIntegerToDate $Setting.RebootScheduleStartDate.Value
						$txt = "Server Settings\Reboot Behavior\Reboot schedule start date"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$Tmp = $Null
					}
					If($Setting.RebootScheduleTime.State -ne "NotConfigured")
					{
						$tmp = ConvertNumberToTime $Setting.RebootScheduleTime.Value 						
						$txt = "Server Settings\Reboot Behavior\Reboot schedule time"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$Tmp = $Null
					}
					If($Setting.RebootWarningInterval.State -ne "NotConfigured")
					{
						Switch ($Setting.RebootWarningInterval.Value)
						{
							"Every1Minute"   {$tmp = "Every 1 Minute"}
							"Every3Minutes"  {$tmp = "Every 3 Minutes"}
							"Every5Minutes"  {$tmp = "Every 5 Minutes"}
							"Every10Minutes" {$tmp = "Every 10 Minutes"}
							"Every15Minutes" {$tmp = "Every 15 Minutes"}
							Default {$tmp = "Reboot warning interval could not be determined: $($Setting.RebootWarningInterval.Value)"}
						}
						$txt = "Server Settings\Reboot Behavior\Reboot warning interval"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$Tmp = $Null
					}
					If($Setting.RebootWarningStartTime.State -ne "NotConfigured")
					{
						Switch ($Setting.RebootWarningStartTime.Value)
						{
							"Start5MinutesBeforeReboot"  {$tmp = "Start 5 Minutes Before Reboot"}
							"Start10MinutesBeforeReboot" {$tmp = "Start 10 Minutes Before Reboot"}
							"Start15MinutesBeforeReboot" {$tmp = "Start 15 Minutes Before Reboot"}
							"Start30MinutesBeforeReboot" {$tmp = "Start 30 Minutes Before Reboot"}
							"Start60MinutesBeforeReboot" {$tmp = "Start 60 Minutes Before Reboot"}
							Default {$tmp = "Reboot warning start time could not be determined: $($Setting.RebootWarningStartTime.Value)"}
						}
						$txt = "Server Settings\Reboot Behavior\Reboot warning start time"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $tmp;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$Tmp = $Null
					}
					If($Setting.RebootWarningMessage.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Reboot Behavior\Reboot warning to users"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.RebootWarningMessage.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.RebootWarningMessage.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.RebootWarningMessage.State 
						}
					}
					If($Setting.ScheduledReboots.State -ne "NotConfigured")
					{
						$txt = "Server Settings\Reboot Behavior\Scheduled reboots"
						If($MSWord -or $PDF)
						{
							$SettingsWordTable += @{
							Text = $txt;
							Value = $Setting.ScheduledReboots.State;
							}
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ScheduledReboots.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ScheduledReboots.State 
						}
					}

					Write-Host "$(Get-Date): `t`t`tVirtual Delivery Agent Settings" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting ControllerRegistrationPort State ) -and ($Setting.ControllerRegistrationPort.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controller registration port"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ControllerRegistrationPort.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ControllerRegistrationPort.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ControllerRegistrationPort.Value 
						}
					}
					If((validStateProp $Setting ControllerSIDs State ) -and ($Setting.ControllerSIDs.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controller SIDs"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ControllerSIDs.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ControllerSIDs.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ControllerSIDs.Value 
						}
					}
					If((validStateProp $Setting Controllers State ) -and ($Setting.Controllers.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Controllers"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.Controllers.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.Controllers.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.Controllers.Value 
						}
					}
					
					Write-Host "$(Get-Date): `t`t`tVirtual Delivery Agent Settings\CPU Usage Monitoring" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting CPUUsageMonitoring_Enable State ) -and ($Setting.CPUUsageMonitoring_Enable.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\CPU Usage Monitoring\Enable Monitoring"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.CPUUsageMonitoring_Enable.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.CPUUsageMonitoring_Enable.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.CPUUsageMonitoring_Enable.State 
						}
					}
					If((validStateProp $Setting CPUUsageMonitoring_Period State ) -and ($Setting.CPUUsageMonitoring_Period.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\CPU Usage Monitoring\Monitoring Period (seconds)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.CPUUsageMonitoring_Period.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.CPUUsageMonitoring_Period.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.CPUUsageMonitoring_Period.Value 
						}
					}
					If((validStateProp $Setting CPUUsageMonitoring_Threshold State ) -and ($Setting.CPUUsageMonitoring_Threshold.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\CPU Usage Monitoring\Threshold (percent)"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.CPUUsageMonitoring_Threshold.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.CPUUsageMonitoring_Threshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.CPUUsageMonitoring_Threshold.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tVirtual Delivery Agent Settings\HDX3DPro" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting EnableLossless State ) -and ($Setting.EnableLossless.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\HDX3DPro\Enable lossless"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.EnableLossless.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.EnableLossless.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.EnableLossless.State 
						}
					}
					If((validStateProp $Setting ProGraphicsObj State ) -and ($Setting.ProGraphicsObj.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\HDX3DPro\HDX3DPro quality settings"
						$tmp = ""
						$xMin = [math]::floor($Setting.ProGraphicsObj.Value%65536).ToString()
						$xMax = [math]::floor($Setting.ProGraphicsObj.Value/65536).ToString()
						[string]$tmp = "Minimum: $($xMin) Maximum: $($xMax)"
						
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $tmp;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$tmp,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $tmp 
						}
						$tmp = $Null
					}

					Write-Host "$(Get-Date): `t`t`tVirtual Delivery Agent Settings\ICA Latency Monitoring" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting ICALatencyMonitoring_Enable State ) -and ($Setting.ICALatencyMonitoring_Enable.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\ICA Latency Monitoring\Enable Monitoring"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ICALatencyMonitoring_Enable.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ICALatencyMonitoring_Enable.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ICALatencyMonitoring_Enable.State 
						}
					}
					If((validStateProp $Setting ICALatencyMonitoring_Period State ) -and ($Setting.ICALatencyMonitoring_Period.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\ICA Latency Monitoring\Monitoring Period seconds"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ICALatencyMonitoring_Period.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ICALatencyMonitoring_Period.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ICALatencyMonitoring_Period.Value 
						}
					}
					If((validStateProp $Setting ICALatencyMonitoring_Threshold State ) -and ($Setting.ICALatencyMonitoring_Threshold.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\ICA Latency Monitoring\Threshold milliseconds"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ICALatencyMonitoring_Threshold.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ICALatencyMonitoring_Threshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ICALatencyMonitoring_Threshold.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tVirtual Delivery Agent Settings\Monitoring" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting SiteGUID State ) -and ($Setting.SiteGUID.State -ne "NotConfigured"))
					{
						#AD specific setting
						$txt = "Virtual Delivery Agent Settings\Site GUID"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.SiteGUID.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.SiteGUID.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.SiteGUID.Value 
						}
					}
					
					Write-Host "$(Get-Date): `t`t`tVirtual Delivery Agent Settings\Profile Load Time Monitoring" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting ProfileLoadTimeMonitoring_Enable State ) -and ($Setting.ProfileLoadTimeMonitoring_Enable.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Profile Load Time Monitoring\Enable Monitoring"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ProfileLoadTimeMonitoring_Enable.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProfileLoadTimeMonitoring_Enable.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProfileLoadTimeMonitoring_Enable.State 
						}
					}
					If((validStateProp $Setting ProfileLoadTimeMonitoring_Threshold State ) -and ($Setting.ProfileLoadTimeMonitoring_Threshold.State -ne "NotConfigured"))
					{
						$txt = "Virtual Delivery Agent Settings\Profile Load Time Monitoring\Threshold seconds"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.ProfileLoadTimeMonitoring_Threshold.Value;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.ProfileLoadTimeMonitoring_Threshold.Value,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.ProfileLoadTimeMonitoring_Threshold.Value 
						}
					}

					Write-Host "$(Get-Date): `t`t`tVirtual IP" -BackgroundColor Black -ForegroundColor Yellow
					If((validStateProp $Setting VirtualLoopbackSupport State ) -and ($Setting.VirtualLoopbackSupport.State -ne "NotConfigured"))
					{
						$txt = "Virtual IP\Virtual IP loopback support"
						If($MSWord -or $PDF)
						{
							$WordTableRowHash = @{
							Text = $txt;
							Value = $Setting.VirtualLoopbackSupport.State;
							}
							$SettingsWordTable += $WordTableRowHash;
						}
						ElseIf($HTML)
						{
							$rowdata += @(,(
							$txt,$htmlbold,
							$Setting.VirtualLoopbackSupport.State,$htmlwhite))
						}
						ElseIf($Text)
						{
							OutputPolicySetting $txt $Setting.VirtualLoopbackSupport.State 
						}
					}
					If((validStateProp $Setting VirtualLoopbackPrograms State ) -and ($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured"))
					{
						$txt = "Virtual IP\Virtual IP virtual loopback programs list"
						If((validStateProp $Setting VirtualLoopbackPrograms State ) -and ($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured"))
						{
							$tmpArray = $Setting.VirtualLoopbackPrograms.Values
							$array = $Null
							$tmp = ""
							$cnt = 0
							ForEach($Thing in $TmpArray)
							{
								If($Null -eq $Thing)
								{
									$Thing = ''
								}
								$cnt++
								$tmp = "$($Thing) "
								If($cnt -eq 1)
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = $txt;
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										$txt,$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting $txt $tmp
									}
								}
								Else
								{
									If($MSWord -or $PDF)
									{
										$WordTableRowHash = @{
										Text = "";
										Value = $tmp;
										}
										$SettingsWordTable += $WordTableRowHash;
									}
									ElseIf($HTML)
									{
										$rowdata += @(,(
										"",$htmlbold,
										$tmp,$htmlwhite))
									}
									ElseIf($Text)
									{
										OutputPolicySetting "" $tmp
									}
								}
							}
							$TmpArray = $Null
							$tmp = $Null
						}
						Else
						{
							$tmp = "No Virtual IP virtual loopback programs list were found"
							If($MSWord -or $PDF)
							{
								$WordTableRowHash = @{
								Text = $txt;
								Value = $tmp;
								}
								$SettingsWordTable += $WordTableRowHash;
							}
							ElseIf($HTML)
							{
								$rowdata += @(,(
								$txt,$htmlbold,
								$tmp,$htmlwhite))
							}
							ElseIf($Text)
							{
								OutputPolicySetting $txt $tmp
							}
						}
					}
				}
				If($MSWord -or $PDF)
				{
					If($SettingsWordTable.Count -gt 0)
					{
						$Table = AddWordTable -Hashtable $SettingsWordTable `
						-Columns  Text,Value `
						-Headers  "Setting Key","Value"`
						-Format $wdTableLightListAccent3 `
						-NoInternalGridLines `
						-AutoFit $wdAutoFitFixed;

						SetWordCellFormat -Collection $Table -Size 9
						
						SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

						$Table.Columns.Item(1).Width = 300;
						$Table.Columns.Item(2).Width = 200;

						$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)
					}
					Else
					{
						WriteWordLine 0 1 "There are no policy settings"
					}
					FindWordDocumentEnd
					$Table = $Null
				}
				ElseIf($Text)
				{
					Line 0 ""
				}
				ElseIf($HTML)
				{
					If($rowdata.count -gt 0)
					{
						$columnHeaders = @(
						'Setting Key',($htmlsilver -bor $htmlbold),
						'Value',($htmlsilver -bor $htmlbold))

						$msg = ""
						$columnWidths = @("400","300")
						FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths -tablewidth "700"
						WriteHTMLLine 0 0 " "
					}
				}
			}
			Else
			{
				$txt = "Unable to retrieve settings"
				If($MSWord -or $PDF)
				{
					WriteWordLine 0 1 $txt
				}
				ElseIf($Text)
				{
					Line 2 $txt
				}
				ElseIf($HTML)
				{
					WriteHTMLLine 0 1 $txt
				}
			}
			$Filter = $Null
			$Settings = $Null
			Write-Host "$(Get-Date): `t`tFinished $($Policy.PolicyName)" -BackgroundColor Black -ForegroundColor Yellow
			Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
		}
	}
	ElseIf(!$?)
	{
		Write-Warning "Citrix Policy information could not be retrieved"
	}
	Else
	{
		Write-Warning "No results Returned for Citrix Policy information"
	}
	
	$HDXPolicies = $Null
	If($xDriveName -ne "localfarmgpo")
	{
		Write-Host "$(Get-Date): `tRemoving $($xDriveName) PSDrive" -BackgroundColor Black -ForegroundColor Yellow
		Remove-PSDrive $xDriveName -EA 0 4>Null
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputPolicySetting
{
	Param([string] $outputText, [string] $outputData)

	If($outputText -ne "")
	{
		$xLength = $outputText.Length
		If($outputText.Substring($xLength-2,2) -ne ": ")
		{
			$outputText += ": "
		}
	}
	Line 2 $outputText $outputData
}

Function Get-PrinterModifiedSettings
{
	Param([string]$Value, [string]$xelement)
	
	[string]$ReturnStr = ""

	Switch ($Value)
	{
		"copi" 
		{
			$txt="Copies: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"coll"
		{
			$txt="Collate: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"scal"
		{
			$txt="Scale (%): "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"colo"
		{
			$txt="Color: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Monochrome"; Break}
					2 {$tmp2 = "Color"; Break}
					Default {$tmp2 = "Color could not be determined: $($xelement) "; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"prin"
		{
			$txt="Print Quality: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					-1 {$tmp2 = "150 dpi"; Break}
					-2 {$tmp2 = "300 dpi"; Break}
					-3 {$tmp2 = "600 dpi"; Break}
					-4 {$tmp2 = "1200 dpi"; Break}
					Default {$tmp2 = "Custom...X resolution: $tmp1"; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"yres"
		{
			$txt="Y resolution: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"orie"
		{
			$txt="Orientation: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					"portrait"  {$tmp2 = "Portrait"; Break}
					"landscape" {$tmp2 = "Landscape"; Break}
					Default {$tmp2 = "Orientation could not be determined: $($xelement) ; Break"}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"dupl"
		{
			$txt="Duplex: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Simplex"; Break}
					2 {$tmp2 = "Vertical"; Break}
					3 {$tmp2 = "Horizontal"; Break}
					Default {$tmp2 = "Duplex could not be determined: $($xelement) "; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"pape"
		{
			$txt="Paper Size: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1   {$tmp2 = "Letter"; Break}
					2   {$tmp2 = "Letter Small"; Break}
					3   {$tmp2 = "Tabloid"; Break}
					4   {$tmp2 = "Ledger"; Break}
					5   {$tmp2 = "Legal"; Break}
					6   {$tmp2 = "Statement"; Break}
					7   {$tmp2 = "Executive"; Break}
					8   {$tmp2 = "A3"; Break}
					9   {$tmp2 = "A4"; Break}
					10  {$tmp2 = "A4 Small"; Break}
					11  {$tmp2 = "A5"; Break}
					12  {$tmp2 = "B4 (JIS)"; Break}
					13  {$tmp2 = "B5 (JIS)"; Break}
					14  {$tmp2 = "Folio"; Break}
					15  {$tmp2 = "Quarto"; Break}
					16  {$tmp2 = "10X14"; Break}
					17  {$tmp2 = "11X17"; Break}
					18  {$tmp2 = "Note"; Break}
					19  {$tmp2 = "Envelope #9"; Break}
					20  {$tmp2 = "Envelope #10"; Break}
					21  {$tmp2 = "Envelope #11"; Break}
					22  {$tmp2 = "Envelope #12"; Break}
					23  {$tmp2 = "Envelope #14"; Break}
					24  {$tmp2 = "C Size Sheet"; Break}
					25  {$tmp2 = "D Size Sheet"; Break}
					26  {$tmp2 = "E Size Sheet"; Break}
					27  {$tmp2 = "Envelope DL"; Break}
					28  {$tmp2 = "Envelope C5"; Break}
					29  {$tmp2 = "Envelope C3"; Break}
					30  {$tmp2 = "Envelope C4"; Break}
					31  {$tmp2 = "Envelope C6"; Break}
					32  {$tmp2 = "Envelope C65"; Break}
					33  {$tmp2 = "Envelope B4"; Break}
					34  {$tmp2 = "Envelope B5"; Break}
					35  {$tmp2 = "Envelope B6"; Break}
					36  {$tmp2 = "Envelope Italy"; Break}
					37  {$tmp2 = "Envelope Monarch"; Break}
					38  {$tmp2 = "Envelope Personal"; Break}
					39  {$tmp2 = "US Std Fanfold"; Break}
					40  {$tmp2 = "German Std Fanfold"; Break}
					41  {$tmp2 = "German Legal Fanfold"; Break}
					42  {$tmp2 = "B4 (ISO)"; Break}
					43  {$tmp2 = "Japanese Postcard"; Break}
					44  {$tmp2 = "9X11"; Break}
					45  {$tmp2 = "10X11"; Break}
					46  {$tmp2 = "15X11"; Break}
					47  {$tmp2 = "Envelope Invite"; Break}
					48  {$tmp2 = "Reserved - DO NOT USE"; Break}
					49  {$tmp2 = "Reserved - DO NOT USE"; Break}
					50  {$tmp2 = "Letter Extra"; Break}
					51  {$tmp2 = "Legal Extra"; Break}
					52  {$tmp2 = "Tabloid Extra"; Break}
					53  {$tmp2 = "A4 Extra"; Break}
					54  {$tmp2 = "Letter Transverse"; Break}
					55  {$tmp2 = "A4 Transverse"; Break}
					56  {$tmp2 = "Letter Extra Transverse"; Break}
					57  {$tmp2 = "A Plus"; Break}
					58  {$tmp2 = "B Plus"; Break}
					59  {$tmp2 = "Letter Plus"; Break}
					60  {$tmp2 = "A4 Plus"; Break}
					61  {$tmp2 = "A5 Transverse"; Break}
					62  {$tmp2 = "B5 (JIS) Transverse"; Break}
					63  {$tmp2 = "A3 Extra"; Break}
					64  {$tmp2 = "A5 Extra"; Break}
					65  {$tmp2 = "B5 (ISO) Extra"; Break}
					66  {$tmp2 = "A2"; Break}
					67  {$tmp2 = "A3 Transverse"; Break}
					68  {$tmp2 = "A3 Extra Transverse"; Break}
					69  {$tmp2 = "Japanese Double Postcard"; Break}
					70  {$tmp2 = "A6"; Break}
					71  {$tmp2 = "Japanese Envelope Kaku #2"; Break}
					72  {$tmp2 = "Japanese Envelope Kaku #3"; Break}
					73  {$tmp2 = "Japanese Envelope Chou #3"; Break}
					74  {$tmp2 = "Japanese Envelope Chou #4"; Break}
					75  {$tmp2 = "Letter Rotated"; Break}
					76  {$tmp2 = "A3 Rotated"; Break}
					77  {$tmp2 = "A4 Rotated"; Break}
					78  {$tmp2 = "A5 Rotated"; Break}
					79  {$tmp2 = "B4 (JIS) Rotated"; Break}
					80  {$tmp2 = "B5 (JIS) Rotated"; Break}
					81  {$tmp2 = "Japanese Postcard Rotated"; Break}
					82  {$tmp2 = "Double Japanese Postcard Rotated"; Break}
					83  {$tmp2 = "A6 Rotated"; Break}
					84  {$tmp2 = "Japanese Envelope Kaku #2 Rotated"; Break}
					85  {$tmp2 = "Japanese Envelope Kaku #3 Rotated"; Break}
					86  {$tmp2 = "Japanese Envelope Chou #3 Rotated"; Break}
					87  {$tmp2 = "Japanese Envelope Chou #4 Rotated"; Break}
					88  {$tmp2 = "B6 (JIS)"; Break}
					89  {$tmp2 = "B6 (JIS) Rotated"; Break}
					90  {$tmp2 = "12X11"; Break}
					91  {$tmp2 = "Japanese Envelope You #4"; Break}
					92  {$tmp2 = "Japanese Envelope You #4 Rotated"; Break}
					93  {$tmp2 = "PRC 16K"; Break}
					94  {$tmp2 = "PRC 32K"; Break}
					95  {$tmp2 = "PRC 32K(Big)"; Break}
					96  {$tmp2 = "PRC Envelope #1"; Break}
					97  {$tmp2 = "PRC Envelope #2"; Break}
					98  {$tmp2 = "PRC Envelope #3"; Break}
					99  {$tmp2 = "PRC Envelope #4"; Break}
					100 {$tmp2 = "PRC Envelope #5"; Break}
					101 {$tmp2 = "PRC Envelope #6"; Break}
					102 {$tmp2 = "PRC Envelope #7"; Break}
					103 {$tmp2 = "PRC Envelope #8"; Break}
					104 {$tmp2 = "PRC Envelope #9"; Break}
					105 {$tmp2 = "PRC Envelope #10"; Break}
					106 {$tmp2 = "PRC 16K Rotated"; Break}
					107 {$tmp2 = "PRC 32K Rotated"; Break}
					108 {$tmp2 = "PRC 32K(Big) Rotated"; Break}
					109 {$tmp2 = "PRC Envelope #1 Rotated"; Break}
					110 {$tmp2 = "PRC Envelope #2 Rotated"; Break}
					111 {$tmp2 = "PRC Envelope #3 Rotated"; Break}
					112 {$tmp2 = "PRC Envelope #4 Rotated"; Break}
					113 {$tmp2 = "PRC Envelope #5 Rotated"; Break}
					114 {$tmp2 = "PRC Envelope #6 Rotated"; Break}
					115 {$tmp2 = "PRC Envelope #7 Rotated"; Break}
					116 {$tmp2 = "PRC Envelope #8 Rotated"; Break}
					117 {$tmp2 = "PRC Envelope #9 Rotated"; Break}
					Default {$tmp2 = "Paper Size could not be determined: $($xelement) "; Break}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"form"
		{
			$txt="Form Name: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				If($tmp2.length -gt 0)
				{
					$ReturnStr = "$txt $tmp2"
				}
			}
		}
		"true"
		{
			$txt="TrueType: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Bitmap"; Break}
					2 {$tmp2 = "Download"; Break}
					3 {$tmp2 = "Substitute"; Break}
					4 {$tmp2 = "Outline"; Break}
					Default {$tmp2 = "TrueType could not be determined: $($xelement) "; Break}
				}
			}
			$ReturnStr = "$txt $tmp2"
		}
		"mode" 
		{
			$txt="Printer Model: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"loca" 
		{
			$txt="Location: "
			$index = $xelement.SubString(0).IndexOf('=')
			If($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				If($tmp2.length -gt 0)
				{
					$ReturnStr = "$txt $tmp2"
				}
			}
		}
		Default {$ReturnStr = "Session printer setting could not be determined: $($xelement) "}
	}
	Return $ReturnStr
}

Function GetCtxGPOsInAD
{
	#thanks to the Citrix Engineering Team for pointers and for Michael B. Smith for creating the function
	#updated 07-Nov-13 to work in a Windows Workgroup environment
	#update 12-Dec-2018 to work in PoSH V2
	Write-Host "$(Get-Date): Testing for an Active Directory environment" -BackgroundColor Black -ForegroundColor Yellow
	$root = [ADSI]"LDAP://RootDSE"
	If([String]::IsNullOrEmpty($root.PSBase.Name))
	{
		Write-Host "$(Get-Date): `tNot in an Active Directory environment" -BackgroundColor Black -ForegroundColor Yellow
		$root = $Null
		$xArray = @()
	}
	Else
	{
		Write-Host "$(Get-Date): `tIn an Active Directory environment" -BackgroundColor Black -ForegroundColor Yellow
		$domainNC = $root.Properties[ 'defaultNamingContext' ].Value
		$root = $Null
		$xArray = @()

		$domain = $domainNC.Replace( 'DC=', '' ).Replace( ',', '.' )
		Write-Host "$(Get-Date): `tSearching \\$($domain)\sysvol\$($domain)\Policies" -BackgroundColor Black -ForegroundColor Yellow
		$sysvolFiles = @()
		$sysvolFiles = Get-ChildItem -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies' ) -EA 0
		If($sysvolFiles.Count -eq 0)
		{
			Write-Host "$(Get-Date): `tSearch timed out.  Retrying.  Searching \\ + $($domain)\sysvol\$($domain)\Policies a second time." -BackgroundColor Black -ForegroundColor Yellow
			$sysvolFiles = Get-ChildItem -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies' ) -EA 0
		}
		ForEach( $file in $sysvolFiles )
		{
			If( -not $file.PSIsContainer )
			{
				#$file.FullName  ### name of the policy file
				If( $file.FullName -like "*\Citrix\GroupPolicy\Policies.gpf" )
				{
					#"have match " + $file.FullName ### name of the Citrix policies file
					$array = $file.FullName.Split( '\' )
					If( $array.Length -gt 7 )
					{
						$gp = $array[ 6 ].ToString()
						$gpObject = [ADSI]( "LDAP://" + "CN=" + $gp + ",CN=Policies,CN=System," + $domainNC )
						#If(!$xArray.Contains($gpObject.DisplayName))
						#If(!$xArray -Contains $gpObject.DisplayName)
						#{
						#	$xArray += $gpObject.DisplayName	### name of the group policy object
						#}
						$dispName = $gpObject.Properties[ 'displayName' ][0]
						If(!( $xArray –Contains $dispName ) )
						{
							$xArray += $dispName ### name of the group policy object
						}
					}
				}
			}
		}
	}
	Return ,$xArray
}
#endregion

#region Appendix A functions
Function OutputAppendixA
{
	If(!$Summary -and ($Section -eq "All"))
	{
		#	The Session Sharing Key is generated by the XML Broker in XenApp 6.5.  
		#	Web Interface or StoreFront send the following information to the XML Broker:"
		#	Audio Quality (Policy Setting)"
		#	Client Printer Port Mapping (Policy Setting)"
		#	Client Printer Spooling (Policy Setting)"
		#	Color Depth (Application Setting)"
		#	COM Port Mapping (Policy Setting)"
		#	Display Size (Application Setting)"
		#	Domain Name (Logon)"
		#	EnableSessionSharing (ICA file or Client Registry Setting)"
		#	Encryption Level (Application Setting and Policy Setting.  Policy wins.)"
		#	Farm Name (Web Interface/StoreFront)"
		#	Special Folder Redirection (Policy Setting)"
		#	TWIDisableSessionSharing(ICA file or Client Registry Setting)"
		#	User Name (Logon)"
		#	Virtual COM Port Emulation (Policy Setting)"
		#
		#	This table consists of the above application settings plus
		#	the application settings from CTX159159
		#	Color depth
		#	Screen Size
		#	Access Control Filters (for SmartAccess)
		#	Encryption
		#
		#	In addition, a XenApp server can have Session Sharing disabled in a registry key
		#	To disable session sharing, the following registry key must be present.
		#	This information has been added to the Server Appendix B section
		#
		#	Add the following value to disable this feature (this value does not exist by default):
		#	HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Citrix\Wfshell\TWI\:
		#	Type: REG_DWORD
		#	Value: SeamlessFlags = 1

		Write-Host "$(Get-Date): Create Appendix A Session Sharing Items" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWord -or $PDF)
		{
			$selection.InsertNewPage()
			WriteWordLine 1 0 "Appendix A - Session Sharing Items from CTX159159"
			## Create an array of hashtables to store our services
			[System.Collections.Hashtable[]] $ItemsWordTable = @();
		
			ForEach($Item in $Script:SessionSharingItems)
			{
				If($Item.AccessControlFilters -is [array])
				{
					$cnt = -1
					ForEach($x in $Item.AccessControlFilters)
					{
						$cnt++
						If($cnt -eq 0)
						{
							$WordTableRowHash = @{ ApplicationName = $Item.ApplicationName;
							MaximumColorQuality = $Item.MaximumColorQuality;
							SessionWindowSize = $Item.SessionWindowSize; 
							AccessControlFilters = $x;
							Encryption = $Item.Encryption}
							$ItemsWordTable += $WordTableRowHash;
						}
						Else
						{
							$WordTableRowHash = @{ ApplicationName = "";
							MaximumColorQuality = "";
							SessionWindowSize = ""; 
							AccessControlFilters = $x;
							Encryption = ""}
							$ItemsWordTable += $WordTableRowHash;
						}
					}
				}
				Else
				{
					$WordTableRowHash = @{ ApplicationName = $Item.ApplicationName;
					MaximumColorQuality = $Item.MaximumColorQuality;
					SessionWindowSize = $Item.SessionWindowSize; 
					AccessControlFilters = $Item.AccessControlFilters;
					Encryption = $Item.Encryption}
					$ItemsWordTable += $WordTableRowHash;
				}
			}

			## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ApplicationName, MaximumColorQuality, SessionWindowSize, AccessControlFilters, Encryption `
			-Headers "Application Name", "Maximum color quality", "Session window size", "Access Control Filters", "Encryption" `
			-AutoFit $wdAutoFitContent;

			## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		ElseIf($Text)
		{
			Line 0 "Appendix A - Session Sharing Items from CTX159159"
			ForEach($Item in $Script:SessionSharingItems)
			{
				If($Item.AccessControlFilters -is [array])
				{
					Line 1 "Application Name`t: " $Item.ApplicationName
					Line 1 "Maximum color quality`t: " $Item.MaximumColorQuality
					Line 1 "Session window size`t: " $Item.SessionWindowSize
					$cnt = -1
					ForEach($AccessCondition in $Item.AccessControlFilters)
					{
						$cnt++
						If($cnt -eq 0)
						{
							[string]$Tmp = $AccessCondition
							[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
							[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
							Line 1 "Access Control Filters`t: $($AGFarm) $($AGFilter)"
						}
						Else
						{
							[string]$Tmp = $AccessCondition
							[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
							[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
							Line 4 "  $($AGFarm) $($AGFilter)"
						}
					}
					Line 1 "Encryption`t`t: " $Item.Encryption
					Line 0 ""
				}
				Else
				{
					Line 1 "Application Name`t: " $Item.ApplicationName
					Line 1 "Maximum color quality`t: " $Item.MaximumColorQuality
					Line 1 "Session window size`t: " $Item.SessionWindowSize
					Line 1 "Access Control Filters`t: " $Item.AccessControlFilters
					Line 1 "Encryption`t`t: " $Item.Encryption
					Line 0 ""
				}
			}
			$tmp = $Null
			$AGFarm = $Null
			$AGFilter = $Null
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 1 0 "Appendix A - Session Sharing Items from CTX159159"
			$rowdata = @()
			$columnHeaders = @(
			'Application Name',($htmlsilver -bor $htmlbold),
			'Maximum color quality',($htmlsilver -bor $htmlbold),
			'Session window size',($htmlsilver -bor $htmlbold),
			'Access Control Filters',($htmlsilver -bor $htmlbold),
			'Encryption',($htmlsilver -bor $htmlbold))
		
			ForEach($Item in $Script:SessionSharingItems)
			{
				If($Item.AccessControlFilters -is [array])
				{
					$cnt = -1
					ForEach($x in $Item.AccessControlFilters)
					{
						$cnt++
						If($cnt -eq 0)
						{
							$rowdata += @(,(
							$Item.ApplicationName,$htmlwhite,
							$Item.MaximumColorQuality,$htmlwhite,
							$Item.SessionWindowSize,$htmlwhite,
							$x,$htmlwhite,
							$Item.Encryption,$htmlwhite))
						}
						Else
						{
							$rowdata += @(,(
							"",$htmlwhite,
							"",$htmlwhite,
							"",$htmlwhite,
							$x,$htmlwhite,
							"",$htmlwhite))
						}
					}
				}
				Else
				{
					$rowdata += @(,(
					$Item.ApplicationName,$htmlwhite,
					$Item.MaximumColorQuality,$htmlwhite,
					$Item.SessionWindowSize,$htmlwhite,
					$Item.AccessControlFilters,$htmlwhite,
					$Item.Encryption,$htmlwhite))
				}
			}
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 ""

		}
		
		Write-Host "$(Get-Date): Finished Create Appendix A - Session Sharing Items" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}
#endregion

#region Appendix B functions
Function OutputAppendixB
{
	If(!$Summary -and ($Section -eq "All"))
	{
		Write-Host "$(Get-Date): Create Appendix B Server Major Items" -BackgroundColor Black -ForegroundColor Yellow
		If($MSWord -or $PDF)
		{
			$selection.InsertNewPage()
			WriteWordLine 1 0 "Appendix B - Server Major Items"
			## Create an array of hashtables to store our services
			[System.Collections.Hashtable[]] $ItemsWordTable = @();
			## Seed the row index from the second row
			[int] $CurrentServiceIndex = 2;

			$Tmp = ""
			ForEach($Item in $ServerItems)
			{
				$Tmp = $Null
				If([String]::IsNullOrEmpty($Item.LicenseServer))
				{
					$Tmp = "Set by policy"
				}
				Else
				{
					$Tmp = $Item.LicenseServer
				}
				## Add the required key/values to the hashtable
				$WordTableRowHash = @{ ServerName = $Item.ServerName;
				ZoneName = $Item.ZoneName;
				OSVersion = $Item.OSVersion;
				CitrixVersion = $Item.CitrixVersion;
				ProductEdition = $Item.ProductEdition;
				LicenseServer = $Tmp
				SessionSharing = $Item.SessionSharing}
				## Add the hash to the array
				$ItemsWordTable += $WordTableRowHash;

				$CurrentServiceIndex++;
				$Tmp = $Null
			}

			## Add the table to the document, using the hashtable (-Alt is short for -AlternateBackgroundColor!)
			$Table = AddWordTable -Hashtable $ItemsWordTable `
			-Columns ServerName, ZoneName, OSVersion, CitrixVersion, ProductEdition, LicenseServer, SessionSharing `
			-Headers "Server Name", "Zone Name", "OS Version", "Citrix Version", "Product Edition", "License Server", "Session Sharing" `
			-AutoFit $wdAutoFitContent;

			## IB - Set the header row format after the SetWordTableAlternateRowColor function as it will paint the header row!
			SetWordCellFormat -Collection $Table.Rows.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

			$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

			FindWordDocumentEnd
			$Table = $Null
		}
		ElseIf($Text)
		{
			Line 0 "Appendix B - Server Major Items"

			$Tmp = ""
			ForEach($Item in $ServerItems)
			{
				If([String]::IsNullOrEmpty($Item.LicenseServer))
				{
					$Tmp = "Set by policy"
				}
				Else
				{
					$Tmp = $Item.LicenseServer
				}
				Line 1 "Server Name`t: " $Item.ServerName
				Line 1 "Zone Name`t: " $Item.ZoneName
				Line 1 "OS Version`t: " $Item.OSVersion
				Line 1 "Citrix Version`t: " $Item.CitrixVersion
				Line 1 "Product Edition`t: " $Item.ProductEdition
				Line 1 "License Server`t: " $Tmp
				Line 1 "Session Sharing`t: " $Item.SessionSharing
				Line 0 ""
				$Tmp = $Null
			}
		}
		ElseIf($HTML)
		{
			WriteHTMLLine 1 0 "Appendix B - Server Major Items"

			$rowdata = @()
			$columnHeaders = @(
			'Server Name',($htmlsilver -bor $htmlbold),
			'Zone Name',($htmlsilver -bor $htmlbold),
			'OS Version',($htmlsilver -bor $htmlbold),
			'Citrix Version',($htmlsilver -bor $htmlbold),
			'Product Edition',($htmlsilver -bor $htmlbold),
			'License Server',($htmlsilver -bor $htmlbold),
			'Session Sharing',($htmlsilver -bor $htmlbold))

			ForEach($Item in $ServerItems)
			{
				If([String]::IsNullOrEmpty($Item.LicenseServer))
				{
					$Tmp = "Set by policy"
				}
				Else
				{
					$Tmp = $Item.LicenseServer
				}
				$rowdata += @(,(
				$Item.ServerName,$htmlwhite,
				$Item.ZoneName,$htmlwhite,
				$Item.OSVersion,$htmlwhite,
				$Item.CitrixVersion,$htmlwhite,
				$Item.ProductEdition,$htmlwhite,
				$Tmp,$htmlwhite,
				$Item.SessionSharing,$htmlwhite))
				$Tmp = $Null
			}
			$msg = ""
			FormatHTMLTable $msg "auto" -rowArray $rowdata -columnArray $columnHeaders
			WriteHTMLLine 0 0 ""
		}
		
		Write-Host "$(Get-Date): Finished Create Appendix B - Server Major Items" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}
#endregion

#region summary page functions
Function ProcessSummaryPage
{
	If($Section -eq "All")
	{
		#summary page
		Write-Host "$(Get-Date): Create Summary Page" -BackgroundColor Black -ForegroundColor Yellow
		If(!$Summary)
		{
			OutputSummaryPage
		}
		Else
		{
			OutputSummarySummaryPage
		}

		Write-Host "$(Get-Date): Finished Create Summary Page" -BackgroundColor Black -ForegroundColor Yellow
		Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow
	}
}

Function OutputSummarySummaryPage
{
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Summary Page"
		WriteWordLine 0 0 "Administrators"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Administrators"; Value = $Script:TotalAdmins; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Applications"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Applications"; Value = $Script:TotalApps; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Load Balancing Policies"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Load Balancing Policies"; Value = $Script:TotalLBPolicies; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Load Evaluators"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Load Evaluators"; Value = $Script:TotalLoadEvaluators; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Servers"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Servers"; Value = $Script:TotalServers; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Worker Groups"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Worker Groups"; Value = $Script:TotalWGs; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Zones"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Zones"; Value = $Script:TotalZones; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Policies"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "IMA Policies"; Value = $Script:TotalIMAPolicies; }
		$ScriptInformation += @{ Data = "Citrix AD Policies Processed"; Value = $Script:TotalADPolicies; }
		$ScriptInformation += @{ Data = "Total Policies"; Value = $Script:TotalPolicies; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 "AD Policies can contain multiple Citrix policies" -fontsize 8
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Summary Page"
		Line 0 "Administrators"
		Line 1 "Total Administrators`t`t: " $Script:TotalAdmins
		Line 0 ""

		Line 0 "Applications"
		Line 1 "Total Applications`t`t: " $Script:TotalApps
		Line 0 ""

		Line 0 "Load Balancing Policies"
		Line 1 "Total Load Balancing Policies`t: " $Script:TotalLBPolicies
		Line 0 ""

		Line 0 "Load Evaluators"
		Line 1 "Total Load Evaluators`t`t: " $Script:TotalLoadEvaluators
		Line 0 ""

		Line 0 "Servers"
		Line 1 "Total Servers`t`t`t: " $Script:TotalServers
		Line 0 ""

		Line 0 "Worker Groups"
		Line 1 "Total Worker Groups`t`t: " $Script:TotalWGs
		Line 0 ""

		Line 0 "Zones"
		Line 1 "Total Zones`t`t`t: " $Script:TotalZones
		Line 0 ""

		Line 0 "Policies"
		Line 1 "IMA Policies`t`t`t: " $Script:TotalIMAPolicies
		Line 1 "Citrix AD Policies Processed`t: $($Script:TotalADPolicies)`t(AD Policies can contain multiple Citrix policies)"
		Line 1 "Total Policies`t`t`t: " $Script:TotalPolicies
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Summary Page"
		$rowdata = @()
		$columnHeaders = @("Total Administrators",($htmlsilver -bor $htmlbold),"$($Script:TotalAdmins)",$htmlwhite)

		$msg = "Administrators"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total Applications",($htmlsilver -bor $htmlbold),"$($Script:TotalApps)",$htmlwhite)

		$msg = "Applications"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total Load Balancing Policies",($htmlsilver -bor $htmlbold),"$($Script:TotalLBPolicies)",$htmlwhite)

		$msg = "Load Balancing Policies"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total Load Evaluators",($htmlsilver -bor $htmlbold),"$($Script:TotalLoadEvaluators)",$htmlwhite)

		$msg = "Load Evaluators"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total Servers",($htmlsilver -bor $htmlbold),"$($Script:TotalServers)",$htmlwhite)

		$msg = "Servers"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total Worker Groups",($htmlsilver -bor $htmlbold),"$($Script:TotalWGs)",$htmlwhite)

		$msg = "Worker Groups"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total Zones",($htmlsilver -bor $htmlbold),"$($Script:TotalZones)",$htmlwhite)

		$msg = "Zones"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("IMA Policies",($htmlsilver -bor $htmlbold),"$($Script:TotalIMAPolicies)",$htmlwhite)
		$rowdata += @(,('Citrix AD Policies Processed',($htmlsilver -bor $htmlbold),"$($Script:TotalADPolicies)",$htmlwhite))
		$rowdata += @(,('Total Policies',($htmlsilver -bor $htmlbold),"$($Script:TotalPolicies)",$htmlwhite))

		$msg = "Policies"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 "AD Policies can contain multiple Citrix policies" -fontsize 1
		WriteHTMLLine 0 0 ""
	}
}

Function OutputSummaryPage
{
	$Script:TotalAdmins = ($Script:TotalFullAdmins + $Script:TotalViewAdmins + $Script:TotalCustomAdmins)
	$Script:TotalApps = ($Script:TotalPublishedApps + $Script:TotalPublishedContent + $Script:TotalPublishedDesktops + $Script:TotalStreamedApps)
	$Script:TotalServers = ($Script:TotalControllers + $Script:TotalWorkers)
	$Script:TotalWGs = ($Script:TotalWGByServerName + $Script:TotalWGByServerGroup + $Script:TotalWGByOU)
	$Script:TotalPolicies = ($Script:TotalComputerPolicies + $Script:TotalUserPolicies)
	
	If($MSWord -or $PDF)
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Summary Page"
		WriteWordLine 0 0 "Administrators"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Full Administrators"; Value = $Script:TotalFullAdmins; }
		$ScriptInformation += @{ Data = "Total View Administrators"; Value = $Script:TotalViewAdmins; }
		$ScriptInformation += @{ Data = "Total Custom Administrators"; Value = $Script:TotalCustomAdmins; }
		$ScriptInformation += @{ Data = "     Total Administrators"; Value = $Script:TotalAdmins; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Applications"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Published Applications"; Value = $Script:TotalPublishedApps; }
		$ScriptInformation += @{ Data = "Total Published Content"; Value = $Script:TotalPublishedContent; }
		$ScriptInformation += @{ Data = "Total Published Desktops"; Value = $Script:TotalPublishedDesktops; }
		$ScriptInformation += @{ Data = "Total Streamed Applications"; Value = $Script:TotalStreamedApps; }
		$ScriptInformation += @{ Data = "     Total Applications"; Value = $TotalApps; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Configuration Logging"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "     Total Config Log Items"; Value = $Script:TotalConfigLogItems; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Load Balancing Policies"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "     Total Load Balancing Policies"; Value = $Script:TotalLBPolicies; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Load Evaluators"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "     Total Load Evaluators"; Value = $Script:TotalLoadEvaluators; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Servers"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Controllers"; Value = $Script:TotalControllers; }
		$ScriptInformation += @{ Data = "Total Workers"; Value = $Script:TotalWorkers; }
		$ScriptInformation += @{ Data = "     Total Servers"; Value = $Script:TotalServers; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Worker Groups"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total WGs by Server Name"; Value = $Script:TotalWGByServerName; }
		$ScriptInformation += @{ Data = "Total WGs by Server Group"; Value = $Script:TotalWGByServerGroup; }
		$ScriptInformation += @{ Data = "Total WGs by AD Container"; Value = $Script:TotalWGByOU; }
		$ScriptInformation += @{ Data = "     Total Worker Groups"; Value = $Script:TotalWGs; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Zones"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "     Total Zones"; Value = $Script:TotalZones; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 ""

		WriteWordLine 0 0 "Policies"
		[System.Collections.Hashtable[]] $ScriptInformation = @()
		$ScriptInformation += @{ Data = "Total Computer Policies"; Value = $Script:TotalComputerPolicies; }
		$ScriptInformation += @{ Data = "Total User Policies"; Value = $Script:TotalUserPolicies; }
		$ScriptInformation += @{ Data = "     Total Policies"; Value = $TotalPolicies; }
		$ScriptInformation += @{ Data = "IMA Policies"; Value = $Script:TotalIMAPolicies; }
		$ScriptInformation += @{ Data = "Citrix AD Policies Processed"; Value = $Script:TotalADPolicies; }
		$ScriptInformation += @{ Data = "Citrix AD Policies not Processed"; Value = $Script:TotalADPoliciesNotProcessed; }

		$Table = AddWordTable -Hashtable $ScriptInformation `
		-Columns Data,Value `
		-List `
		-Format $wdTableGrid `
		-AutoFit $wdAutoFitFixed;

		## IB - Set the header row format
		SetWordCellFormat -Collection $Table.Columns.Item(1).Cells -Bold -BackgroundColor $wdColorGray15;

		$Table.Columns.Item(1).Width = 200;
		$Table.Columns.Item(2).Width = 150;

		$Table.Rows.SetLeftIndent($Indent0TabStops,$wdAdjustProportional)

		FindWordDocumentEnd
		$Table = $Null
		WriteWordLine 0 0 "AD Policies can contain multiple Citrix policies" -fontsize 8
		WriteWordLine 0 0 ""
	}
	ElseIf($Text)
	{
		Line 0 "Summary Page"
		Line 0 "Administrators"
		Line 1 "Total Full Administrators`t: " $Script:TotalFullAdmins
		Line 1 "Total View Administrators`t: " $Script:TotalViewAdmins
		Line 1 "Total Custom Administrators`t: " $Script:TotalCustomAdmins
		Line 2 "Total Administrators`t: " $Script:TotalAdmins
		Line 0 ""

		Line 0 "Applications"
		Line 1 "Total Published Applications`t: " $Script:TotalPublishedApps
		Line 1 "Total Published Content`t`t: " $Script:TotalPublishedContent
		Line 1 "Total Published Desktops`t: " $Script:TotalPublishedDesktops
		Line 1 "Total Streamed Applications`t: " $Script:TotalStreamedApps
		Line 2 "Total Applications`t: " $Script:TotalApps
		Line 0 ""

		Line 0 "Configuration Logging"
		Line 1 "Total Config Log Items`t`t: " $Script:TotalConfigLogItems 
		Line 0 ""

		Line 0 "Load Balancing Policies"
		Line 1 "Total Load Balancing Policies`t: " $Script:TotalLBPolicies
		Line 0 ""

		Line 0 "Load Evaluators"
		Line 1 "Total Load Evaluators`t`t: " $Script:TotalLoadEvaluators
		Line 0 ""

		Line 0 "Servers"
		Line 1 "Total Controllers`t`t: " $Script:TotalControllers
		Line 1 "Total Workers`t`t`t: " $Script:TotalWorkers
		Line 2 "Total Servers`t`t: " $Script:TotalServers
		Line 0 ""

		Line 0 "Worker Groups"
		Line 1 "Total WGs by Server Name`t: " $Script:TotalWGByServerName
		Line 1 "Total WGs by Server Group`t: " $Script:TotalWGByServerGroup
		Line 1 "Total WGs by AD Container`t: " $Script:TotalWGByOU
		Line 2 "Total Worker Groups`t: " $Script:TotalWGs
		Line 0 ""

		Line 0 "Zones"
		Line 1 "Total Zones`t`t`t: " $Script:TotalZones
		Line 0 ""

		Line 0 "Policies"
		Line 1 "Total Computer Policies`t`t: " $Script:TotalComputerPolicies
		Line 1 "Total User Policies`t`t: " $Script:TotalUserPolicies
		Line 2 "Total Policies`t`t: " $Script:TotalPolicies
		Line 1 "IMA Policies`t`t`t: " $Script:TotalIMAPolicies
		Line 1 "Citrix AD Policies Processed`t: $($Script:TotalADPolicies)`t(AD Policies can contain multiple Citrix policies)"
		Line 1 "Citrix AD Policies not Processed: " $Script:TotalADPoliciesNotProcessed
		Line 0 ""
	}
	ElseIf($HTML)
	{
		WriteHTMLLine 1 0 "Summary Page"
		$rowdata = @()
		$columnHeaders = @("Total Full Administrators",($htmlsilver -bor $htmlbold),"$($Script:TotalFullAdmins)",$htmlwhite)
		$rowdata += @(,('Total View Administrators',($htmlsilver -bor $htmlbold),"$($Script:TotalViewAdmins)",$htmlwhite))
		$rowdata += @(,('Total Custom Administrators',($htmlsilver -bor $htmlbold),"$($Script:TotalCustomAdmins)",$htmlwhite))
		$rowdata += @(,('     Total Administrators',($htmlsilver -bor $htmlbold),"$($Script:TotalAdmins)",$htmlwhite))

		$msg = "Administrators"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total Published Applications",($htmlsilver -bor $htmlbold),"$($Script:TotalPublishedApps)",$htmlwhite)
		$rowdata += @(,('Total Published Content',($htmlsilver -bor $htmlbold),"$($Script:TotalPublishedContent)",$htmlwhite))
		$rowdata += @(,('Total Published Desktops',($htmlsilver -bor $htmlbold),"$($Script:TotalPublishedDesktops)",$htmlwhite))
		$rowdata += @(,('Total Streamed Applications',($htmlsilver -bor $htmlbold),"$($Script:TotalStreamedApps)",$htmlwhite))
		$rowdata += @(,('     Total Applications',($htmlsilver -bor $htmlbold),"$($Script:TotalApps)",$htmlwhite))

		$msg = "Applications"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("     Total Config Log Items",($htmlsilver -bor $htmlbold),"$($Script:TotalConfigLogItems)",$htmlwhite)

		$msg = "Configuration Logging"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("     Total Load Balancing Policies",($htmlsilver -bor $htmlbold),"$($Script:TotalLBPolicies)",$htmlwhite)

		$msg = "Load Balancing Policies"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("     Total Load Evaluators",($htmlsilver -bor $htmlbold),"$($Script:TotalLoadEvaluators)",$htmlwhite)

		$msg = "Load Evaluators"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total Controllers",($htmlsilver -bor $htmlbold),"$($Script:TotalControllers)",$htmlwhite)
		$rowdata += @(,('Total Workers',($htmlsilver -bor $htmlbold),"$($Script:TotalWorkers)",$htmlwhite))
		$rowdata += @(,('     Total Servers',($htmlsilver -bor $htmlbold),"$($Script:TotalServers)",$htmlwhite))

		$msg = "Servers"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total WGs by Server Name",($htmlsilver -bor $htmlbold),"$($Script:TotalWGByServerName)",$htmlwhite)
		$rowdata += @(,('Total WGs by Server Group',($htmlsilver -bor $htmlbold),"$($Script:TotalWGByServerGroup)",$htmlwhite))
		$rowdata += @(,('Total WGs by AD Container',($htmlsilver -bor $htmlbold),"$($Script:TotalWGByOU)",$htmlwhite))
		$rowdata += @(,('     Total Worker Groups',($htmlsilver -bor $htmlbold),"$($Script:TotalWGs)",$htmlwhite))

		$msg = "Worker Groups"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("     Total Zones",($htmlsilver -bor $htmlbold),"$($Script:TotalZones)",$htmlwhite)

		$msg = "Zones"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 ""

		$rowdata = @()
		$columnHeaders = @("Total Computer Policies",($htmlsilver -bor $htmlbold),"$($Script:TotalComputerPolicies)",$htmlwhite)
		$rowdata += @(,('Total User Policies',($htmlsilver -bor $htmlbold),"$($Script:TotalUserPolicies)",$htmlwhite))
		$rowdata += @(,('     Total Policies',($htmlsilver -bor $htmlbold),"$($Script:TotalPolicies)",$htmlwhite))
		$rowdata += @(,('IMA Policies',($htmlsilver -bor $htmlbold),"$($Script:TotalIMAPolicies)",$htmlwhite))
		$rowdata += @(,('Citrix AD Policies Processed',($htmlsilver -bor $htmlbold),"$($Script:TotalADPolicies)",$htmlwhite))
		$rowdata += @(,('Citrix AD Policies not Processed',($htmlsilver -bor $htmlbold),"$($Script:TotalADPoliciesNotProcessed)",$htmlwhite))

		$msg = "Policies"
		$columnWidths = @("200","150")
		FormatHTMLTable $msg -rowArray $rowdata -columnArray $columnHeaders -fixedWidth $columnWidths
		WriteHTMLLine 0 0 "AD Policies can contain multiple Citrix policies" -fontsize 1
		WriteHTMLLine 0 0 ""
	}
}
#endregion

#region script end
Function ProcessScriptEnd
{
	Write-Host "$(Get-Date): Script has completed" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): " -BackgroundColor Black -ForegroundColor Yellow

	#clear remote connection if the script set it up
	If(![String]::IsNullOrEmpty($Script:RemoteXAServer))
	{
		Write-Host "$(Get-Date): Clearing remote connection to $Script:RemoteXAServer" -BackgroundColor Black -ForegroundColor Yellow
		Clear-XADefaultComputerName -Scope LocalMachine -EA 0
	}

	#http://poshtips.com/measuring-elapsed-time-in-powershell/
	Write-Host "$(Get-Date): Script started: $($Script:StartTime)" -BackgroundColor Black -ForegroundColor Yellow
	Write-Host "$(Get-Date): Script ended: $(Get-Date)" -BackgroundColor Black -ForegroundColor Yellow
	$runtime = $(Get-Date) - $Script:StartTime
	$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds",
		$runtime.Days,
		$runtime.Hours,
		$runtime.Minutes,
		$runtime.Seconds,
		$runtime.Milliseconds)
	Write-Host "$(Get-Date): Elapsed time: $($Str)" -BackgroundColor Black -ForegroundColor Yellow

	If($Dev)
	{
		Out-File -FilePath $Script:DevErrorFile -InputObject $error -Append
	}

	If($ScriptInfo)
	{
		$SIFile = "$($pwd.Path)\XA65V5InventoryScriptInfo_$(Get-Date -f yyyy-MM-dd_HHmm).txt"
		Out-File -FilePath $SIFile -InputObject ""
		Out-File -FilePath $SIFile -Append -InputObject "Add DateTime       : $($AddDateTime)"
		Out-File -FilePath $SIFile -Append -InputObject "AdminAddress       : $($AdminAddress)"
		Out-File -FilePath $SIFile -Append -InputObject "Administrators     : $($Administrators)"
		Out-File -FilePath $SIFile -Append -InputObject "Applications       : $($Applications)"
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Company Name       : $($Script:CoName)"		
			Out-File -FilePath $SIFile -Append -InputObject "Company Address    : $($CompanyAddress)"		
			Out-File -FilePath $SIFile -Append -InputObject "Company Email      : $($CompanyEmail)"		
			Out-File -FilePath $SIFile -Append -InputObject "Company Fax        : $($CompanyFax)"		
			Out-File -FilePath $SIFile -Append -InputObject "Company Phone      : $($CompanyPhone)"		
			Out-File -FilePath $SIFile -Append -InputObject "Cover Page         : $($CoverPage)"
		}
		Out-File -FilePath $SIFile -Append -InputObject "Dev                : $($Dev)"
		If($Dev)
		{
			Out-File -FilePath $SIFile -Append -InputObject "DevErrorFile       : $($Script:DevErrorFile)"
		}
		Out-File -FilePath $SIFile -Append -InputObject "Filename1          : $($Script:FileName1)"
		If($PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Filename2          : $($Script:FileName2)"
		}
		Out-File -FilePath $SIFile -Append -InputObject "Folder             : $($Folder)"
		Out-File -FilePath $SIFile -Append -InputObject "HW Inventory       : $($Hardware)"
		Out-File -FilePath $SIFile -Append -InputObject "Log                : $($Log)"
		Out-File -FilePath $SIFile -Append -InputObject "Logging            : $($Logging)"
		If($Logging)
		{
			Out-File -FilePath $SIFile -Append -InputObject "   Start Date      : $($StartDate)"
			Out-File -FilePath $SIFile -Append -InputObject "   End Date        : $($EndDate)"
		}
		Out-File -FilePath $SIFile -Append -InputObject "MaxDetails         : $($MaxDetails)"
		Out-File -FilePath $SIFile -Append -InputObject "NoADPolicies       : $($NoADPolicies)"
		Out-File -FilePath $SIFile -Append -InputObject "NoPolicies         : $($NoPolicies)"
		Out-File -FilePath $SIFile -Append -InputObject "Policies           : $($Policies)"
		Out-File -FilePath $SIFile -Append -InputObject "Save As HTML       : $($HTML)"
		Out-File -FilePath $SIFile -Append -InputObject "Save As PDF        : $($PDF)"
		Out-File -FilePath $SIFile -Append -InputObject "Save As TEXT       : $($TEXT)"
		Out-File -FilePath $SIFile -Append -InputObject "Save As WORD       : $($MSWORD)"
		Out-File -FilePath $SIFile -Append -InputObject "Script Info        : $($ScriptInfo)"
		Out-File -FilePath $SIFile -Append -InputObject "Section            : $($Section)"
		Out-File -FilePath $SIFile -Append -InputObject "Title              : $($Script:Title)"
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "User Name          : $($UserName)"
		}
		Out-File -FilePath $SIFile -Append -InputObject ""
		Out-File -FilePath $SIFile -Append -InputObject "OS Detected        : $($Script:RunningOS)"
		Out-File -FilePath $SIFile -Append -InputObject "PoSH version       : $($Host.Version)"
		Out-File -FilePath $SIFile -Append -InputObject "PSCulture          : $($PSCulture)"
		Out-File -FilePath $SIFile -Append -InputObject "PSUICulture        : $($PSUICulture)"
		If($MSWORD -or $PDF)
		{
			Out-File -FilePath $SIFile -Append -InputObject "Word language      : $($Script:WordLanguageValue)"
			Out-File -FilePath $SIFile -Append -InputObject "Word version       : $($Script:WordProduct)"
		}
		Out-File -FilePath $SIFile -Append -InputObject ""
		Out-File -FilePath $SIFile -Append -InputObject "Script start       : $($Script:StartTime)"
		Out-File -FilePath $SIFile -Append -InputObject "Elapsed time       : $($Str)"
	}

	#stop transcript logging
	If($Log -eq $True) 
	{
		If($Script:StartLog -eq $true) 
		{
			try 
			{
				Stop-Transcript | Out-Null
				Write-Host "$(Get-Date): $Script:LogPath is ready for use" -BackgroundColor Black -ForegroundColor Yellow
			} 
			catch 
			{
				Write-Host "$(Get-Date): Transcript/log stop failed" -BackgroundColor Black -ForegroundColor Yellow
			}
		}
	}
	$ErrorActionPreference = $SaveEAPreference
}
#endregion

#region script core
#Script begins

ProcessScriptSetup

SetFileName1andFileName2 "$($Script:FarmName)"

ProcessConfigLogSettings

ProcessAdministrators

ProcessApplications

ProcessConfigLogging

ProcessLoadBalancingPolicies

ProcessLoadEvaluators

ProcessServers

ProcessWorkerGroups

ProcessZones

If($Section -eq "All" -or $Section -eq "Policies")
{
	If($NoPolicies -or $Script:DoPolicies -eq $False)
	{
		#don't process policies
	}
	Else
	{
		ProcessPolicies
	}
}

OutputAppendixA

OutputAppendixB

ProcessSummaryPage
#endregion

#region finish script
Write-Host "$(Get-Date): Finishing up document" -BackgroundColor Black -ForegroundColor Yellow
#end of document processing

$AbstractTitle = "Citrix XenApp 6.5 Inventory"
$SubjectTitle = "XenApp 6.5 Farm Inventory"
UpdateDocumentProperties $AbstractTitle $SubjectTitle

ProcessDocumentOutput

ProcessScriptEnd
#endregion