<#
.SYNOPSIS
	Creates a complete inventory of a Citrix XenApp 6.5 farm using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a Citrix XenApp 6.5 farm using Microsoft Word and PowerShell.
	Creates a Word document named after the XenApp 6.5 farm.
	Document includes a Cover Page, Table of Contents and Footer.
.PARAMETER CompanyName
	Company Name to use for the Cover Page.  
	Default value is contained in HKCU:\Software\Microsoft\Office\Common\UserInfo\CompanyName or
	HKCU:\Software\Microsoft\Office\Common\UserInfo\Company, whichever is populated on the 
	computer running the script.
	This parameter has an alias of CN.
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	(default cover pages in Word en-US)
	Valid input is:
		Alphabet (Word 2007/2010. Works)
		Annual (Word 2007/2010. Doesn't really work well for this report)
		Austere (Word 2007/2010. Works)
		Austin (Word 2010/2013. Doesn't work in 2013, mostly works in 2007/2010 but Subtitle/Subject & Author fields need to me moved after title box is moved up)
		Banded (Word 2013. Works)
		Conservative (Word 2007/2010. Works)
		Contrast (Word 2007/2010. Works)
		Cubicles (Word 2007/2010. Works)
		Exposure (Word 2007/2010. Works if you like looking sideways)
		Facet (Word 2013. Works)
		Filigree (Word 2013. Works)
		Grid (Word 2010/2013.Works in 2010)
		Integral (Word 2013. Works)
		Ion (Dark) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Ion (Light) (Word 2013. Top date doesn't fit, box needs to be manually resized or font changed to 8 point)
		Mod (Word 2007/2010. Works)
		Motion (Word 2007/2010/2013. Works if top date is manually changed to 36 point)
		Newsprint (Word 2010. Works but date is not populated)
		Perspective (Word 2010. Works)
		Pinstripes (Word 2007/2010. Works)
		Puzzle (Word 2007/2010. Top date doesn't fit, box needs to be manually resized or font changed to 14 point)
		Retrospect (Word 2013. Works)
		Semaphore (Word 2013. Works)
		Sideline (Word 2007/2010/2013. Doesn't work in 2013, works in 2007/2010)
		Slice (Dark) (Word 2013. Doesn't work)
		Slice (Light) (Word 2013. Doesn't work)
		Stacks (Word 2007/2010. Works)
		Tiles (Word 2007/2010. Date doesn't fit unless changed to 26 point)
		Transcend (Word 2007/2010. Works)
		ViewMaster (Word 2013. Works)
		Whisp (Word 2013. Works)
	Default value is Motion.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V31.ps1
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V31.ps1 -verbose
	
	Will use all default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Motion for the Cover Page format.
	Administrator for the User Name.
	Will display verbose messages as the script is running.
.EXAMPLE
	PS C:\PSScript .\XA65_Inventory_V31.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\XA65_Inventory_V31.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.LINK
	http://www.carlwebster.com/documenting-a-citrix-xenapp-6-5-farm-with-microsoft-powershell-and-word-version-3-1
.NOTES
	NAME: XA65_Inventory_V31.ps1
	VERSION: 3.14
	AUTHOR: Carl Webster (with a lot of help from Michael B. Smith and Jeff Wouters)
	LASTEDIT: July 1, 2013
.REMARKS
	To see the examples, type: "Get-Help .\XA65_Inventory_V31.ps1 -examples".
	For more information, type: "Get-Help .\XA65_Inventory_V31.ps1 -detailed".
#>


#thanks to @jeffwouters for helping me with these parameters
[CmdletBinding( SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "" ) ]

Param(	[parameter(
	Position = 0, 
	Mandatory=$false )
	] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(
	Position = 1, 
	Mandatory=$false )
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Motion", 

	[parameter(
	Position = 2, 
	Mandatory=$false )
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username )

	
Set-StrictMode -Version 2

#Original Script created 8/17/2010 by Michael Bogobowicz, Citrix Systems.
#To contact, please message @mikebogo on Twitter
#This script is designed to be run on a XenApp 6.5 server

#Modifications by Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#modified from the original script for XenApp 6.5
#Word version of script based on version 2 of XA65 script
#updated February 18, 2013:
#	Fixed typos
#	Add more write-verbose statements
#	Fixed issues found by running in set-strictmode -version 2.0
#	Test for CompanyName in two different registry locations
#	Test if template DOTX file loads properly.  If not, skip Cover Page and Table of Contents
#	Disable Spell and Grammer Check to resolve issue and improve performance (from Pat Coughlin)
#	Added in the missing Load evaluator settings for Load Throttling and Server User Load 
#	Test XenApp server for availability before getting services and hotfixes
#	Move table of Citrix services to align with text above table
#	Created a table for Citrix installed hotfixes
#	Created a table for Microsoft hotfixes
#Updated March 14, 2013
#	?{?_.SessionId -eq $SessionID} should have been ?{$_.SessionId -eq $SessionID} in the CheckWordPrereq function
#Updated March 15, 2013
#	Include updated hotfix lists from CTX129229
#Updated April 21, 2013
#	Fixed a compatibility issue with the way the Word file was saved and Set-StrictMode -Version 2
#Updated May 4, 2013
#	Include updated hotfix lists from CTX129229
#Updated June 7, 2013
#	Fixed the content of and the detail contained in the Table of Contents
#	Citrix services that are Stopped will now show in a Red cell with bold, black text
#	Recommended hotfixes that are Not Installed will now show in a Red cell with bold, black text
#	Added a few more Write-Verbose statements
#Updated July 1, 2013
#	Include updated hotfix lists from CTX129229


Function CheckWordPrereq
{
	if ((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "This script directly outputs to Microsoft Word, please install Microsoft Word"
		exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $null
	if ($wordrunning)
	{
		Write-Host "Please close all instances of Microsoft Word before running this report."
		exit
	}
}

Function ValidateCompanyName
{
	$xResult = Test-RegistryValue "HKCU:\Software\Microsoft\Office\Common\UserInfo" "CompanyName"
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

#http://stackoverflow.com/questions/5648931/test-if-registry-value-exists
# This function just gets $true or $false
function Test-RegistryValue($path, $name)
{
    $key = Get-Item -LiteralPath $path -EA 0
    $key -and $null -ne $key.GetValue($name, $null)
}

# Gets the specified registry value or $null if it is missing
function Get-RegistryValue($path, $name)
{
    $key = Get-Item -LiteralPath $path -EA 0
    if ($key) {
        $key.GetValue($name, $null)
    }
}
	
Function ValidateCoverPage
{
	Param( [int]$xWordVersion, [string]$xCP )
	
	$xArray = ""
	If( $xWordVersion -eq 15)
	{
		#word 2013
		$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral", "Ion (Dark)", "Ion (Light)", "Motion", "Retrospect", "Semaphore", "Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
	}
	ElseIf( $xWordVersion -eq 14)
	{
		#word 2010
		$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative", "Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint", "Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
	}
	ElseIf( $xWordVersion -eq 12)
	{
		#word 2007
		$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast", "Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend" )
	}
	
	If ($xArray -contains $xCP)
	{
		Return $True
	}
	Else
	{
		Return $False
	}
}

Function MultiPortPolicyPriority
{
	Param( [int]$PriorityValue = 3 )
	
	switch ($PriorityValue)
	{ 
        0 {"Very High"} 
        1 {"High"} 
        2 {"Medium"} 
        3 {"Low"} 
        default {"Unknown Priority Value"}
    }
	Return $PriorityValue
}

Function ConvertNumberToTime
{
	Param( [int]$val = 0 )
	
	#this is stored as a number between 0 (00:00 AM) and 1439 (23:59 PM)
	#180 = 3AM
	#900 = 3PM
	#1027 = 5:07 PM
	#[int] (1027/60) = 17 or 5PM
	#1027 % 60 leaves 7 or 7 minutes
	
	#thanks to MBS for the next line
	$hour = [System.Math]::Floor( ( [int] $val ) / ( [int] 60 ) )
	$minute = $val % 60
	$Strminute = $minute.ToString()
	$tempminute = ""
	If($Strminute.length -lt 2)
	{
		$tempMinute = "0" + $Strminute
	}
	else
	{
		$tempminute = $strminute
	}
	$AMorPM = "AM"
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
	#thanks to MBS for helping me on this function
	Param( [int]$DateAsInteger = 0 )
	
	#this is stored as an integer but is actually a bitmask
	#01/01/2013 = 131924225 = 11111011101 00000001 00000001
	#01/17/2013 = 131924241 = 11111011101 00000001 00010001
	#
	# last 8 bits are the day
	# previous 8 bits are the month
	# the rest (up to 16) are the year
	
	$year     = [Math]::Floor( $DateAsInteger / 65536 )
	$month    = [Math]::Floor( $DateAsInteger / 256 ) % 256
	$day      = $DateAsInteger % 256

	Return "$Month/$Day/$Year"
}
	
Function Check-LoadedModule
#function created by Jeff Wouters
#@JeffWouters on Twitter
#modified by Michael B. Smith to handle when the module doesn't exist on server
#modified by @andyjmorgan
#bug fixed by @schose
#This function handles all three scenarios:
#
# 1. Module is already imported into current session
# 2. Module is not already imported into current session, it does exists on the server and is imported
# 3. Module does not exist on the server

{
	Param( [parameter(Mandatory = $true)][alias("Module")][string]$ModuleName)
	#$LoadedModules = Get-Module | Select Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module |% { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work if the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	if (!$ModuleFound) 
	{
		$module = Import-Module -Name $ModuleName -PassThru -EA 0
		If( $module -and $? )
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
		Return $true
	}
}

Function Check-NeededPSSnapins
{
	Param( [parameter(Mandatory = $true)][alias("Snapin")][string[]]$Snapins)
	
	#function specifics
	$MissingSnapins=@()
	$FoundMissingSnapin=$false
	$LoadedSnapins = @()
	$RegisteredSnapins = @()
    
	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += get-pssnapin | % {$_.name}
	$registeredSnapins += get-pssnapin -Registered | % {$_.name}
    
	foreach ($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		if (!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			if (!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				if (!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}#End Registered If 
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Write-Host "Loading Windows PowerShell snap-in: $snapin"
				Add-PSSnapin -Name $snapin
			}
		}#End Loaded If
		#Snapin is registered and loaded
		else
		{
			write-debug "Windows PowerShell snap-in: $snapin - Already Loaded"
		}
	}#End For
    
	if ($FoundMissingSnapin)
	{
		write-warning "Missing Windows PowerShell snap-ins Detected:"
		$missingSnapins | % {write-warning "($_)"}
		return $False
	}#End If
	Else
	{
		Return $true
	}#End Else
    
}#End Function

Function WriteWordLine
#function created by Ryan Revord
#@rsrevord on Twitter
#function created to make output to Word easy in this script
{
	Param( [int]$style=0, [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "'n", [switch]$nonewline)
	$output=""
	#Build output style
	switch ($style)
	{
		0 {$Selection.Style = "No Spacing"}
		1 {$Selection.Style = "Heading 1"}
		2 {$Selection.Style = "Heading 2"}
		3 {$Selection.Style = "Heading 3"}
		Default {$Selection.Style = "No Spacing"}
	}
	#build # of tabs
	While( $tabs -gt 0 ) { 
		$output += "`t"; $tabs--; 
	}
		
	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
	
	#test for new WriteWordLine 0.
	If($nonewline){
		# Do nothing.
	} Else {
		$Selection.TypeParagraph()
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop=$properties | foreach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$null,$_,$null)
		if ($propname -eq $Name) 
		{
			Return $_
		}
	} #foreach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$null,$prop,$Value)
}

#Script begins

if (!(Check-NeededPSSnapins "Citrix.Common.Commands","Citrix.Common.GroupPolicy","Citrix.XenApp.Commands")){
    #We're missing Citrix Snapins that we need
    write-error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Are you sure you are running this script on a XenApp 6.5 Server? Script will now close."
    break
}

CheckWordPreReq

$Remoting = $False
$tmp = Get-XADefaultComputerName
If(![String]::IsNullOrEmpty( $tmp ))
{
	$Remoting = $True
}

If($Remoting)
{
	write-verbose "Remoting is enabled to XenApp server $tmp"
}
Else
{
	write-verbose "Remoting is not being used"
	
	#now need to make sure the script is not being run on a session-only host
	$ServerName = (Get-Childitem env:computername).value
	$Server = Get-XAServer -ServerName $ServerName
	If($Server.ElectionPreference -eq "WorkerMode")
	{
		Write-Warning "This script cannot be run on a Session-only Host Server if Remoting is not enabled."
		Write-Warning "Use Set-XADefaultComputerName XA65ControllerServerName or run the script on a controller."
		Write-Error "Script cannot continue.  See messages above."
		Exit
	}
}

# Get farm information
write-verbose "Getting Farm data"
$farm = Get-XAFarm -EA 0

If( $? )
{
	#first check to make sure this is a XenApp 6.5 farm
	If($Farm.ServerVersion.ToString().SubString(0,3) -eq "6.5")
	{
		#this is a XenApp 6.5 farm, script can proceed
	}
	Else
	{
		#this is not a XenApp 6.5 farm, script cannot proceed
		write-warning "This script is designed for XenApp 6.5 and should not be run on previous versions of XenApp"
		Return 1
	}
	$FarmName = $farm.FarmName
	$Title="Inventory Report for the $($FarmName) Farm"
	$filename="$($pwd.path)\$($farm.FarmName).docx"
} 
Else 
{
	$FarmName = "Unable to retrieve"
	$Title="XenApp 6.5 Farm Inventory Report"
	$filename="$($pwd.path)\XenApp 6.5 Farm Inventory.docx"
	write-warning "Farm information could not be retrieved"
}
$farm = $null

write-verbose "Setting up Word"
#these values were attained from 
#http://groovy.codehaus.org/modules/scriptom/1.6.0/scriptom-office-2K3-tlb/apidocs/
#http://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
$wdAlignPageNumberRight = 2
$wdColorGray15 = 14277081
$wdFormatDocument = 0
$wdMove = 0
$wdSeekMainDocument = 0
$wdSeekPrimaryFooter = 4
$wdStory = 6
$wdColorRed = 255
$wdColorBlack = 0

# Setup word for output
write-verbose "Create Word comObject.  If you are not running Word 2007, ignore the next message."
$Word = New-Object -comobject "Word.Application"
$WordVersion = [int] $Word.Version
If( $WordVersion -eq 15)
{
	write-verbose "Running Microsoft Word 2013"
	$WordProduct = "Word 2013"
}
Elseif ( $WordVersion -eq 14)
{
	write-verbose "Running Microsoft Word 2010"
	$WordProduct = "Word 2010"
}
Elseif ( $WordVersion -eq 12)
{
	write-verbose "Running Microsoft Word 2007"
	$WordProduct = "Word 2007"
}
Elseif ( $WordVersion -eq 11)
{
	write-verbose "Running Microsoft Word 2003"
	Write-error "This script does not work with Word 2003. Script will end."
	$word.quit()
	exit
}
Else
{
	Write-error "You are running an untested or unsupported version of Microsoft Word.  Script will end.  Please send info on your version of Word to webster@carlwebster.com"
	$word.quit()
	exit
}

write-verbose "Validate company name"
#only validate CompanyName if the field is blank
If([String]::IsNullOrEmpty($CompanyName))
{
	$CompanyName = ValidateCompanyName
	If([String]::IsNullOrEmpty($CompanyName))
	{
		write-error "Company Name cannot be blank.  Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value.  Script cannot continue."
		$Word.Quit()
		exit
	}
}

write-verbose "Validate cover page"
$ValidCP = ValidateCoverPage $WordVersion $CoverPage
If(!$ValidCP)
{
	write-error "For $WordProduct, $CoverPage is not a valid Cover Page option.  Script cannot continue."
	$Word.Quit()
	exit
}

Write-Verbose "Company Name: $CompanyName"
Write-Verbose "Cover Page  : $CoverPage"
Write-Verbose "User Name   : $UserName"
Write-Verbose "Farm Name   : $FarmName"
Write-Verbose "Title       : $Title"
Write-Verbose "Filename    : $filename"

$Word.Visible = $False

#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
#using Jeff's Demo-WordReport.ps1 file for examples
#down to $global:configlog = $false is from Jeff Hicks
write-verbose "Load Word Templates"
$CoverPagesExist = $False
$word.Templates.LoadBuildingBlocks()
If ( $WordVersion -eq 12)
{
	#word 2007
	$BuildingBlocks=$word.Templates | Where {$_.name -eq "Building Blocks.dotx"}
}
Else
{
	#word 2010/2013
	$BuildingBlocks=$word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}
}

If($BuildingBlocks -ne $Null)
{
	$CoverPagesExist = $True
	$part=$BuildingBlocks.BuildingBlockEntries.Item($CoverPage)
}
Else
{
	$CoverPagesExist = $False
}

write-verbose "Create empty word doc"
$Doc = $Word.Documents.Add()
$global:Selection = $Word.Selection

#Disable Spell and Grammer Check to resolve issue and improve performance (from Pat Coughlin)
write-verbose "disable spell checking"
$Word.Options.CheckGrammarAsYouType=$false
$Word.Options.CheckSpellingAsYouType=$false

If($CoverPagesExist)
{
	#insert new page, getting ready for table of contents
	write-verbose "insert new page, getting ready for table of contents"
	$part.Insert($selection.Range,$True) | out-null
	$selection.InsertNewPage()

	#table of contents
	write-verbose "table of contents"
	$toc=$BuildingBlocks.BuildingBlockEntries.Item("Automatic Table 2")
	$toc.insert($selection.Range,$True) | out-null
}
Else
{
	write-verbose "Cover Pages are not installed."
	write-warning "Cover Pages are not installed so this report will not have a cover page."
	write-verbose "Table of Contents are not installed."
	write-warning "Table of Contents are not installed so this report will not have a Table of Contents."
}

#set the footer
write-verbose "set the footer"
[string]$footertext="Report created by $username"

#get the footer
write-verbose "get the footer and format font"
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekPrimaryFooter
#get the footer and format font
$footers=$doc.Sections.Last.Footers
foreach ($footer in $footers) 
{
	if ($footer.exists) 
	{
		$footer.range.Font.name="Calibri"
		$footer.range.Font.size=8
		$footer.range.Font.Italic=$True
		$footer.range.Font.Bold=$True
	}
} #end Foreach
write-verbose "Footer text"
$selection.HeaderFooter.Range.Text=$footerText

#add page numbering
write-verbose "add page numbering"
$selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

#return focus to main document
write-verbose "return focus to main document"
$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

#move to the end of the current document
write-verbose "move to the end of the current document"
$selection.EndKey($wdStory,$wdMove) | Out-Null
#end of Jeff Hicks 

write-verbose "Processing Configuration Logging"
$global:ConfigLog = $False
$ConfigurationLogging = Get-XAConfigurationLog -EA 0

If( $? )
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Configuration Logging"
	If ($ConfigurationLogging.LoggingEnabled ) 
	{
		$global:ConfigLog = $True
		WriteWordLine 0 1 "Configuration Logging is enabled."
		WriteWordLine 0 1 "Allow changes to the farm when logging database is disconnected: " $ConfigurationLogging.ChangesWhileDisconnectedAllowed
		WriteWordLine 0 1 "Require administrator to enter credentials before clearing the log: " $ConfigurationLogging.CredentialsOnClearLogRequired
		WriteWordLine 0 1 "Database type: " $ConfigurationLogging.DatabaseType
		WriteWordLine 0 1 "Authentication mode: " $ConfigurationLogging.AuthenticationMode
		WriteWordLine 0 1 "Connection string: " 
		$Tmp = "`t`t" + $ConfigurationLogging.ConnectionString.replace(";","`n`t`t`t")
		WriteWordLine 0 1 $Tmp -NoNewline
		WriteWordLine 0 0 ""
		WriteWordLine 0 1 "User name: " $ConfigurationLogging.UserName
		$Tmp = $null
	}
	Else 
	{
		WriteWordLine 0 1 "Configuration Logging is disabled."
	}
}
Else 
{
	write-warning  "Configuration Logging could not be retrieved"
}
$ConfigurationLogging = $null

write-verbose "Processing Administrators"
$Administrators = Get-XAAdministrator -EA 0 | sort-object AdministratorName

If( $? )
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Administrators:"
	ForEach($Administrator in $Administrators)
	{
		WriteWordLine 2 0 $Administrator.AdministratorName
		WriteWordLine 0 1 "Administrator type: " -nonewline
		switch ($Administrator.AdministratorType)
		{
			"Unknown"  {WriteWordLine 0 0 "Unknown"}
			"Full"     {WriteWordLine 0 0 "Full Administration"}
			"ViewOnly" {WriteWordLine 0 0 "View Only"}
			"Custom"   {WriteWordLine 0 0 "Custom"}
			Default    {WriteWordLine 0 0 "Administrator type could not be determined: $($Administrator.AdministratorType)"}
		}
		WriteWordLine 0 1 "Administrator account is " -NoNewLine
		If($Administrator.Enabled)
		{
			WriteWordLine 0 0 "Enabled" 
		} 
		Else
		{
			WriteWordLine 0 0 "Disabled" 
		}
		If ($Administrator.AdministratorType -eq "Custom") 
		{
			WriteWordLine 0 1 "Farm Privileges:"
			ForEach($farmprivilege in $Administrator.FarmPrivileges) 
			{
				switch ($farmprivilege)
				{
					"Unknown"                   {WriteWordLine 0 2 "Unknown"}
					"ViewFarm"                  {WriteWordLine 0 2 "View farm management"}
					"EditZone"                  {WriteWordLine 0 2 "Edit zones"}
					"EditConfigurationLog"      {WriteWordLine 0 2 "Configure logging for the farm"}
					"EditFarmOther"             {WriteWordLine 0 2 "Edit all other farm settings"}
					"ViewAdmins"                {WriteWordLine 0 2 "View Citrix administrators"}
					"LogOnConsole"              {WriteWordLine 0 2 "Log on to console"}
					"LogOnWIConsole"            {WriteWordLine 0 2 "Logon on to Web Interface console"}
					"ViewLoadEvaluators"        {WriteWordLine 0 2 "View load evaluators"}
					"AssignLoadEvaluators"      {WriteWordLine 0 2 "Assign load evaluators"}
					"EditLoadEvaluators"        {WriteWordLine 0 2 "Edit load evaluators"}
					"ViewLoadBalancingPolicies" {WriteWordLine 0 2 "View load balancing policies"}
					"EditLoadBalancingPolicies" {WriteWordLine 0 2 "Edit load balancing policies"}
					"ViewPrinterDrivers"        {WriteWordLine 0 2 "View printer drivers"}
					"ReplicatePrinterDrivers"   {WriteWordLine 0 2 "Replicate printer drivers"}
					Default {WriteWordLine 0 2 "Farm privileges could not be determined: $($farmprivilege)"}
				}
			}
	
			WriteWordLine 0 1 "Folder Privileges:"
			ForEach($folderprivilege in $Administrator.FolderPrivileges) 
			{
				#The Citrix PoSH cmdlet only returns data for three folders:
				#Servers
				#WorkerGroups
				#Applications
				
				WriteWordLine 0 2 $FolderPrivilege.FolderPath
				ForEach($FolderPermission in $FolderPrivilege.FolderPrivileges)
				{
					switch ($folderpermission)
					{
						"Unknown"                          {WriteWordLine 0 3 "Unknown"}
						"ViewApplications"                 {WriteWordLine 0 3 "View applications"}
						"EditApplications"                 {WriteWordLine 0 3 "Edit applications"}
						"TerminateProcessApplication"      {WriteWordLine 0 3 "Terminate process that is created as a result of launching a published application"}
						"AssignApplicationsToServers"      {WriteWordLine 0 3 "Assign applications to servers"}
						"ViewServers"                      {WriteWordLine 0 3 "View servers"}
						"EditOtherServerSettings"          {WriteWordLine 0 3 "Edit other server settings"}
						"RemoveServer"                     {WriteWordLine 0 3 "Remove a bad server from farm"}
						"TerminateProcess"                 {WriteWordLine 0 3 "Terminate processes on a server"}
						"ViewSessions"                     {WriteWordLine 0 3 "View ICA/RDP sessions"}
						"ConnectSessions"                  {WriteWordLine 0 3 "Connect sessions"}
						"DisconnectSessions"               {WriteWordLine 0 3 "Disconnect sessions"}
						"LogOffSessions"                   {WriteWordLine 0 3 "Log off sessions"}
						"ResetSessions"                    {WriteWordLine 0 3 "Reset sessions"}
						"SendMessages"                     {WriteWordLine 0 3 "Send messages to sessions"}
						"ViewWorkerGroups"                 {WriteWordLine 0 3 "View worker groups"}
						"AssignApplicationsToWorkerGroups" {WriteWordLine 0 3 "Assign applications to worker groups"}
						Default {WriteWordLine 0 3 "Folder permission could not be determined: $($folderpermissions)"}
					}
				}
			}
		}		
		WriteWordLine 0 0 ""
	}
}
Else 
{
	write-warning "Administrator information could not be retrieved"
}

$Administrators = $null

write-verbose "Processing Applications"
$Applications = Get-XAApplication -EA 0 | sort-object FolderPath, DisplayName

If( $? -and $Applications)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Applications:"

	ForEach($Application in $Applications)
	{
		$AppServerInfoResults = $False
		$AppServerInfo = Get-XAApplicationReport -BrowserName $Application.BrowserName -EA 0
		If( $? )
		{
			$AppServerInfoResults = $True
		}
		$streamedapp = $False
		If($Application.ApplicationType -Contains "streamedtoclient" -or $Application.ApplicationType -Contains "streamedtoserver")
		{
			$streamedapp = $True
		}
		#name properties
		WriteWordLine 2 0 $Application.DisplayName
		WriteWordLine 0 1 "Application name`t`t: " $Application.BrowserName
		WriteWordLine 0 1 "Disable application`t`t: " -NoNewLine
		#weird, if application is enabled, it is disabled!
		If ($Application.Enabled) 
		{
			WriteWordLine 0 0 "No"
		} 
		Else
		{
			WriteWordLine 0 0 "Yes"
			WriteWordLine 0 1 "Hide disabled application`t: " -nonewline
			If($Application.HideWhenDisabled)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}

		If(![String]::IsNullOrEmpty( $Application.Description))
		{
			WriteWordLine 0 1 "Application description`t`t: " $Application.Description
		}
		
		#type properties
		WriteWordLine 0 1 "Application Type`t`t: " -nonewline
		switch ($Application.ApplicationType)
		{
			"Unknown"                            {WriteWordLine 0 0 "Unknown"}
			"ServerInstalled"                    {WriteWordLine 0 0 "Installed application"}
			"ServerDesktop"                      {WriteWordLine 0 0 "Server desktop"}
			"Content"                            {WriteWordLine 0 0 "Content"}
			"StreamedToServer"                   {WriteWordLine 0 0 "Streamed to server"}
			"StreamedToClient"                   {WriteWordLine 0 0 "Streamed to client"}
			"StreamedToClientOrInstalled"        {WriteWordLine 0 0 "Streamed if possible, otherwise accessed from server as Installed application"}
			"StreamedToClientOrStreamedToServer" {WriteWordLine 0 0 "Streamed if possible, otherwise Streamed to server"}
			Default {WriteWordLine 0 0 "Application Type could not be determined: $($Application.ApplicationType)"}
		}
		If(![String]::IsNullOrEmpty( $Application.FolderPath))
		{
			WriteWordLine 0 1 "Folder path`t`t`t: " $Application.FolderPath
		}
		If(![String]::IsNullOrEmpty( $Application.ContentAddress))
		{
			WriteWordLine 0 1 "Content Address`t`t: " $Application.ContentAddress
		}
	
		#if a streamed app
		If($streamedapp)
		{
			WriteWordLine 0 1 "Citrix streaming app profile address`t`t: " 
			WriteWordLine 0 2 $Application.ProfileLocation
			WriteWordLine 0 1 "App to launch from Citrix stream app profile`t: " 
			WriteWordLine 0 2 $Application.ProfileProgramName
			If(![String]::IsNullOrEmpty( $Application.ProfileProgramArguments))
			{
				WriteWordLine 0 1 "Extra command line parameters`t`t`t: " 
				WriteWordLine 0 2 $Application.ProfileProgramArguments
			}
			#if streamed, OffWriteWordLine 0 access properties
			If($Application.OfflineAccessAllowed)
			{
				WriteWordLine 0 1 "Enable offline access`t`t`t`t: " -nonewline
				If($Application.OfflineAccessAllowed)
				{
					WriteWordLine 0 0 "Yes"
				}
				Else
				{
					WriteWordLine 0 0 "No"
				}
			}
			If($Application.CachingOption)
			{
				WriteWordLine 0 1 "Cache preference`t`t`t`t: " -nonewline
				switch ($Application.CachingOption)
				{
					"Unknown"   {WriteWordLine 0 0 "Unknown"}
					"PreLaunch" {WriteWordLine 0 0 "Cache application prior to launching"}
					"AtLaunch"  {WriteWordLine 0 0 "Cache application during launch"}
					Default {WriteWordLine 0 0 "Could not be determined: $($Application.CachingOption)"}
				}
			}
		}
		
		#location properties
		If(!$streamedapp)
		{
			If(![String]::IsNullOrEmpty( $Application.CommandLineExecutable))
			{
				If($Application.CommandLineExecutable.Length -lt 40)
				{
					WriteWordLine 0 1 "Command Line`t`t`t: " $Application.CommandLineExecutable
				}
				Else
				{
					WriteWordLine 0 1 "Command Line: " 
					WriteWordLine 0 2 $Application.CommandLineExecutable
				}
			}
			If(![String]::IsNullOrEmpty( $Application.WorkingDirectory))
			{
				If($Application.WorkingDirectory.Length -lt 40)
				{
					WriteWordLine 0 1 "Working directory`t`t: " $Application.WorkingDirectory
				}
				Else
				{
					WriteWordLine 0 1 "Working directory: " 
					WriteWordLine 0 2 $Application.WorkingDirectory
				}
			}
			
			#servers properties
			If($AppServerInfoResults)
			{
				If(![String]::IsNullOrEmpty( $AppServerInfo.ServerNames))
				{
					WriteWordLine 0 1 "Servers:"
					ForEach($servername in $AppServerInfo.ServerNames)
					{
						WriteWordLine 0 2 $servername
					}
				}
				If(![String]::IsNullOrEmpty($AppServerInfo.WorkerGroupNames))
				{
					WriteWordLine 0 1 "Worker Groups:"
					ForEach($workergroup in $AppServerInfo.WorkerGroupNames)
					{
						WriteWordLine 0 2 $workergroup
					}
				}
			}
			Else
			{
				WriteWordLine 0 2 "Unable to retrieve a list of Servers or Worker Groups for this application"
			}
		}
	
		#users properties
		If($Application.AnonymousConnectionsAllowed)
		{
			WriteWordLine 0 1 "Allow anonymous users: " $Application.AnonymousConnectionsAllowed
		}
		Else
		{
			If($AppServerInfoResults)
			{
				WriteWordLine 0 1 "Users:"
				ForEach($user in $AppServerInfo.Accounts)
				{
					WriteWordLine 0 2 $user
				}
			}
			Else
			{
				WriteWordLine 0 2 "Unable to retrieve a list of Users for this application"
			}
		}	

		#shortcut presentation properties
		#application icon is ignored
		If(![String]::IsNullOrEmpty($Application.ClientFolder))
		{
			If($Application.ClientFolder.Length -lt 30)
			{
				WriteWordLine 0 1 "Client application folder`t`t`t`t: " $Application.ClientFolder
			}
			Else
			{
				WriteWordLine 0 1 "Client application folder`t`t`t`t: " 
				WriteWordLine 0 2 $Application.ClientFolder
			}
		}
		If($Application.AddToClientStartMenu)
		{
			WriteWordLine 0 1 "Add to client's start menu"
			If($Application.StartMenuFolder)
			{
				WriteWordLine 0 2 "Start menu folder`t`t`t: " $Application.StartMenuFolder
			}
		}
		If($Application.AddToClientDesktop)
		{
			WriteWordLine 0 1 "Add shortcut to the client's desktop"
		}
	
		#access control properties
		If($Application.ConnectionsThroughAccessGatewayAllowed)
		{
			WriteWordLine 0 1 "Allow connections made through AGAE`t`t: " -nonewline
			If($Application.ConnectionsThroughAccessGatewayAllowed)
			{
				WriteWordLine 0 0 "Yes"
			} 
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		If($Application.OtherConnectionsAllowed)
		{
			WriteWordLine 0 1 "Any connection`t`t`t`t`t: " -nonewline
			If($Application.OtherConnectionsAllowed)
			{
				WriteWordLine 0 0 "Yes"
			} 
			Else
			{
				WriteWordLine 0 0 "No"
			}
		}
		If($Application.AccessSessionConditionsEnabled)
		{
			WriteWordLine 0 1 "Any connection that meets any of the following filters: " $Application.AccessSessionConditionsEnabled
			WriteWordLine 0 1 "Access Gateway Filters:"
			ForEach($filter in $Application.AccessSessionConditions)
			{
				WriteWordLine 0 2 $filter
			}
		}
	
		#content redirection properties
		If($AppServerInfoResults)
		{
			If($AppServerInfo.FileTypes)
			{
				WriteWordLine 0 1 "File type associations:"
				ForEach($filetype in $AppServerInfo.FileTypes)
				{
					WriteWordLine 0 2 $filetype
				}
			}
			Else
			{
				WriteWordLine 0 1 "File Type Associations for this application`t: None"
			}
		}
		Else
		{
			WriteWordLine 0 1 "Unable to retrieve the list of FTAs for this application"
		}
	
		#if streamed app, Alternate profiles
		If($streamedapp)
		{
			If($Application.AlternateProfiles)
			{
				WriteWordLine 0 1 "Primary application profile location`t`t: " $Application.AlternateProfiles
			}
		
			#if streamed app, User privileges properties
			If($Application.RunAsLeastPrivilegedUser)
			{
				WriteWordLine 0 1 "Run app as a least-privileged user account`t: " $Application.RunAsLeastPrivilegedUser
			}
		}
	
		#limits properties
		WriteWordLine 0 1 "Limit instances allowed to run in server farm`t: " -NoNewLine

		If($Application.InstanceLimit -eq -1)
		{
			WriteWordLine 0 0 "No limit set"
		}
		Else
		{
			WriteWordLine 0 0 $Application.InstanceLimit
		}
	
		WriteWordLine 0 1 "Allow only 1 instance of app for each user`t: " -NoNewLine
	
		If ($Application.MultipleInstancesPerUserAllowed) 
		{
			WriteWordLine 0 0 "No"
		} 
		Else
		{
			WriteWordLine 0 0 "Yes"
		}
	
		If($Application.CpuPriorityLevel)
		{
			WriteWordLine 0 1 "Application importance`t`t`t`t: " -nonewline
			switch ($Application.CpuPriorityLevel)
			{
				"Unknown"     {WriteWordLine 0 0 "Unknown"}
				"BelowNormal" {WriteWordLine 0 0 "Below Normal"}
				"Low"         {WriteWordLine 0 0 "Low"}
				"Normal"      {WriteWordLine 0 0 "Normal"}
				"AboveNormal" {WriteWordLine 0 0 "Above Normal"}
				"High"        {WriteWordLine 0 0 "High"}
				Default {WriteWordLine 0 0 "Application importance could not be determined: $($Application.CpuPriorityLevel)"}
			}
		}
		
		#client options properties
		WriteWordLine 0 1 "Enable legacy audio`t`t`t`t: " -nonewline
		switch ($Application.AudioType)
		{
			"Unknown" {WriteWordLine 0 0 "Unknown"}
			"None"    {WriteWordLine 0 0 "Not Enabled"}
			"Basic"   {WriteWordLine 0 0 "Enabled"}
			Default {WriteWordLine 0 0 "Enable legacy audio could not be determined: $($Application.AudioType)"}
		}
		WriteWordLine 0 1 "Minimum requirement`t`t`t`t: " -nonewline
		If($Application.AudioRequired)
		{
			WriteWordLine 0 0 "Enabled"
		}
		Else
		{
			WriteWordLine 0 0 "Disabled"
		}
		If($Application.SslConnectionEnabled)
		{
			WriteWordLine 0 1 "Enable SSL and TLS protocols`t`t`t: " -nonewline
			If($Application.SslConnectionEnabled)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
		}
		If($Application.EncryptionLevel)
		{
			WriteWordLine 0 1 "Encryption`t`t`t`t`t: " -nonewline
			switch ($Application.EncryptionLevel)
			{
				"Unknown" {WriteWordLine 0 0 "Unknown"}
				"Basic"   {WriteWordLine 0 0 "Basic"}
				"LogOn"   {WriteWordLine 0 0 "128-Bit Login Only (RC-5)"}
				"Bits40"  {WriteWordLine 0 0 "40-Bit (RC-5)"}
				"Bits56"  {WriteWordLine 0 0 "56-Bit (RC-5)"}
				"Bits128" {WriteWordLine 0 0 "128-Bit (RC-5)"}
				Default {WriteWordLine 0 0 "Encryption could not be determined: $($Application.EncryptionLevel)"}
			}
		}
		If($Application.EncryptionRequired)
		{
			WriteWordLine 0 1 "Minimum requirement`t`t`t`t: " -nonewline
			If($Application.EncryptionRequired)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
		}
	
		WriteWordLine 0 1 "Start app w/o waiting for printer creation`t: " -NoNewLine
		#another weird one, if True then this is Disabled
		If ($Application.WaitOnPrinterCreation) 
		{
			WriteWordLine 0 0 "Disabled"
		} 
		Else
		{
			WriteWordLine 0 0 "Enabled"
		}
		
		#appearance properties
		If($Application.WindowType)
		{
			WriteWordLine 0 1 "Session window size`t`t`t`t: " $Application.WindowType
		}
		If($Application.ColorDepth)
		{
			WriteWordLine 0 1 "Maximum color quality`t`t`t`t: " -nonewline
			switch ($Application.ColorDepth)
			{
				"Unknown"     {WriteWordLine 0 0 "Unknown color depth"}
				"Colors8Bit"  {WriteWordLine 0 0 "256-color (8-bit)"}
				"Colors16Bit" {WriteWordLine 0 0 "Better Speed (16-bit)"}
				"Colors32Bit" {WriteWordLine 0 0 "Better Appearance (32-bit)"}
				Default {WriteWordLine 0 0 "Maximum color quality could not be determined: $($Application.ColorDepth)"}
			}
		}
		If($Application.TitleBarHidden)
		{
			WriteWordLine 0 1 "Hide application title bar`t`t`t: " -nonewline
			If($Application.TitleBarHidden)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
		}
		If($Application.MaximizedOnStartup)
		{
			WriteWordLine 0 1 "Maximize application at startup`t`t`t: " -nonewline
			If($Application.MaximizedOnStartup)
			{
				WriteWordLine 0 0 "Enabled"
			}
			Else
			{
				WriteWordLine 0 0 "Disabled"
			}
		}
	$AppServerInfo = $null
	}
}
Else 
{
	write-warning "Application information could not be retrieved"
}

$Applications = $null

write-verbose "Processing Configuration Logging/History Report"
If( $Global:ConfigLog )
{
	#history AKA Configuration Logging report
	#only process if $Global:ConfigLog = $True and .\XA65ConfigLog.udl file exists
	#build connection string
	#User ID is account that has access permission for the configuration logging database
	#Initial Catalog is the name of the Configuration Logging SQL Database
	If ( Test-Path .\XA65ConfigLog.udl )
	{
		$ConnectionString = Get-Content .\xa65configlog.udl | select-object -last 1
		$ConfigLogReport = get-CtxConfigurationLogReport -connectionstring $ConnectionString -EA 0

		If( $? -and $ConfigLogReport)
		{
			$selection.InsertNewPage()
			WriteWordLine 1 0 "History:"
			ForEach($ConfigLogItem in $ConfigLogReport)
			{
				WriteWordLine 0 1 "Date`t`t`t: " $ConfigLogItem.Date
				WriteWordLine 0 1 "Account`t`t: " $ConfigLogItem.Account
				WriteWordLine 0 1 "Change description`t: " $ConfigLogItem.Description
				WriteWordLine 0 1 "Type of change`t`t: " $ConfigLogItem.TaskType
				WriteWordLine 0 1 "Type of item`t`t: " $ConfigLogItem.ItemType
				WriteWordLine 0 1 "Name of item`t`t: " $ConfigLogItem.ItemName
				WriteWordLine 0 0 ""
			}
		} 
		Else 
		{
			WriteWordLine 0 0 "History information could not be retrieved"
		}
		$ConnectionString = $null
		$ConfigLogReport = $null
		$global:output = $null
	}
	Else 
	{
		WriteWordLine 1 0 "Configuration Logging is enabled but the XA65ConfigLog.udl file was not found"
	}
}

#load balancing policies
write-verbose "Processing Load Balancing Policies"
$LoadBalancingPolicies = Get-XALoadBalancingPolicy -EA 0 | sort-object PolicyName

If( $? -and $LoadBalancingPolicies)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Load Balancing Policies:"
	ForEach($LoadBalancingPolicy in $LoadBalancingPolicies)
	{
		$LoadBalancingPolicyConfiguration = Get-XALoadBalancingPolicyConfiguration -PolicyName $LoadBalancingPolicy.PolicyName
		$LoadBalancingPolicyFilter = Get-XALoadBalancingPolicyFilter -PolicyName $LoadBalancingPolicy.PolicyName 
	
		WriteWordLine 2 0 $LoadBalancingPolicy.PolicyName
		WriteWordLine 0 1 "Description`t: " $LoadBalancingPolicy.Description
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
				If($LoadBalancingPolicyFilter.AllowOtherConnections)
				{
					WriteWordLine 0 2 "Any connection"
				} 
				Else
				{
					WriteWordLine 0 2 "Any connection that meets any of the following filters"
					If($LoadBalancingPolicyFilter.AccessSessionConditions)
					{
						ForEach($AccessSessionCondition in $LoadBalancingPolicyFilter.AccessSessionConditions)
						{
							WriteWordLine 4 $AccessSessionCondition
						}
					}
				}
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
					WriteWordLine 0 2 "Worker Group: " $WorkerGroupPreference
				}
			}
		}
		If($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Enabled")
		{
			WriteWordLine 0 1 "Set the delivery protocols for applications streamed to client"
			WriteWordLine 0 2 "" -nonewline
			switch ($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)
			{
				"Unknown"                {WriteWordLine 0 0 "Unknown"}
				"ForceServerAccess"      {WriteWordLine 0 0 "Do not allow applications to stream to the client"}
				"ForcedStreamedDelivery" {WriteWordLine 0 0 "Force applications to stream to the client"}
				Default {WriteWordLine 0 0 "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"}
			}
		}
		Elseif($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Disabled")
		{
			#In the GUI, if "Set the delivery protocols for applications streamed to client" IS selected AND 
			#"Allow applications to stream to the client or run on a Terminal Server (default)" IS selected
			#then "Set the delivery protocols for applications streamed to client" is set to Disabled
			WriteWordLine 0 1 "Set the delivery protocols for applications streamed to client"
			WriteWordLine 0 2 "Allow applications to stream to the client or run on a Terminal Server (default)"
		}
		Else
		{
			WriteWordLine 0 1 "Streamed App Delivery is not configured"
		}
	
		$LoadBalancingPolicyConfiguration = $null
		$LoadBalancingPolicyFilter = $null
	}
}
Else 
{
	Write-warning "Load balancing policy information could not be retrieved"
}
$LoadBalancingPolicies = $null

#load evaluators
write-verbose "Processing Load Evaluators"
$LoadEvaluators = Get-XALoadEvaluator -EA 0 | sort-object LoadEvaluatorName

If( $? )
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Load Evaluators:"
	ForEach($LoadEvaluator in $LoadEvaluators)
	{
		WriteWordLine 2 0 $LoadEvaluator.LoadEvaluatorName
		WriteWordLine 0 1 "Description: " $LoadEvaluator.Description
		
		If($LoadEvaluator.IsBuiltIn)
		{
			WriteWordLine 0 1 "Built-in Load Evaluator"
		} 
		Else 
		{
			WriteWordLine 0 1 "User created load evaluator"
		}
	
		If($LoadEvaluator.ApplicationUserLoadEnabled)
		{
			WriteWordLine 0 1 "Application User Load Settings"
			WriteWordLine 0 2 "Report full load when the # of users for this application equals: " $LoadEvaluator.ApplicationUserLoad
			WriteWordLine 0 2 "Application: " $LoadEvaluator.ApplicationBrowserName
		}
	
		If($LoadEvaluator.ContextSwitchesEnabled)
		{
			WriteWordLine 0 1 "Context Switches Settings"
			WriteWordLine 0 2 "Report full load when the # of context switches per second is > than: " $LoadEvaluator.ContextSwitches[1]
			WriteWordLine 0 2 "Report no load when the # of context switches per second is <= to: " $LoadEvaluator.ContextSwitches[0]
		}
	
		If($LoadEvaluator.CpuUtilizationEnabled)
		{
			WriteWordLine 0 1 "CPU Utilization Settings"
			WriteWordLine 0 2 "Report full load when the processor utilization % is > than: " $LoadEvaluator.CpuUtilization[1]
			WriteWordLine 0 2 "Report no load when the processor utilization % is <= to: " $LoadEvaluator.CpuUtilization[0]
		}
	
		If($LoadEvaluator.DiskDataIOEnabled)
		{
			WriteWordLine 0 1 "Disk Data I/O Settings"
			WriteWordLine 0 2 "Report full load when the total disk I/O in kbps is > than: " $LoadEvaluator.DiskDataIO[1]
			WriteWordLine 0 2 "Report no load when the total disk I/O in kbps per second is <= to: " $LoadEvaluator.DiskDataIO[0]
		}
	
		If($LoadEvaluator.DiskOperationsEnabled)
		{
			WriteWordLine 0 1 "Disk Operations Settings"
			WriteWordLine 0 2 "Report full load when the total # of R/W operations per second is > than: " $LoadEvaluator.DiskOperations[1]
			WriteWordLine 0 2 "Report no load when the total # of R/W operations per second is <= to: " $LoadEvaluator.DiskOperations[0]
		}
	
		If($LoadEvaluator.IPRangesEnabled)
		{
			WriteWordLine 0 1 "IP Range Settings"
			If($LoadEvaluator.IPRangesAllowed)
			{
				WriteWordLine 0 2 "Allow " -NoNewLine
			} 
			Else 
			{
				WriteWordLine 0 2 "Deny " -NoNewLine
			}
			WriteWordLine 0 0 "client connections from the listed IP Ranges"
			ForEach($IPRange in $LoadEvaluator.IPRanges)
			{
				WriteWordLine 0 3 "IP Address Ranges: " $IPRange
			}
		}
	
		If($LoadEvaluator.LoadThrottlingEnabled)
		{
			WriteWordLine 0 1 "Load Throttling Settings"
			WriteWordLine 0 2 "Impact of logons on load: " -nonewline
			switch ($LoadEvaluator.LoadThrottling)
			{
				"Unknown"    {WriteWordLine 0 0 "Unknown"}
				"Extreme"    {WriteWordLine 0 0 "Extreme"}
				"High"       {WriteWordLine 0 0 "High (Default)"}
				"MediumHigh" {WriteWordLine 0 0 "Medium High"}
				"Medium"     {WriteWordLine 0 0 "Medium"}
				"MediumLow"  {WriteWordLine 0 0 "Medium Low"}
				Default {WriteWordLine 0 0 "Impact of logons on load could not be determined: $($LoadEvaluator.LoadThrottling)"}
			}
		}
	
		If($LoadEvaluator.MemoryUsageEnabled)
		{
			WriteWordLine 0 1 "Memory Usage Settings"
			WriteWordLine 0 2 "Report full load when the memory usage is > than: " $LoadEvaluator.MemoryUsage[1]
			WriteWordLine 0 2 "Report no load when the memory usage is <= to: " $LoadEvaluator.MemoryUsage[0]
		}
	
		If($LoadEvaluator.PageFaultsEnabled)
		{
			WriteWordLine 0 1 "Page Faults Settings"
			WriteWordLine 0 2 "Report full load when the # of page faults per second is > than: " $LoadEvaluator.PageFaults[1]
			WriteWordLine 0 2 "Report no load when the # of page faults per second is <= to: " $LoadEvaluator.PageFaults[0]
		}
	
		If($LoadEvaluator.PageSwapsEnabled)
		{
			WriteWordLine 0 1 "Page Swaps Settings"
			WriteWordLine 0 2 "Report full load when the # of page swaps per second is > than: " $LoadEvaluator.PageSwaps[1]
			WriteWordLine 0 2 "Report no load when the # of page swaps per second is <= to: " $LoadEvaluator.PageSwaps[0]
		}
	
		If($LoadEvaluator.ScheduleEnabled)
		{
			WriteWordLine 0 1 "Scheduling Settings"
			WriteWordLine 0 2 "Sunday Schedule`t: " $LoadEvaluator.SundaySchedule
			WriteWordLine 0 2 "Monday Schedule`t: " $LoadEvaluator.MondaySchedule
			WriteWordLine 0 2 "Tuesday Schedule`t: " $LoadEvaluator.TuesdaySchedule
			WriteWordLine 0 2 "Wednesday Schedule`t: " $LoadEvaluator.WednesdaySchedule
			WriteWordLine 0 2 "Thursday Schedule`t: " $LoadEvaluator.ThursdaySchedule
			WriteWordLine 0 2 "Friday Schedule`t`t: " $LoadEvaluator.FridaySchedule
			WriteWordLine 0 2 "Saturday Schedule`t: " $LoadEvaluator.SaturdaySchedule
		}
	
		If($LoadEvaluator.ServerUserLoadEnabled)
		{
			WriteWordLine 0 1 "Server User Load Settings"
			WriteWordLine 0 2 "Report full load when the # of server users equals: " $LoadEvaluator.ServerUserLoad
		}
	
		WriteWordLine 0 0 ""
	}
}
Else 
{
	Write-warning "Load Evaluator information could not be retrieved"
}
$LoadEvaluators = $null

#servers
write-verbose "Processing Servers"
$servers = Get-XAServer -EA 0 | sort-object FolderPath, ServerName

If( $? )
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Servers:"
	ForEach($server in $servers)
	{
		WriteWordLine 2 0 $server.ServerName
		WriteWordLine 0 1 "Product`t`t`t`t: " $server.CitrixProductName
		WriteWordLine 0 1 "Edition`t`t`t`t: " $server.CitrixEdition
		WriteWordLine 0 1 "Version`t`t`t`t: " $server.CitrixVersion
		WriteWordLine 0 1 "Service Pack`t`t`t: " $server.CitrixServicePack
		WriteWordLine 0 1 "IP Address`t`t`t: " $server.IPAddresses
		WriteWordLine 0 1 "Logons`t`t`t`t: " -NoNewLine
		If($server.LogOnsEnabled)
		{
			WriteWordLine 0 0 "Enabled"
		} 
		Else 
		{
			WriteWordLine 0 0 "Disabled"
		}
		WriteWordLine 0 1 "Logon Control Mode`t`t: " -nonewline
		switch ($Server.LogOnMode)
		{
			"Unknown"                       {WriteWordLine 0 0 "Unknown"}
			"AllowLogOns"                   {WriteWordLine 0 0 "Allow logons and reconnections"}
			"ProhibitNewLogOnsUntilRestart" {WriteWordLine 0 0 "Prohibit logons until server restart"}
			"ProhibitNewLogOns "            {WriteWordLine 0 0 "Prohibit logons only"}
			"ProhibitLogOns "               {WriteWordLine 0 0 "Prohibit logons and reconnections"}
			Default {WriteWordLine 0 0 "Logon control mode could not be determined: $($Server.LogOnMode)"}
		}

		WriteWordLine 0 1 "Product Installation Date`t: " $server.CitrixInstallDate
		WriteWordLine 0 1 "Operating System Version`t: " $server.OSVersion -NoNewLine
		WriteWordLine 0 0 " " $server.OSServicePack
		WriteWordLine 0 1 "Zone`t`t`t`t: " $server.ZoneName
		WriteWordLine 0 1 "Election Preference`t`t: " -nonewline
		switch ($server.ElectionPreference)
		{
			"Unknown"           {WriteWordLine 0 0 "Unknown"}
			"MostPreferred"     {WriteWordLine 0 0 "Most Preferred"}
			"Preferred"         {WriteWordLine 0 0 "Preferred"}
			"DefaultPreference" {WriteWordLine 0 0 "Default Preference"}
			"NotPreferred"      {WriteWordLine 0 0 "Not Preferred"}
			"WorkerMode"        {WriteWordLine 0 0 "Worker Mode"}
			Default {WriteWordLine 0 0 "Server election preference could not be determined: $($server.ElectionPreference)"}
		}
		WriteWordLine 0 1 "Folder`t`t`t`t: " $server.FolderPath
		WriteWordLine 0 1 "Product Installation Path`t: " $server.CitrixInstallPath
		If($server.LicenseServerName)
		{
			WriteWordLine 0 1 "License Server Name`t`t: " $server.LicenseServerName
			WriteWordLine 0 1 "License Server Port`t`t: " $server.LicenseServerPortNumber
		}
		If($server.ICAPortNumber -gt 0)
		{
			WriteWordLine 0 1 "ICA Port Number`t`t: " $server.ICAPortNumber
		}
		
		WriteWordLine 0 0 ""
		
		#applications published to server
		$Applications = Get-XAApplication -ServerName $server.ServerName -EA 0 | sort-object FolderPath, DisplayName
		If( $? -and $Applications )
		{
			WriteWordLine 0 1 "Published applications:"
			ForEach($app in $Applications)
			{
				WriteWordLine 0 2 "Display name`t: " $app.DisplayName
				WriteWordLine 0 2 "Folder path`t: " $app.FolderPath
				WriteWordLine 0 0 ""
			}
		}
		#list citrix services
		Write-Verbose "`t`tTesting to see if $($server.ServerName) is online and reachable"
		If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
		{
			Write-Verbose "`t`t$($server.ServerName) is online.  Citrix Services and Hotfix areas processed."
			Write-Verbose "`t`tProcessing Citrix services for server $($server.ServerName)"
			$services = get-service -ComputerName $server.ServerName -EA 0 | where-object {$_.DisplayName -like "*Citrix*"} | sort-object DisplayName
			WriteWordLine 0 1 "Citrix Services"
			Write-Verbose "`t`tCreate Word Table for Citrix services"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			[int]$Rows = $services.count + 1
			Write-Verbose "`t`tadd Citrix services table to doc"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = "Table Grid"
			$table.Borders.InsideLineStyle = 1
			$table.Borders.OutsideLineStyle = 1
			[int]$xRow = 1
			Write-Verbose "`t`tformat first row with column headings"
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Display Name"
			$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,2).Range.Font.Bold = $True
			$Table.Cell($xRow,2).Range.Text = "Status"
			ForEach($Service in $Services)
			{
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $Service.DisplayName
				If($Service.Status -eq "Stopped")
				{
					$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
					$Table.Cell($xRow,2).Range.Font.Bold  = $True
					$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
				}
				$Table.Cell($xRow,2).Range.Text = $Service.Status
			}

			Write-Verbose "`t`tMove table of Citrix services to the right"
			$Table.Rows.SetLeftIndent(43,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			Write-Verbose "`t`treturn focus back to document"
			$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

			#move to the end of the current document
			Write-Verbose "`t`tmove to the end of the current document"
			$selection.EndKey($wdStory,$wdMove) | Out-Null

			#Citrix hotfixes installed
			Write-Verbose "`t`tGet list of Citrix hotfixes installed"
			$hotfixes = Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | sort-object HotfixName
			If( $? -and $hotfixes )
			{
				$Rows = 1
				$Single_Row = (Get-Member -Type Property -Name Length -InputObject $hotfixes -EA 0) -eq $null
				If(-not $Single_Row)
				{
					$Rows = $Hotfixes.length
				}
				$Rows++
				
				Write-Verbose "`t`tnumber of hotfixes is $($Rows-1)"
				$HotfixArray = ""
				$HRP1Installed = $False
				WriteWordLine 0 0 ""
				WriteWordLine 0 1 "Citrix Installed Hotfixes:"
				Write-Verbose "`t`tCreate Word Table for Citrix Hotfixes"
				$TableRange = $doc.Application.Selection.Range
				$Columns = 5
				Write-Verbose "`t`tadd Citrix installed hotfix table to doc"
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = 1
				$table.Borders.OutsideLineStyle = 1
				$xRow = 1
				Write-Verbose "`t`tformat first row with column headings"
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Font.Size = "10"
				$Table.Cell($xRow,1).Range.Text = "Hotfix"
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,2).Range.Font.Bold = $True
				$Table.Cell($xRow,2).Range.Font.Size = "10"
				$Table.Cell($xRow,2).Range.Text = "Installed By"
				$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,3).Range.Font.Bold = $True
				$Table.Cell($xRow,3).Range.Font.Size = "10"
				$Table.Cell($xRow,3).Range.Text = "Install Date"
				$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,4).Range.Font.Bold = $True
				$Table.Cell($xRow,4).Range.Font.Size = "10"
				$Table.Cell($xRow,4).Range.Text = "Type"
				$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,5).Range.Font.Bold = $True
				$Table.Cell($xRow,5).Range.Font.Size = "10"
				$Table.Cell($xRow,5).Range.Text = "Valid"
				ForEach($hotfix in $hotfixes)
				{
					$xRow++
					$HotfixArray += $hotfix.HotfixName
					If( $hotfix.HotfixName -eq "XA650W2K8R2X64R01")
					{
						$HRP1Installed = $True
					}
					$InstallDate = $hotfix.InstalledOn.ToString()
					
					$Table.Cell($xRow,1).Range.Font.Size = "10"
					$Table.Cell($xRow,1).Range.Text = $hotfix.HotfixName
					$Table.Cell($xRow,2).Range.Font.Size = "10"
					$Table.Cell($xRow,2).Range.Text = $hotfix.InstalledBy
					$Table.Cell($xRow,3).Range.Font.Size = "10"
					$Table.Cell($xRow,3).Range.Text = $InstallDate.SubString(0,$InstallDate.IndexOf(" "))
					$Table.Cell($xRow,4).Range.Font.Size = "10"
					$Table.Cell($xRow,4).Range.Text = $hotfix.HotfixType
					$Table.Cell($xRow,5).Range.Font.Size = "10"
					$Table.Cell($xRow,5).Range.Text = $hotfix.Valid
				}
				Write-Verbose "`t`tMove table of Citrix installed hotfixes to the right"
				$Table.Rows.SetLeftIndent(43,1)
				$table.AutoFitBehavior(1)

				#return focus back to document
				Write-Verbose "`t`treturn focus back to document"
				$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

				#move to the end of the current document
				Write-Verbose "`t`tmove to the end of the current document"
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				WriteWordLine 0 0 ""

				#compare Citrix hotfixes to recommended Citrix hotfixes from CTX129229
				#hotfix lists are from CTX129229 dated 29-MAY-2013
				Write-Verbose "`t`tcompare Citrix hotfixes to recommended Citrix hotfixes from CTX129229"
				# as of the 29-apr-2013 update, there are recommended hotfixes for pre and post R01
				Write-Verbose "`t`tProcessing Citrix hotfix list for server $($server.ServerName)"
				WriteWordLine 0 1 "Citrix Recommended Hotfixes:"
				If( !$HRP1Installed )
				{
					$RecommendedList = @("XA650W2K8R2X64001","XA650W2K8R2X64011","XA650W2K8R2X64019","XA650W2K8R2X64025")
				}
				Else
				{
					$RecommendedList = @("XA650R01W2K8R2X64061")
				}
				Write-Verbose "`t`tCreate Word Table for Citrix Hotfixes"
				$TableRange = $doc.Application.Selection.Range
				$Columns = 2
				$Rows = $RecommendedList.count + 1
				Write-Verbose "`t`tadd Citrix recommended hotfix table to doc"
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = 1
				$table.Borders.OutsideLineStyle = 1
				$xRow = 1
				Write-Verbose "`t`tformat first row with column headings"
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Citrix Hotfix"
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,2).Range.Font.Bold = $True
				$Table.Cell($xRow,2).Range.Text = "Status"
				ForEach($element in $RecommendedList)
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = $element
					If(!$HotfixArray -contains $element)
					{
						#missing a recommended Citrix hotfix
						#WriteWordLine 0 2 "Recommended Citrix Hotfix $element is not installed"
						$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
						$Table.Cell($xRow,2).Range.Font.Bold  = $True
						$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
						$Table.Cell($xRow,2).Range.Text = "Not Installed"
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = "Installed"
					}
				}
				Write-Verbose "`t`tMove table of Citrix hotfixes to the right"
				$Table.Rows.SetLeftIndent(43,1)
				$table.AutoFitBehavior(1)

				#return focus back to document
				Write-Verbose "`t`treturn focus back to document"
				$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

				#move to the end of the current document
				Write-Verbose "`t`tmove to the end of the current document"
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				WriteWordLine 0 0 ""
				#build list of installed Microsoft hotfixes
				Write-Verbose "`t`tProcessing Microsoft hotfixes for server $($server.ServerName)"
				$MSInstalledHotfixes = Get-HotFix -computername $Server.ServerName -EA 0 | select-object -Expand HotFixID | sort-object HotFixID
				If($server.OSServicePack.IndexOf('1') -gt 0)
				{
					#Server 2008 R2 SP1 installed
					$RecommendedList = @("KB2444328", "KB2465772", "KB2551503", "KB2571388", 
										"KB2578159", "KB2617858", "KB2620656", "KB2647753",
										"KB2661001", "KB2661332", "KB2731847", "KB2748302",
										"KB2775511", "KB2778831", "KB917607")
				}
				Else
				{
					#Server 2008 R2 without SP1 installed
					$RecommendedList = @("KB2265716", "KB2388142", "KB2383928", "KB2444328", 
										"KB2465772", "KB2551503", "KB2571388", "KB2578159", 
										"KB2617858", "KB2620656", "KB2647753", "KB2661001",
										"KB2661332", "KB2731847", "KB2748302", "KB2778831", "KB917607", 
										"KB975777", "KB979530", "KB980663", "KB983460")
				}
				
				WriteWordLine 0 1 "Microsoft Recommended Hotfixes:"
				Write-Verbose "`t`tCreate Word Table for Microsoft Hotfixes"
				$TableRange = $doc.Application.Selection.Range
				$Columns = 2
				$Rows = $RecommendedList.count + 1
				Write-Verbose "`t`tadd Microsoft hotfix table to doc"
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$table.Style = "Table Grid"
				$table.Borders.InsideLineStyle = 1
				$table.Borders.OutsideLineStyle = 1
				$xRow = 1
				Write-Verbose "`t`tformat first row with column headings"
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Microsoft Hotfix"
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,2).Range.Font.Bold = $True
				$Table.Cell($xRow,2).Range.Text = "Status"

				$results = @{}
				foreach( $hotfix in $RecommendedList )
				{
					$xRow++
					$Table.Cell($xRow,1).Range.Text = $hotfix
					If(!($MSInstalledHotfixes -contains $hotfix))
					{
						$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorRed
						$Table.Cell($xRow,2).Range.Font.Bold  = $True
						$Table.Cell($xRow,2).Range.Font.Color = $WDColorBlack
						$Table.Cell($xRow,2).Range.Text = "Not Installed"
					}
					Else
					{
						$Table.Cell($xRow,2).Range.Text = "Installed"
					}
				}
				Write-Verbose "`t`tMove table of Microsoft hotfixes to the right"
				$Table.Rows.SetLeftIndent(43,1)
				$table.AutoFitBehavior(1)

				#return focus back to document
				Write-Verbose "`t`treturn focus back to document"
				$doc.ActiveWindow.ActivePane.view.SeekView=$wdSeekMainDocument

				#move to the end of the current document
				Write-Verbose "`t`tmove to the end of the current document"
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				WriteWordLine 0 1 "Not all missing Microsoft hotfixes may be needed for this server"
			}
		}
		Else
		{
			Write-Verbose "`t`t$($server.ServerName) is offline or unreachable.  Citrix Services and Hotfix areas skipped."
			WriteWordLine 0 0 "Server $($server.ServerName) was offline or unreachable at "(get-date).ToString()
			WriteWordLine 0 0 "The Citrix Services and Hotfix areas were skipped."
		}
		WriteWordLine 0 0 "" 
	}
}
Else 
{
	Write-warning "Server information could not be retrieved"
}
$servers = $null

#worker groups
write-verbose "Processing Worker Groups"
$WorkerGroups = Get-XAWorkerGroup -EA 0 | sort-object WorkerGroupName

If( $? -and $WorkerGroups)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Worker Groups:"
	ForEach($WorkerGroup in $WorkerGroups)
	{
		WriteWordLine 2 0 $WorkerGroup.WorkerGroupName
		WriteWordLine 0 1 "Description: " $WorkerGroup.Description
		WriteWordLine 0 1 "Folder Path: " $WorkerGroup.FolderPath
		If($WorkerGroup.ServerNames)
		{
			WriteWordLine 0 1 "Farm Servers:"
			$TempArray = $WorkerGroup.ServerNames | Sort-Object
			ForEach($ServerName in $TempArray)
			{
				WriteWordLine 0 2 $ServerName
			}
			$TempArray = $null
		}
		If($WorkerGroup.ServerGroups)
		{
			WriteWordLine 0 1 "Server Group Accounts:"
			$TempArray = $WorkerGroup.ServerGroups | Sort-Object
			ForEach($ServerGroup in $TempArray)
			{
				WriteWordLine 0 2 $ServerGroup
			}
			$TempArray = $null
		}
		If($WorkerGroup.OUs)
		{
			WriteWordLine 0 1 "Organizational Units:"
			$TempArray = $WorkerGroup.OUs | Sort-Object
			ForEach($OU in $TempArray)
			{
				WriteWordLine 0 2 $OU
			}
			$TempArray = $null
		}
		#applications published to worker group
		$Applications = Get-XAApplication -WorkerGroup $WorkerGroup.WorkerGroupName -EA 0 | sort-object FolderPath, DisplayName
		If( $? -and $Applications )
		{
			WriteWordLine 0 1 "Published applications:"
			ForEach($app in $Applications)
			{
				WriteWordLine 0 2 "Display name: " $app.DisplayName
				WriteWordLine 0 2 "Folder path: " $app.FolderPath
				WriteWordLine 0 0 ""
			}
		}
		WriteWordLine 0 0 ""
	}
}
Else 
{
	Write-warning "Worker Group information could not be retrieved"
}
$WorkerGroups = $null

#zones
write-verbose "Processing Zones"
$Zones = Get-XAZone -EA 0 | sort-object ZoneName
If( $? )
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Zones:"
	ForEach($Zone in $Zones)
	{
		WriteWordLine 2 0 $Zone.ZoneName
		WriteWordLine 0 1 "Current Data Collector: " $Zone.DataCollector
		$Servers = Get-XAServer -ZoneName $Zone.ZoneName -EA 0 | sort-object ElectionPreference, ServerName
		If( $? )
		{		
			WriteWordLine 0 1 "Servers in Zone"
	
			ForEach($Server in $Servers)
			{
				WriteWordLine 0 2 "Server Name and Preference: " $server.ServerName -NoNewLine
				WriteWordLine 0 0  " - " -nonewline
				switch ($server.ElectionPreference)
				{
					"Unknown"           {WriteWordLine 0 0 "Unknown"}
					"MostPreferred"     {WriteWordLine 0 0 "Most Preferred"}
					"Preferred"         {WriteWordLine 0 0 "Preferred"}
					"DefaultPreference" {WriteWordLine 0 0 "Default Preference"}
					"NotPreferred"      {WriteWordLine 0 0 "Not Preferred"}
					"WorkerMode"        {WriteWordLine 0 0 "Worker Mode"}
					Default {WriteWordLine 0 0 "Zone preference could not be determined: $($server.ElectionPreference)"}
				}
			}
		}
		Else
		{
			WriteWordLine 0 1 "Unable to enumerate servers in the zone"
		}
		$Servers = $Null
	}
}
Else 
{
	Write-warning "Zone information could not be retrieved"
}
$Servers = $null
$Zones = $null

#if remoting is enabled, the citrix.grouppolicy.commands module does not work with remoting so skip it
If($Remoting)
{
	write-warning "Remoting is enabled."
	write-warning "The Citrix.GroupPolicy.Commands module does not work with Remoting."
	write-warning "Citrix Policy documentation will not take place."
}
Else
{
	#make sure Citrix.GroupPolicy.Commands module is loaded
	If(!(Check-LoadedModule "Citrix.GroupPolicy.Commands"))
	{
		write-warning "The Citrix Group Policy module Citrix.GroupPolicy.Commands.psm1 does not exist (http://support.citrix.com/article/CTX128625), Citrix Policy documentation will not take place."
	}
	else
	{
		write-verbose "Processing Citrix IMA Policies"
		$Policies = Get-CtxGroupPolicy -EA 0 | sort-object Type,Priority
		If( $? )
		{
			$selection.InsertNewPage()
			WriteWordLine 1 0 "Policies:"
			ForEach($Policy in $Policies)
			{
				write-verbose "`t$($Policy.PolicyName)`t$($Policy.Type)"
				WriteWordLine 2 0 $Policy.PolicyName
				WriteWordLine 0 1 "Type`t`t: " $Policy.Type
				If(![String]::IsNullOrEmpty($Policy.Description))
				{
					WriteWordLine 0 1 "Description`t: " $Policy.Description
				}
				WriteWordLine 0 1 "Enabled`t`t: " $Policy.Enabled
				WriteWordLine 0 1 "Priority`t`t: " $Policy.Priority

				$filters = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName -EA 0

				If( $? )
				{
					If(![String]::IsNullOrEmpty($filters))
					{
						WriteWordLine 0 1 "Filter(s)`t`t:"
						ForEach($Filter in $Filters)
						{
							WriteWordLine 0 2 "Filter name`t: " $filter.FilterName
							WriteWordLine 0 2 "Filter type`t: " $filter.FilterType
							WriteWordLine 0 2 "Filter enabled`t: " $filter.Enabled
							WriteWordLine 0 2 "Filter mode`t: " $filter.Mode
							WriteWordLine 0 2 "Filter value`t: " $filter.FilterValue
							WriteWordLine 0 2 ""
						}
					}
					Else
					{
						WriteWordLine 0 1 "Filter(s)`t`t: None"
						#WriteWordLine 0 1 "No filter information"
					}
				}
				Else
				{
					WriteWordLine 0 1 "Unable to retrieve Filter settings"
				}

				$Settings = Get-CtxGroupPolicyConfiguration -PolicyName $Policy.PolicyName -EA 0
				If( $? )
				{
					ForEach($Setting in $Settings)
					{
						If($Setting.Type -eq "Computer")
						{
							WriteWordLine 0 1 "Computer settings:"
							If($Setting.IcaListenerTimeout.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\ICA listener connection timeout (milliseconds): " $Setting.IcaListenerTimeout.Value
							}
							If($Setting.IcaListenerPortNumber.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\ICA listener port number: " $Setting.IcaListenerPortNumber.Value
							}
							If($Setting.AutoClientReconnect.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Auto Client Reconnect\Auto client reconnect: " $Setting.AutoClientReconnect.State
							}
							If($Setting.AutoClientReconnectLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Auto Client Reconnect\Auto client reconnect logging: "
								switch ($Setting.AutoClientReconnectLogging.Value)
								{
									"DoNotLogAutoReconnectEvents" {WriteWordLine 0 3 "Do Not Log auto-reconnect events"}
									"LogAutoReconnectEvents"      {WriteWordLine 0 3 "Log auto-reconnect events"}
									Default {WriteWordLine 0 3 "Auto client reconnect logging could not be determined: $($Setting.AutoClientReconnectLogging.Value)"}
								}
							}
							If($Setting.IcaRoundTripCalculation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\End User Monitoring\ICA round trip calculation: " $Setting.IcaRoundTripCalculation.State
							}
							If($Setting.IcaRoundTripCalculationInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\End User Monitoring\ICA round trip calculation interval (seconds): " $Setting.IcaRoundTripCalculationInterval.Value
							}
							If($Setting.IcaRoundTripCalculationWhenIdle.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\End User Monitoring\ICA round trip calculations for idle connections: " 
								WriteWordLine 0 3 $Setting.IcaRoundTripCalculationWhenIdle.State
							}
							If($Setting.DisplayMemoryLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Display memory limit (KB): " $Setting.DisplayMemoryLimit.Value
							}
							If($Setting.DisplayDegradePreference.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Display mode degrade preference: "
								
								switch ($Setting.DisplayDegradePreference.Value)
								{
									"ColorDepth" {WriteWordLine 0 3 "Degrade color depth first"}
									"Resolution" {WriteWordLine 0 3 "Degrade resolution first"}
									Default {WriteWordLine 0 3 "Display mode degrade preference could not be determined: $($Setting.DisplayDegradePreference.Value)"}
								}
							}
							If($Setting.DynamicPreview.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Dynamic Windows Preview: " $Setting.DynamicPreview.State
							}
							If($Setting.ImageCaching.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Image caching: " $Setting.ImageCaching.State
							}
							If($Setting.MaximumColorDepth.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Maximum allowed color depth: "
								switch ($Setting.MaximumColorDepth.Value)
								{
									"BitsPerPixel8"  {WriteWordLine 0 3 "8 Bits Per Pixel"}
									"BitsPerPixel15" {WriteWordLine 0 3 "15 Bits Per Pixel"}
									"BitsPerPixel16" {WriteWordLine 0 3 "16 Bits Per Pixel"}
									"BitsPerPixel24" {WriteWordLine 0 3 "24 Bits Per Pixel"}
									"BitsPerPixel32" {WriteWordLine 0 3 "32 Bits Per Pixel"}
									Default {WriteWordLine 0 3 "Maximum allowed color depth could not be determined: $($Setting.MaximumColorDepth.Value)"}
								}
							}
							If($Setting.DisplayDegradeUserNotification.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Notify user when display mode is degraded: " $Setting.DisplayDegradeUserNotification.State
							}
							If($Setting.QueueingAndTossing.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Queueing and tossing: " $Setting.QueueingAndTossing.State
							}
							If($Setting.PersistentCache.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Graphics\Caching\Persistent Cache Threshold (Kbps): " $Setting.PersistentCache.Value
							}
							If($Setting.IcaKeepAliveTimeout.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Keep Alive\ICA keep alive timeout (seconds): " $Setting.IcaKeepAliveTimeout.Value
							}
							If($Setting.IcaKeepAlives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Keep Alive\ICA keep alives: "
								switch ($Setting.IcaKeepAlives.Value)
								{
									"DoNotSendKeepAlives" {WriteWordLine 0 3 "Do not send ICA keep alive messages"}
									"SendKeepAlives"      {WriteWordLine 0 3 "Send ICA keep alive messages"}
									Default {WriteWordLine 0 3 "ICA keep alives could not be determined: $($Setting.IcaKeepAlives.Value)"}
								}
							}
							If($Setting.MultimediaConferencing.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\Multimedia conferencing: " $Setting.MultimediaConferencing.State
							}
							If($Setting.MultimediaAcceleration.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\Windows Media Redirection: " $Setting.MultimediaAcceleration.State
							}
							If($Setting.MultimediaAccelerationDefaultBufferSize.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\Windows Media Redirection Buffer Size (seconds): " $Setting.MultimediaAccelerationDefaultBufferSize.Value
							}
							If($Setting.MultimediaAccelerationUseDefaultBufferSize.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multimedia\Windows Media Redirection Buffer Size Use: " $Setting.MultimediaAccelerationUseDefaultBufferSize.State
							}
							If($Setting.MultiPortPolicy.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\MultiStream Connections\Multi-Port Policy: " 
								$Tmp = $Setting.MultiPortPolicy.Value
								$cgpport1 = $Tmp.substring(0, $Tmp.indexof(";"))
								$cgpport2 = $Tmp.substring($cgpport1.length + 1 , $Tmp.indexof(";"))
								$cgpport3 = $Tmp.substring((($cgpport1.length + 1)+($cgpport2.length + 1)) , $Tmp.indexof(";"))
								$cgpport1priority = multiportpolicypriority $cgpport1.substring($cgpport1.length -1, 1)
								$cgpport2priority = multiportpolicypriority $cgpport2.substring($cgpport2.length -1, 1)
								$cgpport3priority = multiportpolicypriority $cgpport3.substring($cgpport3.length -1, 1)
								$cgpport1 = $cgpport1.substring(0, $cgpport1.indexof(","))
								$cgpport2 = $cgpport2.substring(0, $cgpport2.indexof(","))
								$cgpport3 = $cgpport3.substring(0, $cgpport3.indexof(","))
								WriteWordLine 0 3 "CGP port1: " $cgpport1 -nonewline 
								WriteWordLine 0 1 "priority: " $cgpport1priority[0]
								WriteWordLine 0 3 "CGP port2: " $cgpport2 -nonewline
								WriteWordLine 0 1 "priority: " $cgpport2priority[0]
								WriteWordLine 0 3 "CGP port3: " $cgpport3 -nonewline
								WriteWordLine 0 1 "priority: " $cgpport3priority[0]
								$Tmp = $null
								$cgpport1 = $null
								$cgpport2 = $null
								$cgpport3 = $null
								$cgpport1priority = $null
								$cgpport2priority = $null
								$cgpport3priority = $null
							}
							If($Setting.MultiStreamPolicy.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\MultiStream Connections\Multi-Stream: " $Setting.MultiStreamPolicy.State
							}
							If($Setting.PromptForPassword.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Security\Prompt for password: " $Setting.PromptForPassword.State
							}
							If($Setting.IdleTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Server Limits\Server idle timer interval (milliseconds): " $Setting.IdleTimerInterval.Value
							}
							If($Setting.SessionReliabilityConnections.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Reliability\Session reliability connections: " $Setting.SessionReliabilityConnections.State
							}
							If($Setting.SessionReliabilityPort.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Reliability\Session reliability port number: " $Setting.SessionReliabilityPort.Value
							}
							If($Setting.SessionReliabilityTimeout.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Reliability\Session reliability timeout (seconds): " $Setting.SessionReliabilityTimeout.Value
							}
							If($Setting.Shadowing.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Shadowing: " $Setting.Shadowing.State
							}
							If($Setting.LicenseServerHostName.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Licensing\License server host name: " $Setting.LicenseServerHostName.Value
							}
							If($Setting.LicenseServerPort.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Licensing\License server port: " $Setting.LicenseServerPort.Value
							}
							If($Setting.FarmName.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Power and Capacity Management\Farm name: " $Setting.FarmName.Value
							}
							If($Setting.WorkloadName.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Power and Capacity Management\Workload name: " $Setting.WorkloadName.Value
							}
							If($Setting.ConnectionAccessControl.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Connection access control: "
								switch ($Setting.ConnectionAccessControl.Value)
								{
									"AllowAny"                     {WriteWordLine 0 3 "Any connections"}
									"AllowTicketedConnectionsOnly" {WriteWordLine 0 3 "Citrix Access Gateway, Citrix Receiver, and Web Interface connections only"}
									"AllowAccessGatewayOnly"       {WriteWordLine 0 3 "Citrix Access Gateway connections only"}
									Default {WriteWordLine 0 3 "Connection access control could not be determined: $($Setting.ConnectionAccessControl.Value)"}
								}
							}
							If($Setting.DnsAddressResolution.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\DNS address resolution: " $Setting.DnsAddressResolution.State
							}
							If($Setting.FullIconCaching.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Full icon caching: " $Setting.FullIconCaching.State
							}
							If($Setting.LoadEvaluator.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Load Evaluator Name - Load evaluator: " $Setting.LoadEvaluator.Value
							}
							If($Setting.ProductEdition.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\XenApp product edition: " $Setting.ProductEdition.Value
							}
							If($Setting.ProductModel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\XenApp product model: " -nonewline
								switch ($Setting.ProductModel.Value)
								{
									"XenAppCCU"                  {WriteWordLine 0 0 "XenApp"}
									"XenDesktopConcurrentServer" {WriteWordLine 0 0 "XenDesktop Concurrent"}
									"XenDesktopUserDevice"       {WriteWordLine 0 0 "XenDesktop User Device"}
									Default {WriteWordLine 0 0 "XenApp product model could not be determined: $($Setting.ProductModel.Value)"}
								}
							}
							If($Setting.UserSessionLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Connection Limits\Limit user sessions: " $Setting.UserSessionLimit.Value
							}
							If($Setting.UserSessionLimitAffectsAdministrators.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Connection Limits\Limits on administrator sessions: " $Setting.UserSessionLimitAffectsAdministrators.State
							}
							If($Setting.UserSessionLimitLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Connection Limits\Logging of logon limit events: " $Setting.UserSessionLimitLogging.State
							}
							If($Setting.HealthMonitoring.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Health Monitoring and Recovery\Health monitoring: " $Setting.HealthMonitoring.State
							}
							If($Setting.HealthMonitoringTests.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Health Monitoring and Recovery\Health monitoring tests: " 
								[xml]$XML = $Setting.HealthMonitoringTests.Value
								ForEach($Test in $xml.hmrtests.tests.test)
								{
									WriteWordLine 0 3 "Name: " $test.name
									WriteWordLine 0 3 "File Location: " $test.file
									If($test.arguments)
									{
										WriteWordLine 0 3 "Arguments: " $test.arguments
									}
									WriteWordLine 0 3 "Description: " $test.description
									WriteWordLine 0 3 "Interval: " $test.interval
									WriteWordLine 0 3 "Time-out: " $test.timeout
									WriteWordLine 0 3 "Threshold: " $test.threshold
									WriteWordLine 0 3 "Recovery Action : " -nonewline
									switch ($test.RecoveryAction)
									{
										"AlertOnly"                     {WriteWordLine 0 0 "Alert Only"}
										"RemoveServerFromLoadBalancing" {WriteWordLine 0 0 "Remove Server from load balancing"}
										"RestartIma"                    {WriteWordLine 0 0 "Restart IMA"}
										"ShutdownIma"                   {WriteWordLine 0 0 "Shutdown IMA"}
										"RebootServer"                  {WriteWordLine 0 0 "Reboot Server"}
										Default {WriteWordLine 0 0 "Recovery Action could not be determined: $($test.RecoveryAction)"}
									}
									WriteWordLine 0 0 ""
								}
							}
							If($Setting.MaximumServersOfflinePercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Health Monitoring and Recovery\"
								WriteWordLine 0 3 "Max % of servers with logon control: " $Setting.MaximumServersOfflinePercent.Value
							}
							If($Setting.CpuManagementServerLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\CPU management server level: "
								switch ($Setting.CpuManagementServerLevel.Value)
								{
									"NoManagement" {WriteWordLine 0 3 "No CPU utilization management"}
									"Fair"         {WriteWordLine 0 3 "Fair sharing of CPU between sessions"}
									"Preferential" {WriteWordLine 0 3 "Preferential Load Balancing"}
									Default {WriteWordLine 0 3 "CPU management server level could not be determined: $($Setting.CpuManagementServerLevel.Value)"}
								}
							}
							If($Setting.MemoryOptimization.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization: " $Setting.MemoryOptimization.State
							}
							If($Setting.MemoryOptimizationExcludedPrograms.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization application exclusion list: "
								$array = $Setting.MemoryOptimizationExcludedPrograms.Values
								foreach( $element in $array)
								{
									WriteWordLine 0 3 $element
								}
							}
							If($Setting.MemoryOptimizationIntervalType.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization interval: " -nonewline
								switch ($Setting.MemoryOptimizationIntervalType.Value)
								{
									"AtStartup" {WriteWordLine 0 0 "Only at startup time"}
									"Daily"     {WriteWordLine 0 0 "Daily"}
									"Weekly"    {WriteWordLine 0 0 "Weekly"}
									"Monthly"   {WriteWordLine 0 0 "Monthly"}
									Default {WriteWordLine 0 0 " could not be determined: $($Setting.MemoryOptimizationIntervalType.Value)"}
								}
							}
							If($Setting.MemoryOptimizationDayOfMonth.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization schedule: "
								WriteWordLine 0 3 "day of month: " $Setting.MemoryOptimizationDayOfMonth.Value
							}
							If($Setting.MemoryOptimizationDayOfWeek.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization schedule: "
								WriteWordLine 0 3 "day of week: " $Setting.MemoryOptimizationDayOfWeek.Value
							}
							If($Setting.MemoryOptimizationTime.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization schedule time: " -nonewline
								$tmp = ConvertNumberToTime $Setting.MemoryOptimizationTime.Value
								WriteWordLine 0 0 $tmp
							}
							If($Setting.OfflineClientTrust.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Offline Applications\Offline app client trust: " $Setting.OfflineClientTrust.State
							}
							If($Setting.OfflineEventLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Offline Applications\Offline app event logging: " $Setting.OfflineEventLogging.State
							}
							If($Setting.OfflineLicensePeriod.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Offline Applications\Offline app license period - Days: " $Setting.OfflineLicensePeriod.Value
							}
							If($Setting.OfflineUsers.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Offline Applications\Offline app users: " 
								$array = $Setting.OfflineUsers.Values
								foreach( $element in $array)
								{
									WriteWordLine 0 3 $element
								}
							}
							If($Setting.RebootCustomMessage.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot custom warning: " $Setting.RebootCustomMessage.State
							}
							If($Setting.RebootCustomMessageText.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot custom warning text: " 
								WriteWordLine 0 3 $Setting.RebootCustomMessageText.Value
							}
							If($Setting.RebootDisableLogOnTime.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot logon disable time: "
								switch ($Setting.RebootDisableLogOnTime.Value)
								{
									"DoNotDisableLogOnsBeforeReboot" {WriteWordLine 0 3 "Do not disable logons before reboot"}
									"Disable5MinutesBeforeReboot"    {WriteWordLine 0 3 "Disable 5 minutes before reboot"}
									"Disable10MinutesBeforeReboot"   {WriteWordLine 0 3 "Disable 10 minutes before reboot"}
									"Disable15MinutesBeforeReboot"   {WriteWordLine 0 3 "Disable 15 minutes before reboot"}
									"Disable30MinutesBeforeReboot"   {WriteWordLine 0 3 "Disable 30 minutes before reboot"}
									"Disable60MinutesBeforeReboot"   {WriteWordLine 0 3 "Disable 60 minutes before reboot"}
									Default {WriteWordLine 0 3 "Reboot logon disable time could not be determined: $($Setting.RebootDisableLogOnTime.Value)"}
								}
							}
							If($Setting.RebootScheduleFrequency.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot schedule frequency - Days: " $Setting.RebootScheduleFrequency.Value
							}
							If($Setting.RebootScheduleRandomizationInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot schedule randomization interval"
								WriteWordLine 0 3 "Minutes: " $Setting.RebootScheduleRandomizationInterval.Value
							}
							If($Setting.RebootScheduleStartDate.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot schedule start date: " -nonewline
								$Tmp = ConvertIntegerToDate $Setting.RebootScheduleStartDate.Value
								WriteWordLine 0 0 $Tmp
							}
							If($Setting.RebootScheduleTime.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot schedule time: " -nonewline
								$tmp = ConvertNumberToTime $Setting.RebootScheduleTime.Value 						
								WriteWordLine 0 0 $Tmp
							}
							If($Setting.RebootWarningInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot warning interval: "
								switch ($Setting.RebootWarningInterval.Value)
								{
									"Every1Minute"   {WriteWordLine 0 3 "Every 1 Minute"}
									"Every3Minutes"  {WriteWordLine 0 3 "Every 3 Minutes"}
									"Every5Minutes"  {WriteWordLine 0 3 "Every 5 Minutes"}
									"Every10Minutes" {WriteWordLine 0 3 "Every 10 Minutes"}
									"Every15Minutes" {WriteWordLine 0 3 "Every 15 Minutes"}
									Default {WriteWordLine 0 3 "Reboot warning interval could not be determined: $($Setting.RebootWarningInterval.Value)"}
								}
							}
							If($Setting.RebootWarningStartTime.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot warning start time: "
								switch ($Setting.RebootWarningStartTime.Value)
								{
									"Start5MinutesBeforeReboot"  {WriteWordLine 0 3 "Start 5 Minutes Before Reboot"}
									"Start10MinutesBeforeReboot" {WriteWordLine 0 3 "Start 10 Minutes Before Reboot"}
									"Start15MinutesBeforeReboot" {WriteWordLine 0 3 "Start 15 Minutes Before Reboot"}
									"Start30MinutesBeforeReboot" {WriteWordLine 0 3 "Start 30 Minutes Before Reboot"}
									"Start60MinutesBeforeReboot" {WriteWordLine 0 3 "Start 60 Minutes Before Reboot"}
									Default {WriteWordLine 0 3 "Reboot warning start time could not be determined: $($Setting.RebootWarningStartTime.Value)"}
								}
							}
							If($Setting.RebootWarningMessage.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot warning to users: " $Setting.RebootWarningMessage.State
							}
							If($Setting.ScheduledReboots.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Settings\Reboot Behavior\Scheduled reboots: " $Setting.ScheduledReboots.State
							}
							If($Setting.FilterAdapterAddresses.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP adapter address filtering: " $Setting.FilterAdapterAddresses.State
							}
							If($Setting.EnhancedCompatibilityPrograms.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP compatibility programs list: " 
								$array = $Setting.EnhancedCompatibilityPrograms.Values
								foreach( $element in $array)
								{
									WriteWordLine 0 3 $element
								}
							}
							If($Setting.EnhancedCompatibility.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP enhanced compatibility: " $Setting.EnhancedCompatibility.State
							}
							If($Setting.FilterAdapterAddressesPrograms.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP filter adapter addresses programs list: " 
								$array = $Setting.FilterAdapterAddressesPrograms.Values
								foreach( $element in $array)
								{
									WriteWordLine 0 3 $element
								}
							}
							If($Setting.VirtualLoopbackSupport.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP loopback support: " $Setting.VirtualLoopbackSupport.State
							}
							If($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Virtual IP\Virtual IP virtual loopback programs list: " 
								$array = $Setting.VirtualLoopbackPrograms.Values
								foreach( $element in $array)
								{
									WriteWordLine 0 3 $element
								}
							}
							If($Setting.TrustXmlRequests.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "XML Service\Trust XML requests: " $Setting.TrustXmlRequests.State
							}
							If($Setting.XmlServicePort.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "XML Service\XML service port: " $Setting.XmlServicePort.Value
							}
						}
						Else
						{
							WriteWordLine 0 1 "User settings:"
							If($Setting.ClipboardRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Client clipboard redirection: " $Setting.ClipboardRedirection.State
							}
							If($Setting.DesktopLaunchForNonAdmins.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Desktop launches: " $Setting.DesktopLaunchForNonAdmins.State
							}
							If($Setting.NonPublishedProgramLaunching.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Launching of non-published programs during client connection: " $Setting.NonPublishedProgramLaunching.State
							}
							If($Setting.FlashAcceleration.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash acceleration: " $Setting.FlashAcceleration.State
							}
							If($Setting.FlashUrlColorList.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash background color list: "
								$Values = $Setting.FlashUrlColorList.Values
								ForEach($Value in $Values)
								{
									WriteWordLine 0 3 $Value
								}
								$Values = $null
							}
							If($Setting.FlashBackwardsCompatibility.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash backwards compatibility: " 
								WriteWordLine 0 3 $Setting.FlashBackwardsCompatibility.State
							}
							If($Setting.FlashDefaultBehavior.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash default behavior: "
								switch ($Setting.FlashDefaultBehavior.Value)
								{
									"Block"   {WriteWordLine 0 3 "Block Flash player"}
									"Disable" {WriteWordLine 0 3 "Disable Flash acceleration"}
									"Enable"  {WriteWordLine 0 3 "Enable Flash acceleration"}
									Default {WriteWordLine 0 3 "Flash default behavior could not be determined: $($Setting.FlashDefaultBehavior.Value)"}
								}
							}
							If($Setting.FlashEventLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash event logging: " $Setting.FlashEventLogging.State
							}
							If($Setting.FlashIntelligentFallback.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash intelligent fallback: " $Setting.FlashIntelligentFallback.State
							}
							If($Setting.FlashLatencyThreshold.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash latency threshold"
								WriteWordLine 0 3 "Value (milliseconds): " $Setting.FlashLatencyThreshold.Value
							}
							If($Setting.FlashServerSideContentFetchingWhitelist.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash server-side content fetching "
								WriteWordLine 0 3 "URL list: "
								$Values = $Setting.FlashServerSideContentFetchingWhitelist.Values
								ForEach($Value in $Values)
								{
									WriteWordLine 0 4 $Value
								}
								$Values = $null
							}
							If($Setting.FlashUrlCompatibilityList.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash URL compatibility list: " 
								$Values = $Setting.FlashUrlCompatibilityList.Values
								ForEach($Value in $Values)
								{
									$Spc = $Value.indexof(" ")
									$Action = $Value.substring(0, $Spc)
									If($Action -eq "CLIENT")
									{
										$Action = "Render On Client"
									}
									elseif ($Action -eq "SERVER")
									{
										$Action = "Render On Server"
									}
									elseif ($Action -eq "BLOCK")
									{
										$Action = "BLOCK           "
									}
									$Url = $Value.substring($Spc +1)
									WriteWordLine 0 3 "Action: " $Action -NoNewLine
									WriteWordLine 0 1 "URL: "$Url
								}
								$Values = $null
								$Action = $null
								$Url = $null
							}
							If($Setting.AllowSpeedFlash.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Legacy Server Side Optimizations\"
								WriteWordLine 0 3 "Flash quality adjustment: "
								switch ($Setting.AllowSpeedFlash.Value)
								{
									"NoOptimization"      {WriteWordLine 0 3 "Do not optimize Flash animation options"}
									"AllConnections"      {WriteWordLine 0 3 "Optimize Flash animation options for all connections"}
									"RestrictedBandwidth" {WriteWordLine 0 3 "Optimize Flash animation options for low bandwidth connections only"}
									Default {WriteWordLine 0 3 "Flash quality adjustment could not be determined: $($Setting.AllowSpeedFlash.Value)"}
								}
							}
							If($Setting.AudioPlugNPlay.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Audio\Audio Plug N Play: " $Setting.AudioPlugNPlay.State
							}
							If($Setting.AudioQuality.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Audio\Audio quality: "
								switch ($Setting.AudioQuality.Value)
								{
									"Low"    {WriteWordLine 0 3 "Low - for low-speed connections"}
									"Medium" {WriteWordLine 0 3 "Medium - optimized for speech"}
									"High"   {WriteWordLine 0 3 "High - high definition audio"}
									Default {WriteWordLine 0 3 "Audio quality could not be determined: $($Setting.AudioQuality.Value)"}
								}
							}
							If($Setting.ClientAudioRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Audio\Client audio redirection: " $Setting.ClientAudioRedirection.State
							}
							If($Setting.MicrophoneRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Audio\Client microphone redirection: " $Setting.MicrophoneRedirection.State
							}
							If($Setting.AudioBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Audio redirection bandwidth limit (Kbps): " $Setting.AudioBandwidthLimit.Value
							}
							If($Setting.AudioBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Audio redirection bandwidth limit %: " $Setting.AudioBandwidthPercent.Value
							}
							If($Setting.USBBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Client USB device redirection bandwidth limit: " $Setting.USBBandwidthLimit.Value
							}
							If($Setting.USBBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Client USB device redirection bandwidth limit %: " $Setting.USBBandwidthPercent.Value
							}
							If($Setting.ClipboardBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Clipboard redirection bandwidth limit (Kbps): " $Setting.ClipboardBandwidthLimit.Value
							}
							If($Setting.ClipboardBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Clipboard redirection bandwidth limit %: " $Setting.ClipboardBandwidthPercent.Value
							}
							If($Setting.ComPortBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\COM port redirection bandwidth limit (Kbps): " $Setting.ComPortBandwidthLimit.Value
							}
							If($Setting.ComPortBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\COM port redirection bandwidth limit %: " $Setting.ComPortBandwidthPercent.Value
							}
							If($Setting.FileRedirectionBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\File redirection bandwidth limit (Kbps): " $Setting.FileRedirectionBandwidthLimit.Value
							}
							If($Setting.FileRedirectionBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\File redirection bandwidth limit %: " $Setting.FileRedirectionBandwidthPercent.Value
							}
							If($Setting.HDXMultimediaBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration "
								WriteWordLine 0 3 "bandwidth limit (Kbps): " $Setting.HDXMultimediaBandwidthLimit.Value
							}
							If($Setting.HDXMultimediaBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration "
								WriteWordLine 0 3 "bandwidth limit %: " $Setting.HDXMultimediaBandwidthPercent.Value
							}
							If($Setting.LptBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\LPT port redirection bandwidth limit (Kbps): " $Setting.LptBandwidthLimit.Value
							}
							If($Setting.LptBandwidthLimitPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\LPT port redirection bandwidth limit %: " $Setting.LptBandwidthLimitPercent.Value
							}
							If($Setting.OverallBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Overall session bandwidth limit (Kbps): " $Setting.OverallBandwidthLimit.Value
							}
							If($Setting.PrinterBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Printer redirection bandwidth limit (Kbps): " $Setting.PrinterBandwidthLimit.Value
							}
							If($Setting.PrinterBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\Printer redirection bandwidth limit %: " $Setting.PrinterBandwidthPercent.Value
							}
							If($Setting.TwainBandwidthLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\TWAIN device redirection bandwidth limit (Kbps): " $Setting.TwainBandwidthLimit.Value
							}
							If($Setting.TwainBandwidthPercent.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Bandwidth\TWAIN device redirection bandwidth limit %: " $Setting.TwainBandwidthPercent.Value
							}
							If($Setting.DesktopWallpaper.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Desktop UI\Desktop wallpaper: " $Setting.DesktopWallpaper.State
							}
							If($Setting.MenuAnimation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Desktop UI\Menu animation: " $Setting.MenuAnimation.State
							}
							If($Setting.WindowContentsVisibleWhileDragging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Desktop UI\View window contents while dragging: " $Setting.WindowContentsVisibleWhileDragging.State
							}
							If($Setting.AutoConnectDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Auto connect client drives: " $Setting.AutoConnectDrives.State
							}
							If($Setting.ClientDriveRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client drive redirection: " $Setting.ClientDriveRedirection.State
							}
							If($Setting.ClientFixedDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client fixed drives: " $Setting.ClientFixedDrives.State
							}
							If($Setting.ClientFloppyDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client floppy drives: " $Setting.ClientFloppyDrives.State
							}
							If($Setting.ClientNetworkDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client network drives: " $Setting.ClientNetworkDrives.State
							}
							If($Setting.ClientOpticalDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client optical drives: " $Setting.ClientOpticalDrives.State
							}
							If($Setting.ClientRemoveableDrives.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Client removable drives: " $Setting.ClientRemoveableDrives.State
							}
							If($Setting.HostToClientRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Host to client redirection: " $Setting.HostToClientRedirection.State
							}
							If($Setting.ReadOnlyMappedDrive.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Read-only client drive access: " $Setting.ReadOnlyMappedDrive.State
							}
							If($Setting.SpecialFolderRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Special folder redirection: " $Setting.SpecialFolderRedirection.State
							}
							If($Setting.AsynchronousWrites.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\File Redirection\Use asynchronous writes: " $Setting.AsynchronousWrites.State
							}
							If($Setting.MultiStream.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Multi-Stream Connections\Multi-Stream: " $Setting.MultiStream.State
							}
							If($Setting.ClientComPortsAutoConnection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Port Redirection\Auto connect client COM ports: " $Setting.ClientComPortsAutoConnection.State
							}
							If($Setting.ClientLptPortsAutoConnection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Port Redirection\Auto connect client LPT ports: " $Setting.ClientLptPortsAutoConnection.State
							}
							If($Setting.ClientComPortRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Port Redirection\Client COM port redirection: " $Setting.ClientComPortRedirection.State
							}
							If($Setting.ClientLptPortRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Port Redirection\Client LPT port redirection: " $Setting.ClientLptPortRedirection.State
							}
							If($Setting.ClientPrinterRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client printer redirection: " $Setting.ClientPrinterRedirection.State
							}
							If($Setting.DefaultClientPrinter.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Default printer - Choose client's default printer: " 
								switch ($Setting.DefaultClientPrinter.Value)
								{
									"ClientDefault" {WriteWordLine 0 3 "Set default printer to the client's main printer"}
									"DoNotAdjust"   {WriteWordLine 0 3 "Do not adjust the user's default printer"}
									Default {WriteWordLine 0 0 "Default printer could not be determined: $($Setting.DefaultClientPrinter.Value)"}
								}
							}
							If($Setting.AutoCreationEventLogPreference.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Printer auto-creation event log preference: " 
								switch ($Setting.AutoCreationEventLogPreference.Value)
								{
									"LogErrorsOnly"        {WriteWordLine 0 3 "Log errors only"}
									"LogErrorsAndWarnings" {WriteWordLine 0 3 "Log errors and warnings"}
									"DoNotLog"             {WriteWordLine 0 3 "Do not log errors or warnings"}
									Default {WriteWordLine 0 3 "Printer auto-creation event log preference could not be determined: $($Setting.AutoCreationEventLogPreference.Value)"}
								}
							}
							If($Setting.SessionPrinters.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Session printers:" 
								$valArray = $Setting.SessionPrinters.Values
								foreach( $printer in $valArray )
								{
									$prArray = $printer.Split( ',' )
									foreach( $element in $prArray )
									{
										if( $element.SubString( 0, 2 ) -eq "\\" )
										{
											$index = $element.SubString( 2 ).IndexOf( '\' )
											if( $index -ge 0 )
											{
												$server = $element.SubString( 0, $index + 2 )
												$share  = $element.SubString( $index + 3 )
												WriteWordLine 0 3 "Server: $server"
												WriteWordLine 0 3 "Shared Name: $share"
											}
										}
										Else
										{
											$tmp = $element.SubString( 0, 4 )
											Switch ($tmp)
											{
												"copi" 
												{
													$txt="Count:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt $tmp2"
													}
												}
												"coll"
												{
													$txt="Collate:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt $tmp2"
													}
												}
												"scal"
												{
													$txt="Scale (%):"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt $tmp2"
													}
												}
												"colo"
												{
													$txt="Color:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt " -nonewline
														Switch ($tmp2)
														{
															1 {WriteWordLine 0 0 "Monochrome"}
															2 {WriteWordLine 0 0 "Color"}
															Default {WriteWordLine 0 3 "Color could not be determined: $($element)"}
														}
													}
												}
												"prin"
												{
													$txt="Print Quality:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt " -nonewline
														Switch ($tmp2)
														{
															-1 {WriteWordLine 0 0 "150 dpi"}
															-2 {WriteWordLine 0 0 "300 dpi"}
															-3 {WriteWordLine 0 0 "600 dpi"}
															-4 {WriteWordLine 0 0 "1200 dpi"}
															Default 
															{
																WriteWordLine 0 0 "Custom..."
																WriteWordLine 0 3 "X resolution: " $tmp2
															}
														}
													}
												}
												"yres"
												{
													$txt="Y resolution:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt $tmp2"
													}
												}
												"orie"
												{
													$txt="Orientation:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt " -nonewline
														switch ($tmp2)
														{
															"portrait"  {WriteWordLine 0 0 "Portrait"}
															"landscape" {WriteWordLine 0 0 "Landscape"}
															Default {WriteWordLine 0 3 "Orientation could not be determined: $($Element)"}
														}
													}
												}
												"dupl"
												{
													$txt="Duplex:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt " -nonewline
														switch ($tmp2)
														{
															1 {WriteWordLine 0 0 "Simplex"}
															2 {WriteWordLine 0 0 "Vertical"}
															3 {WriteWordLine 0 0 "Horizontal"}
															Default {WriteWordLine 0 3 "Duplex could not be determined: $($Element)"}
														}
													}
												}
												"pape"
												{
													$txt="Paper Size:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt " -nonewline
														switch ($tmp2)
														{
															1   {WriteWordLine 0 0 "Letter"}
															2   {WriteWordLine 0 0 "Letter Small"}
															3   {WriteWordLine 0 0 "Tabloid"}
															4   {WriteWordLine 0 0 "Ledger"}
															5   {WriteWordLine 0 0 "Legal"}
															6   {WriteWordLine 0 0 "Statement"}
															7   {WriteWordLine 0 0 "Executive"}
															8   {WriteWordLine 0 0 "A3"}
															9   {WriteWordLine 0 0 "A4"}
															10  {WriteWordLine 0 0 "A4 Small"}
															11  {WriteWordLine 0 0 "A5"}
															12  {WriteWordLine 0 0 "B4 (JIS)"}
															13  {WriteWordLine 0 0 "B5 (JIS)"}
															14  {WriteWordLine 0 0 "Folio"}
															15  {WriteWordLine 0 0 "Quarto"}
															16  {WriteWordLine 0 0 "10X14"}
															17  {WriteWordLine 0 0 "11X17"}
															18  {WriteWordLine 0 0 "Note"}
															19  {WriteWordLine 0 0 "Envelope #9"}
															20  {WriteWordLine 0 0 "Envelope #10"}
															21  {WriteWordLine 0 0 "Envelope #11"}
															22  {WriteWordLine 0 0 "Envelope #12"}
															23  {WriteWordLine 0 0 "Envelope #14"}
															24  {WriteWordLine 0 0 "C Size Sheet"}
															25  {WriteWordLine 0 0 "D Size Sheet"}
															26  {WriteWordLine 0 0 "E Size Sheet"}
															27  {WriteWordLine 0 0 "Envelope DL"}
															28  {WriteWordLine 0 0 "Envelope C5"}
															29  {WriteWordLine 0 0 "Envelope C3"}
															30  {WriteWordLine 0 0 "Envelope C4"}
															31  {WriteWordLine 0 0 "Envelope C6"}
															32  {WriteWordLine 0 0 "Envelope C65"}
															33  {WriteWordLine 0 0 "Envelope B4"}
															34  {WriteWordLine 0 0 "Envelope B5"}
															35  {WriteWordLine 0 0 "Envelope B6"}
															36  {WriteWordLine 0 0 "Envelope Italy"}
															37  {WriteWordLine 0 0 "Envelope Monarch"}
															38  {WriteWordLine 0 0 "Envelope Personal"}
															39  {WriteWordLine 0 0 "US Std Fanfold"}
															40  {WriteWordLine 0 0 "German Std Fanfold"}
															41  {WriteWordLine 0 0 "German Legal Fanfold"}
															42  {WriteWordLine 0 0 "B4 (ISO)"}
															43  {WriteWordLine 0 0 "Japanese Postcard"}
															44  {WriteWordLine 0 0 "9X11"}
															45  {WriteWordLine 0 0 "10X11"}
															46  {WriteWordLine 0 0 "15X11"}
															47  {WriteWordLine 0 0 "Envelope Invite"}
															48  {WriteWordLine 0 0 "Reserved - DO NOT USE"}
															49  {WriteWordLine 0 0 "Reserved - DO NOT USE"}
															50  {WriteWordLine 0 0 "Letter Extra"}
															51  {WriteWordLine 0 0 "Legal Extra"}
															52  {WriteWordLine 0 0 "Tabloid Extra"}
															53  {WriteWordLine 0 0 "A4 Extra"}
															54  {WriteWordLine 0 0 "Letter Transverse"}
															55  {WriteWordLine 0 0 "A4 Transverse"}
															56  {WriteWordLine 0 0 "Letter Extra Transverse"}
															57  {WriteWordLine 0 0 "A Plus"}
															58  {WriteWordLine 0 0 "B Plus"}
															59  {WriteWordLine 0 0 "Letter Plus"}
															60  {WriteWordLine 0 0 "A4 Plus"}
															61  {WriteWordLine 0 0 "A5 Transverse"}
															62  {WriteWordLine 0 0 "B5 (JIS) Transverse"}
															63  {WriteWordLine 0 0 "A3 Extra"}
															64  {WriteWordLine 0 0 "A5 Extra"}
															65  {WriteWordLine 0 0 "B5 (ISO) Extra"}
															66  {WriteWordLine 0 0 "A2"}
															67  {WriteWordLine 0 0 "A3 Transverse"}
															68  {WriteWordLine 0 0 "A3 Extra Transverse"}
															69  {WriteWordLine 0 0 "Japanese Double Postcard"}
															70  {WriteWordLine 0 0 "A6"}
															71  {WriteWordLine 0 0 "Japanese Envelope Kaku #2"}
															72  {WriteWordLine 0 0 "Japanese Envelope Kaku #3"}
															73  {WriteWordLine 0 0 "Japanese Envelope Chou #3"}
															74  {WriteWordLine 0 0 "Japanese Envelope Chou #4"}
															75  {WriteWordLine 0 0 "Letter Rotated"}
															76  {WriteWordLine 0 0 "A3 Rotated"}
															77  {WriteWordLine 0 0 "A4 Rotated"}
															78  {WriteWordLine 0 0 "A5 Rotated"}
															79  {WriteWordLine 0 0 "B4 (JIS) Rotated"}
															80  {WriteWordLine 0 0 "B5 (JIS) Rotated"}
															81  {WriteWordLine 0 0 "Japanese Postcard Rotated"}
															82  {WriteWordLine 0 0 "Double Japanese Postcard Rotated"}
															83  {WriteWordLine 0 0 "A6 Rotated"}
															84  {WriteWordLine 0 0 "Japanese Envelope Kaku #2 Rotated"}
															85  {WriteWordLine 0 0 "Japanese Envelope Kaku #3 Rotated"}
															86  {WriteWordLine 0 0 "Japanese Envelope Chou #3 Rotated"}
															87  {WriteWordLine 0 0 "Japanese Envelope Chou #4 Rotated"}
															88  {WriteWordLine 0 0 "B6 (JIS)"}
															89  {WriteWordLine 0 0 "B6 (JIS) Rotated"}
															90  {WriteWordLine 0 0 "12X11"}
															91  {WriteWordLine 0 0 "Japanese Envelope You #4"}
															92  {WriteWordLine 0 0 "Japanese Envelope You #4 Rotated"}
															93  {WriteWordLine 0 0 "PRC 16K"}
															94  {WriteWordLine 0 0 "PRC 32K"}
															95  {WriteWordLine 0 0 "PRC 32K(Big)"}
															96  {WriteWordLine 0 0 "PRC Envelope #1"}
															97  {WriteWordLine 0 0 "PRC Envelope #2"}
															98  {WriteWordLine 0 0 "PRC Envelope #3"}
															99  {WriteWordLine 0 0 "PRC Envelope #4"}
															100 {WriteWordLine 0 0 "PRC Envelope #5"}
															101 {WriteWordLine 0 0 "PRC Envelope #6"}
															102 {WriteWordLine 0 0 "PRC Envelope #7"}
															103 {WriteWordLine 0 0 "PRC Envelope #8"}
															104 {WriteWordLine 0 0 "PRC Envelope #9"}
															105 {WriteWordLine 0 0 "PRC Envelope #10"}
															106 {WriteWordLine 0 0 "PRC 16K Rotated"}
															107 {WriteWordLine 0 0 "PRC 32K Rotated"}
															108 {WriteWordLine 0 0 "PRC 32K(Big) Rotated"}
															109 {WriteWordLine 0 0 "PRC Envelope #1 Rotated"}
															110 {WriteWordLine 0 0 "PRC Envelope #2 Rotated"}
															111 {WriteWordLine 0 0 "PRC Envelope #3 Rotated"}
															112 {WriteWordLine 0 0 "PRC Envelope #4 Rotated"}
															113 {WriteWordLine 0 0 "PRC Envelope #5 Rotated"}
															114 {WriteWordLine 0 0 "PRC Envelope #6 Rotated"}
															115 {WriteWordLine 0 0 "PRC Envelope #7 Rotated"}
															116 {WriteWordLine 0 0 "PRC Envelope #8 Rotated"}
															117 {WriteWordLine 0 0 "PRC Envelope #9 Rotated"}
															Default {WriteWordLine 0 3 "Paper Size could not be determined: $($element)"}
														}
													}
												}
												"form"
												{
													$txt="Form Name:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														If($tmp2.length -gt 0)
														{
															WriteWordLine 0 3 "$txt $tmp2"
														}
													}
												}
												"true"
												{
													$txt="TrueType:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt " -nonewline
														switch ($tmp2)
														{
															1 {WriteWordLine 0 0 "Bitmap"}
															2 {WriteWordLine 0 0 "Download"}
															3 {WriteWordLine 0 0 "Substitute"}
															4 {WriteWordLine 0 0 "Outline"}
															Default {WriteWordLine 0 3 "TrueType could not be determined: $($Element)"}
														}
													}
												}
												"mode" 
												{
													$txt="Printer Model:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														WriteWordLine 0 3 "$txt $tmp2"
													}
												}
												"loca" 
												{
													$txt="Location:"
													$index = $element.SubString( 0 ).IndexOf( '=' )
													if( $index -ge 0 )
													{
														$tmp2 = $element.SubString( $index + 1 )
														If($tmp2.length -gt 0)
														{
															WriteWordLine 0 3 "$txt $tmp2"
														}
													}
												}
												Default {WriteWordLine 0 3 "Session printer setting could not be determined: $($Element)"}
											}
										}
									}
									WriteWordLine 0 0 ""
								}
							}
							If($Setting.WaitForPrintersToBeCreated.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Wait for printers to be created (desktop): " $Setting.WaitForPrintersToBeCreated.Values
							}
							If($Setting.ClientPrinterAutoCreation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Auto-create client printers: "
								switch ($Setting.ClientPrinterAutoCreation.Value)
								{
									"DoNotAutoCreate"    {WriteWordLine 0 3 "Do not auto-create client printers"}
									"DefaultPrinterOnly" {WriteWordLine 0 3 "Auto-create the client's default printer only"}
									"LocalPrintersOnly"  {WriteWordLine 0 3 "Auto-create local (non-network) client printers only"}
									"AllPrinters"        {WriteWordLine 0 3 "Auto-create all client printers"}
									Default {WriteWordLine 0 3 "Auto-create client printers could not be determined: $($Setting.ClientPrinterAutoCreation.Value)"}
								}
							}
							If($Setting.GenericUniversalPrinterAutoCreation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Auto-create generic universal printer: " $Setting.GenericUniversalPrinterAutoCreation.Value
							}
							If($Setting.ClientPrinterNames.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Client printer names: " 
								switch ($Setting.ClientPrinterNames.Value)
								{
									"StandardPrinterNames" {WriteWordLine 0 3 "Standard printer names"}
									"LegacyPrinterNames"   {WriteWordLine 0 3 "Legacy printer names"}
									Default {WriteWordLine 0 3 "Client printer names could not be determined: $($Setting.ClientPrinterNames.Value)"}
								}
							}
							If($Setting.DirectConnectionsToPrintServers.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Direct connections to print servers: " $Setting.DirectConnectionsToPrintServers.State
							}
							If($Setting.PrinterDriverMappings.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Printer driver mapping and compatibility: " 
								$array = $Setting.PrinterDriverMappings.Values
								foreach( $element in $array)
								{
									WriteWordLine 0 3 $element
								}
							}
							If($Setting.PrinterPropertiesRetention.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Printer properties retention: " 
								switch ($Setting.PrinterPropertiesRetention.Value)
								{
									"SavedOnClientDevice"   {WriteWordLine 0 3 "Saved on the client device only"}
									"RetainedInUserProfile" {WriteWordLine 0 3 "Retained in user profile only"}
									"FallbackToProfile"     {WriteWordLine 0 3 "Held in profile only if not saved on client"}
									"DoNotRetain"           {WriteWordLine 0 3 "Do not retain printer properties"}
									Default {WriteWordLine 0 3 "Printer properties retention could not be determined: $($Setting.PrinterPropertiesRetention.Value)"}
								}
							}
							If($Setting.RetainedAndRestoredClientPrinters.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Client Printers\Retained and restored client printers: " $Setting.RetainedAndRestoredClientPrinters.State
							}
							If($Setting.InboxDriverAutoInstallation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Drivers\Automatic installation of in-box printer drivers: " $Setting.InboxDriverAutoInstallation.State
							}
							If($Setting.UniversalDriverPriority.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Drivers\Universal driver preference: " 
								WriteWordLine 0 3 $Setting.UniversalDriverPriority.Value
							}
							If($Setting.UniversalPrintDriverUsage.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Drivers\Universal print driver usage: " 
								switch ($Setting.UniversalPrintDriverUsage.Value)
								{
									"SpecificOnly"       {WriteWordLine 0 3 "Use only printer model specific drivers"}
									"UpdOnly"            {WriteWordLine 0 3 "Use universal printing only"}
									"FallbackToUpd"      {WriteWordLine 0 3 "Use universal printing only if requested driver is unavailable"}
									"FallbackToSpecific" {WriteWordLine 0 3 "Use printer model specific drivers only if universal printing is unavailable"}
									Default {WriteWordLine 0 3 "Universal print driver usage could not be determined: $($Setting.UniversalPrintDriverUsage.Value)"}
								}
							}
							If($Setting.EMFProcessingMode.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing EMF processing mode: " 
								switch ($Setting.EMFProcessingMode.Value)
								{
									"ReprocessEMFsForPrinter" {WriteWordLine 0 3 "Reprocess EMFs for printer"}
									"SpoolDirectlyToPrinter"  {WriteWordLine 0 3 "Spool directly to printer"}
									Default {WriteWordLine 0 3 "Universal printing EMF processing mode could not be determined: $($Setting.EMFProcessingMode.Value)"}
								}
							}
							If($Setting.ImageCompressionLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing image compression limit: " 
								switch ($Setting.ImageCompressionLimit.Value)
								{
									"NoCompression"       {WriteWordLine 0 3 "No compression"}
									"LosslessCompression" {WriteWordLine 0 3 "Best quality (lossless compression)"}
									"MinimumCompression"  {WriteWordLine 0 3 "High quality"}
									"MediumCompression"   {WriteWordLine 0 3 "Standard quality"}
									"MaximumCompression"  {WriteWordLine 0 3 "Reduced quality (maximum compression)"}
									Default {WriteWordLine 0 3 "Universal printing image compression limit could not be determined: $($Setting.ImageCompressionLimit.Value)"}
								}
							}
							If($Setting.UPDCompressionDefaults.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing optimization default: "
								$Tmp = $Setting.UPDCompressionDefaults.Value.replace(";","`n`t`t`t`t")
								WriteWordLine 0 3 $Tmp
								$Tmp = $null
							}
							If($Setting.UniversalPrintingPreviewPreference.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing preview preference: " 
								switch ($Setting.UniversalPrintingPreviewPreference.Value)
								{
									"NoPrintPreview"        {WriteWordLine 0 3 "Do not use print preview for auto-created or generic universal printers"}
									"AutoCreatedOnly"       {WriteWordLine 0 3 "Use print preview for auto-created printers only"}
									"GenericOnly"           {WriteWordLine 0 3 "Use print preview for generic universal printers only"}
									"AutoCreatedAndGeneric" {WriteWordLine 0 3 "Use print preview for both auto-created and generic universal printers"}
									Default {WriteWordLine 0 3 "Universal printing preview preference could not be determined: $($Setting.UniversalPrintingPreviewPreference.Value)"}
								}
							}
							If($Setting.DPILimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing print quality limit: " 
								switch ($Setting.DPILimit.Value)
								{
									"Draft"            {WriteWordLine 0 3 "Draft (150 DPI)"}
									"LowResolution"    {WriteWordLine 0 3 "Low Resolution (300 DPI)"}
									"MediumResolution" {WriteWordLine 0 3 "Medium Resolution (600 DPI)"}
									"HighResolution"   {WriteWordLine 0 3 "High Resolution (1200 DPI)"}
									"Unlimited "       {WriteWordLine 0 3 "No Limit"}
									Default {WriteWordLine 0 3 "Universal printing print quality limit could not be determined: $($Setting.DPILimit.Value)"}
								}
							}
							If($Setting.MinimumEncryptionLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Security\SecureICA minimum encryption level: " 
								switch ($Setting.MinimumEncryptionLevel.Value)
								{
									"Unknown" {WriteWordLine 0 3 "Unknown encryption"}
									"Basic"   {WriteWordLine 0 3 "Basic"}
									"LogOn"   {WriteWordLine 0 3 "RC5 (128 bit) logon only"}
									"Bits40"  {WriteWordLine 0 3 "RC5 (40 bit)"}
									"Bits56"  {WriteWordLine 0 3 "RC5 (56 bit)"}
									"Bits128" {WriteWordLine 0 3 "RC5 (128 bit)"}
									Default {WriteWordLine 0 3 "SecureICA minimum encryption level could not be determined: $($Setting.MinimumEncryptionLevel.Value)"}
								}
							}
							If($Setting.ConcurrentLogOnLimit.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session limits\Concurrent logon limit: " $Setting.ConcurrentLogOnLimit.Value
							}
							If($Setting.SessionDisconnectTimer.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Disconnected session timer: " $Setting.SessionDisconnectTimer.State
							}
							If($Setting.SessionDisconnectTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Disconnected session timer interval (minutes): " $Setting.SessionDisconnectTimerInterval.Value
							}
							If($Setting.LingerDisconnectTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Linger Disconnect Timer Interval (minutes): " $Setting.LingerDisconnectTimerInterval.Value
							}
							If($Setting.LingerTerminateTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Linger Terminate Timer Interval - Value (minutes): " $Setting.LingerTerminateTimerInterval.Value
							}
							If($Setting.PrelaunchDisconnectTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Pre-launch Disconnect Timer Interval - Value (minutes): " $Setting.PrelaunchDisconnectTimerInterval.Value
							}
							If($Setting.PrelaunchTerminateTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Pre-launch Terminate Timer Interval - Value (minutes): " $Setting.PrelaunchTerminateTimerInterval.Value
							}
							If($Setting.SessionConnectionTimer.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Session connection timer: " $Setting.SessionConnectionTimer.State
							}
							If($Setting.SessionConnectionTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Session connection timer interval - Value (minutes): " $Setting.SessionConnectionTimerInterval.Value
							}
							If($Setting.SessionIdleTimer.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Session idle timer: " $Setting.SessionIdleTimer.State
							}
							If($Setting.SessionIdleTimerInterval.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Session Limits\Session idle timer interval - Value (minutes): " $Setting.SessionIdleTimerInterval.Value
							}
							If($Setting.ShadowInput.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Input from shadow connections: " $Setting.ShadowInput.State
							}
							If($Setting.ShadowLogging.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Log shadow attempts: " $Setting.ShadowLogging.State
							}
							If($Setting.ShadowUserNotification.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Notify user of pending shadow connections: " $Setting.ShadowUserNotification.State
							}
							If($Setting.ShadowAllowList.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Users who can shadow other users: " 
								$array = $Setting.ShadowAllowList.Values
								#gui only shows computer\account or domain\account
								#what is stored is:
								#0x05/NT/XA65\ANON000/S-1-5-21-1307341077-4083623718-4268213518-1028 (workgroup/local)
								#0x05/NT/XA651\CTX_CPUUSER/S-1-5-21-1200344839-3835835227-1016768578-1002 (domain/local)
								#0x05/NT/WEBSTERSLAB\ADMINISTRATOR/S-1-5-21-3679396586-1061193519-2853834051-500 (domain user)
								#0x05/NT/WEBSTERSLAB\DOMAIN ADMINS/S-1-5-21-3679396586-1061193519-2853834051-512 (domain group)
								#we only need the computer\account or domain\account
								#first 9 characters are 0x05/NT/ for all account types
								#since PoSH starts counting at 0 we don't need the first 9 characters
								#Then we need the position of the first / after the computer\account
								#what is left between the two is what we need
								foreach( $element in $array)
								{
									$x = $element.indexof("/",8)
									$tmp = $element.substring(8,$x-8)
									WriteWordLine 0 3 $tmp
								}
							}
							If($Setting.ShadowDenyList.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Shadowing\Users who cannot shadow other users: " 
								$array = $Setting.ShadowDenyList.Values
								foreach( $element in $array)
								{
									$x = $element.indexof("/",8)
									$tmp = $element.substring(8,$x-8)
									WriteWordLine 0 3 $tmp
								}
							}
							If($Setting.LocalTimeEstimation.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Time Zone Control\Estimate local time for legacy clients: " $Setting.LocalTimeEstimation.State
							}
							If($Setting.SessionTimeZone.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Time Zone Control\Use local time of client: " 
								switch ($Setting.SessionTimeZone.Value)
								{
									"UseServerTimeZone" {WriteWordLine 0 3 "Use server time zone"}
									"UseClientTimeZone" {WriteWordLine 0 3 "Use client time zone"}
									Default {WriteWordLine 0 3 "Use local time of client could not be determined: $($Setting.SessionTimeZone.Value)"}
								}
							}
							If($Setting.TwainRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\TWAIN devices\Client TWAIN device redirection: " $Setting.TwainRedirection.State
							}
							If($Setting.TwainCompressionLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\TWAIN devices\TWAIN compression level: " 
								switch ($Setting.TwainCompressionLevel.Value)
								{
									"None"   {WriteWordLine 0 3 "None"}
									"Low"    {WriteWordLine 0 3 "Low"}
									"Medium" {WriteWordLine 0 3 "Medium"}
									"High"   {WriteWordLine 0 3 "High"}
									Default {WriteWordLine 0 3 "TWAIN compression level could not be determined: $($Setting.TwainCompressionLevel.Value)"}
								}
							}
							If($Setting.UsbDeviceRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\USB devices\Client USB device redirection: " $Setting.UsbDeviceRedirection.State
							}
							If($Setting.UsbDeviceRedirectionRules.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\USB devices\Client USB device redirection rules: " 
								$array = $Setting.UsbDeviceRedirectionRules.Values
								foreach( $element in $array)
								{
									WriteWordLine 0 3 $element
								}
							}
							If($Setting.UsbPlugAndPlayRedirection.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\USB devices\Client USB Plug and Play device redirection: " $Setting.UsbPlugAndPlayRedirection.State
							}
							If($Setting.FramesPerSecond.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Visual Display\Max Frames Per Second (fps): " $Setting.FramesPerSecond.Value
							}
							If($Setting.ProgressiveCompressionLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Visual Display\Moving Images\Progressive compression level: " -nonewline
								switch ($Setting.ProgressiveCompressionLevel.Value)
								{
									"UltraHigh" {WriteWordLine 0 0 "Ultra high"}
									"VeryHigh"  {WriteWordLine 0 0 "Very high"}
									"High"      {WriteWordLine 0 0 "High"}
									"Normal"    {WriteWordLine 0 0 "Normal"}
									"Low"       {WriteWordLine 0 0 "Low"}
									Default {WriteWordLine 0 0 "Progressive compression level could not be determined: $($Setting.ProgressiveCompressionLevel.Value)"}
								}
							}
							If($Setting.ProgressiveCompressionThreshold.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Visual Display\Moving Images\Progressive compression "
								WriteWordLine 0 3 "threshold value (Kbps): " $Setting.ProgressiveCompressionThreshold.Value
							}
							If($Setting.ExtraColorCompression.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Visual Display\Still Images\Extra Color Compression: " $Setting.ExtraColorCompression.State
							}
							If($Setting.ExtraColorCompressionThreshold.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Visual Display\Still Images\Extra Color Compression Threshold (Kbps): " $Setting.ExtraColorCompressionThreshold.Value
							}
							If($Setting.ProgressiveHeavyweightCompression.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Visual Display\Still Images\Heavyweight compression: " $Setting.ProgressiveHeavyweightCompression.State
							}
							If($Setting.LossyCompressionLevel.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Visual Display\Still Images\Lossy compression level: " 
								switch ($Setting.LossyCompressionLevel.Value)
								{
									"None"   {WriteWordLine 0 3 "None"}
									"Low"    {WriteWordLine 0 3 "Low"}
									"Medium" {WriteWordLine 0 3 "Medium"}
									"High"   {WriteWordLine 0 3 "High"}
									Default {WriteWordLine 0 3 "Lossy compression level could not be determined: $($Setting.LossyCompressionLevel.Value)"}
								}
							}
							If($Setting.LossyCompressionThreshold.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "ICA\Visual Display\Still Images\Lossy compression threshold value (Kbps): " 
								WriteWordLine 0 3 $Setting.LossyCompressionThreshold.Value
							}
							If($Setting.SessionImportance.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Session Settings\Session importance: " 
								switch ($Setting.SessionImportance.Value)
								{
									"Low"    {WriteWordLine 0 3 "Low"}
									"Normal" {WriteWordLine 0 3 "Normal"}
									"High"   {WriteWordLine 0 3 "High"}
									Default {WriteWordLine 0 3 "Session importance could not be determined: $($Setting.SessionImportance.Value)"}
								}
							}
							If($Setting.SingleSignOn.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Session Settings\Single Sign-On: " $Setting.SingleSignOn.State
							}
							If($Setting.SingleSignOnCentralStore.State -ne "NotConfigured")
							{
								WriteWordLine 0 2 "Server Session Settings\Single Sign-On central store: " $Setting.SingleSignOnCentralStore.Value
							}
						}
					}
					WriteWordLine 0 0 ""
				}
				Else
				{
					WriteWordLine 0 1 "Unable to retrieve settings"
				}
				$Filter = $null
				$Settings = $null
			}
		}
		Else 
		{
			Write-warning "Citrix Policy information could not be retrieved."
		}
			
		$Policies = $null
	}
}

write-verbose "Finishing up Word document"
#end of document processing
#Update document properties

If($CoverPagesExist)
{
	write-verbose "Set Cover Page Properties"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Company" $CompanyName
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Title" $title
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Subject" "XenApp 6.5 Farm Inventory"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Author" $username

	#Get the Coverpage XML part
	$cp=$doc.CustomXMLParts | where {$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab=$cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
	#set the text
	[string]$abstract="Citrix XenApp 6.5 Inventory for $CompanyName"
	$ab.Text=$abstract

	$ab=$cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
	#set the text
	[string]$abstract=( Get-Date -Format d ).ToString()
	$ab.Text=$abstract

	write-verbose "Update the Table of Contents"
	#update the Table of Contents
	$doc.TablesOfContents.item(1).Update()
}

write-verbose "Save and Close document and Shutdown Word"
If ($WordVersion -eq 12)
{
	#Word 2007
	$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
	$doc.SaveAs($filename, $SaveFormat)
}
Else
{
	#the $saveFormat below passes StrictMode 2
	#I found this at the following two links
	#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
	#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
	$doc.SaveAs([REF]$filename, [ref]$SaveFormat)
}

$doc.Close()
$Word.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
Remove-Variable -Name word
[gc]::collect() 
[gc]::WaitForPendingFinalizers()