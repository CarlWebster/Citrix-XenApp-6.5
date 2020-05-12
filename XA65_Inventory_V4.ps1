#This File is in Unicode format.  Do not edit in an ASCII editor.

<#
.SYNOPSIS
	Creates a complete inventory of a Citrix XenApp 6.5 farm using Microsoft Word.
.DESCRIPTION
	Creates a complete inventory of a Citrix XenApp 6.5 farm using Microsoft Word and PowerShell.
	Creates either a Word document or PDF named after the XenApp 6.5 farm.
	Document includes a Cover Page, Table of Contents and Footer.
	Version 4 includes support for the following language versions of Microsoft Word:
		Catalan
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
.PARAMETER CoverPage
	What Microsoft Word Cover Page to use.
	(Default cover pages in Word en-US)
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
	Default value is Sideline.
	This parameter has an alias of CP.
.PARAMETER UserName
	User name to use for the Cover Page and Footer.
	Default value is contained in $env:username
	This parameter has an alias of UN.
.PARAMETER PDF
	SaveAs PDF file instead of DOCX file.
	This parameter is disabled by default.
	For Word 2007, the Microsoft add-in for saving as a PDF muct be installed.
	For Word 2007, please see http://www.microsoft.com/en-us/download/details.aspx?id=9943
	The PDF file is roughly 5X to 10X larger than the DOCX file.
.PARAMETER Hardware
	Use WMI to gather hardware information on: Computer System, Disks, Processor and Network Interface Cards
	This parameter is disabled by default.
.PARAMETER Software
	Gather software installed by querying the registry.  
	Use SoftwareExclusions.txt to exclude software from the report.
	SoftwareExclusions.txt must exist, and be readable, in the same folder as this script.
	SoftwareExclusions.txt can be an empty file to have no installed applications excluded.
	See Get-Help About-Wildcards for help on formatting the lines to exclude applications.
	This parameter is disabled by default.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V4.ps1
	
	Will use all Default values.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V4.ps1 -PDF -verbose
	
	Will use all Default values and save the document as a PDF file.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will display verbose messages as the script is running.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V4.ps1 -Hardware -verbose
	
	Will use all Default values and add additional information for each server about its hardware.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will display verbose messages as the script is running.
.EXAMPLE
	PS C:\PSScript > .\XA65_Inventory_V4.ps1 -Software -verbose
	
	Will use all Default values and add additional information for each server about its installed applications.
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\CompanyName="Carl Webster" or
	HKEY_CURRENT_USER\Software\Microsoft\Office\Common\UserInfo\Company="Carl Webster"
	$env:username = Administrator

	Carl Webster for the Company Name.
	Sideline for the Cover Page format.
	Administrator for the User Name.
	Will display verbose messages as the script is running.
.EXAMPLE
	PS C:\PSScript .\XA65_Inventory_V4.ps1 -CompanyName "Carl Webster Consulting" -CoverPage "Mod" -UserName "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name.
		Mod for the Cover Page format.
		Carl Webster for the User Name.
.EXAMPLE
	PS C:\PSScript .\XA65_Inventory_V4.ps1 -CN "Carl Webster Consulting" -CP "Mod" -UN "Carl Webster"

	Will use:
		Carl Webster Consulting for the Company Name (alias CN).
		Mod for the Cover Page format (alias CP).
		Carl Webster for the User Name (alias UN).
.INPUTS
	None.  You cannot pipe objects to this script.
.OUTPUTS
	No objects are output from this script.  This script creates a Word document.
.NOTES
	NAME: XA65_Inventory_V4.ps1
	VERSION: 4.02
	AUTHOR: Carl Webster (with a lot of help from Michael B. Smith and Jeff Wouters)
	LASTEDIT: December 5, 2013
#>


#thanks to @jeffwouters for helping me with these parameters
[CmdletBinding(SupportsShouldProcess = $False, ConfirmImpact = "None", DefaultParameterSetName = "") ]

Param([parameter(
	Position = 0, 
	Mandatory=$False)
	] 
	[Alias("CN")]
	[ValidateNotNullOrEmpty()]
	[string]$CompanyName="",
    
	[parameter(
	Position = 1, 
	Mandatory=$False)
	] 
	[Alias("CP")]
	[ValidateNotNullOrEmpty()]
	[string]$CoverPage="Sideline", 

	[parameter(
	Position = 2, 
	Mandatory=$False)
	] 
	[Alias("UN")]
	[ValidateNotNullOrEmpty()]
	[string]$UserName=$env:username,

	[parameter(
	Position = 3, 
	Mandatory=$False)
	] 
	[Switch]$PDF=$False,

	[parameter(
	Position = 4, 
	Mandatory=$False)
	] 
	[Switch]$Hardware=$False, 

	[parameter(
	Position = 5, 
	Mandatory=$False)
	] 
	[Switch]$Software=$False 
	)

#force -verbose on
$PSDefaultParameterValues = @{"*:Verbose"=$True}
	
#Original Script created 8/17/2010 by Michael Bogobowicz, Citrix Systems.
#To contact, please message @mikebogo on Twitter
#This script is designed to be run on a XenApp 6.5 server

#Modifications by Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#modified from the original script for XenApp 6.5
#	Version 4 of script is based on version 3.17 of XA65 script
#	Add ability to process AD based Citrix policies
#	Add Appendix A for Session Sharing information
#	Add Appendix B for Server Major Items
#	Add descriptions for Citrix Policy filter type
#	Add detecting the running Operating System to handle Word 2007 oddness with Server 2003/2008 vs Windows 7 vs Server 2008 R2
#	Add elapsed time to end of script
#	Add extra testing for applications, load balancing policies and worker groups to report if none exist instead of issuing a warning
#	Add get-date to all write-verbose statements
#	Add missing "None" option to ICA\Visual Display\Moving Images\Progressive compression level
#	Add more Write-Verbose statements
#	Add option to SaveAs PDF
#	Add setting Default tab stops at 36 points (1/2 inch in the USA)
#	Add Software Inventory
#	Add Summary Page
#	Add support for non-English versions of Microsoft Word
#	Add WMI hardware information for Computer System, Disks, Processor and Network Interface Cards
#	Change all instances of using $Word.Quit() to also use proper garbage collection
#	Change all occurrences of Access Session Conditions to Tables 
#	Change Default Cover Page to Sideline since Motion is not in German Word
#	Change Get-RegistryValue function to handle $null return value
#	Change most $Global: variables to regular variables
#	Change the test for the existence of XA65ConfigLog.udl from using .\ to $pwd.path
#	Change wording of not being able to load the Citrix.GroupPolicy.Commands.psm1 module
#	Change wording when script aborts from a blank company name
#	Fix issues with Word 2007 SaveAs under (Server 2008 and Windows 7) and Server 2008 R2
#	Fix logic error when comparing Citrix installed hotfixes to the recommended list
#	Fix output and missing items from ICA\Printing\Client Printers\Printer driver mapping and compatibility
#	Fix output of ICA\Adobe Flash Delivery\Flash Redirection\Flash URL compatibility list
#	Fix output of ICA\MultiStream Connections\Multi-Port Policy
#	Fix output of ICA\Printing\Drivers\Universal driver preference
#	Fix output of ICA\Printing\Session printers
#	Fix output of ICA\Printing\Universal Printing\Universal printing optimization defaults
#	Fix output of Server Settings\Health Monitoring and Recovery\Health monitoring tests
#	Fix WaitForPrintersToBeCreated policy setting
#	Fixing ICA\Printing\Session printers and ICA\Printing\Client Printers\Printer driver mapping and compatibility  required a new Function Get-PrinterModifiedSettings to keep from having duplicate code from Session Printers
#	Abort script if Farm information cannot be retrieved
#	Align Tables on Tab stop boundaries
#	Consolidated all the code to properly abort the script into a function AbortScript
#	Force the -verbose common parameter to be $True if running PoSH V3 and later
#	General code cleanup
#	If cover page selected does not exist, abort script
#	If running Word 2007 and the Save As PDF option is selected then verify the Save As PDF add-in is installed.  Abort script if not installed.
#	In the Server section, change Published Application to a Table
#	Load Balancing Policies: fixed display of "Apply to connections made through Access Gateway" and "Configure application connection preference based on worker group"
#	Only process WMI hardware information if the server is online
#	Strongly type all possible variables
#	Update for changes to CTX129229
#	Verify Get-HotFix cmdlet worked.  If not, write error and suggestion to document
#	Verify Word object is created.  If not, write error & suggestion to document and abort script
#Updated 07-Nov-2013
#	Changed link to Citrix.GroupPolicy.Commands.psm module to my Dropbox
#	Changed the GetCtxGPOsInAD function to work in a Windows Workgroup environment
#	The Hotfix array for Citrix hotfixes was not initialized correctly causing all installed Citrix hotfixes to show as not installed.
#	Removed the .LINK section from the help text
#Updated 12-Nov-2013
#	Added back in the French sections that somehow got removed
#Updated 5-Dec-2013
#	Fixed bug where XA65ConfigLog.udl was not found even if it existed
#	Fixed bug where the functions in Citrix.GroupPolicy.Command.psm1 were not found
#	Initialize switch parameters as $False

Set-StrictMode -Version 2

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
[int]$wdColorBlack = 0
[int]$wdWord2007 = 12
[int]$wdWord2010 = 14
[int]$wdWord2013 = 15
[int]$wdSaveFormatPDF = 17
[string]$RunningOS = (Get-WmiObject -class Win32_OperatingSystem).Caption

$hash = @{}

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

Switch ($PSUICulture.Substring(0,3))
{
	'ca-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Taula automática 2';
			}
		}

	'da-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabel 2';
			}
		}

	'de-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatische Tabelle 2';
			}
		}

	'en-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents'  = 'Automatic Table 2';
			}
		}

	'es-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Tabla automática 2';
			}
		}

	'fi-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automaattinen taulukko 2';
			}
		}

	'fr-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Sommaire Automatique 2';
			}
		}

	'nb-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatisk tabell 2';
			}
		}

	'nl-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatische inhoudsopgave 2';
			}
		}

	'pt-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Sumário Automático 2';
			}
		}

	'sv-'	{
			$hash.($($PSUICulture)) = @{
				'Word_TableOfContents' = 'Automatisk innehållsförteckning2';
			}
		}

	Default	{$hash.('en-US') = @{
				'Word_TableOfContents'  = 'Automatic Table 2';
			}
		}
}

# http://www.thedoctools.com/index.php?show=wt_style_names_english_danish_german_french
[int]$wdStyleHeading1 = -2
[int]$wdStyleHeading2 = -3
[int]$wdStyleHeading3 = -4
[int]$wdStyleHeading4 = -5
[int]$wdStyleNoSpacing = -158
[int]$wdTableGrid = -155

$myHash = $hash.$PSUICulture

If($myHash -eq $Null)
{
	$myHash = $hash.('en-US')
}

$myHash.Word_NoSpacing = $wdStyleNoSpacing
$myHash.Word_Heading1 = $wdStyleheading1
$myHash.Word_Heading2 = $wdStyleheading2
$myHash.Word_Heading3 = $wdStyleheading3
$myHash.Word_Heading4 = $wdStyleheading4
$myHash.Word_TableGrid = $wdTableGrid

Function ValidateCoverPage
{
	Param([int]$xWordVersion, [string]$xCP)
	
	$xArray = ""
	
	Switch ($PSUICulture.Substring(0,3))
	{
		'ca-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Anual", "Conservador", "Contrast",
					"Cubicles", "Diplomàtic", "En mosaic", "Exposició", "Línia lateral",
					"Mod", "Moviment", "Piles", "Sobri", "Transcendir", "Trencaclosques")
				}
			}

		'da-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Årlig", "BevægElse", "Eksponering",
					"Enkel", "Firkanter", "Fliser", "Gåde", "Kontrast",
					"Mod", "Nålestribet", "Overskrid", "Sidelinje", "Stakke",
					"Tradionel")
				}
			}

		'de-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Bewegung", "Durchscheinend", "Herausgestellt",
					"Jährlich", "Kacheln", "Kontrast", "Kubistisch", "Modern",
					"Nadelstreifen", "Puzzle", "Randlinie", "Raster", "Schlicht", "Stapel",
					"Traditionell")
				}
			}

		'en-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
					"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
					"Sideline", "Stacks", "Tiles", "Transcend")
				}
			}

		'es-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Anual", "Austero", "Conservador",
					"Contraste", "Cubículos", "Exposición", "Línea lateral",
					"Moderno", "Mosaicos", "Movimiento", "Pilas", "Puzzle",
					"Rayas", "Sobrepasar")
				}
			}

		'fi-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aakkoset", "Alttius", "Kontrasti", "Kuvakkeet ja tiedot",
					"Liike" , "Liituraita" , "Mod" , "Palapeli", "Perinteinen", "Pinot",
					"Sivussa", "Työpisteet", "Vuosittainen", "Yksinkertainen", "Ylitys")
				}
			}

		'fr-'	{
				If($xWordVersion -eq $wdWord2013)
				{
					$xArray = ("ViewMaster", "Secteur (foncé)", "Sémaphore",
					"Rétrospective", "Ion (foncé)", "Ion (clair)", "Intégrale",
					"Filigrane", "Facette", "Secteur (clair)", "À bandes", "Austin",
					"Guide", "Whisp", "Lignes latérales", "Quadrillage")
				}
				ElseIf($xWordVersion -eq $wdWord2010)
				{
					$xArray = ("Mosaïques", "Ligne latérale", "Annuel", "Perspective",
					"Contraste", "Emplacements de bureau", "Moderne", "Blocs empilés",
					"Rayures fines", "Austère", "Transcendant", "Classique", "Quadrillage",
					"Exposition", "Alphabet", "Mots croisés", "Papier journal", "Austin", "Guide")
				}
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alphabet", "Annuel", "Austère", "Blocs empilés", "Blocs superposés",
					"Classique", "Contraste", "Exposition", "Guide", "Ligne latérale", "Moderne",
					"Mosaïques", "Mots croisés", "Rayures fines", "Transcendant")
				}
			}

		'nb-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabet", "Årlig", "Avlukker", "BevegElse", "Engasjement",
					"Enkel", "Fliser", "Konservativ", "Kontrast", "Mod", "Puslespill",
					"Sidelinje", "Smale striper", "Stabler", "Transcenderende")
				}
			}

		'nl-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Aantrekkelijk", "Alfabet", "Bescheiden", "Beweging",
					"Blikvanger", "Contrast", "Eenvoudig", "Jaarlijks", "Krijtstreep",
					"Mod", "Puzzel", "Stapels", "Tegels", "Terzijde", "Werkplekken")
				}
			}

		'pt-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabeto", "Animação", "Anual", "Austero", "Baias", "Conservador",
					"Contraste", "Exposição", "Ladrilhos", "Linha Lateral", "Listras", "Mod",
					"Pilhas", "Quebra-cabeça", "Transcendente")
				}
			}

		'sv-'	{
				If($xWordVersion -eq $wdWord2013)
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
				ElseIf($xWordVersion -eq $wdWord2007)
				{
					$xArray = ("Alfabetmönster", "Årligt", "Enkelt", "Exponering", "Konservativt",
					"Kontrast", "Kritstreck", "Kuber", "Övergående", "Plattor", "Pussel", "RörElse",
					"Sidlinje", "Sobert", "Staplat")
				}
			}

		Default	{
					If($xWordVersion -eq $wdWord2013)
					{
						$xArray = ("Austin", "Banded", "Facet", "Filigree", "Grid", "Integral",
						"Ion (Dark)", "Ion (Light)", "Motion", "Retrospect", "Semaphore",
						"Sideline", "Slice (Dark)", "Slice (Light)", "ViewMaster", "Whisp")
					}
					ElseIf($xWordVersion -eq $wdWord2010)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Austin", "Conservative",
						"Contrast", "Cubicles", "Exposure", "Grid", "Mod", "Motion", "Newsprint",
						"Perspective", "Pinstripes", "Puzzle", "Sideline", "Stacks", "Tiles", "Transcend")
					}
					ElseIf($xWordVersion -eq $wdWord2007)
					{
						$xArray = ("Alphabet", "Annual", "Austere", "Conservative", "Contrast",
						"Cubicles", "Exposure", "Mod", "Motion", "Pinstripes", "Puzzle",
						"Sideline", "Stacks", "Tiles", "Transcend")
					}
				}
	}
	
	If($xArray -contains $xCP)
	{
		Return $True
	}
	Else
	{
		Return $False
	}
}

Function GetComputerWMIInfo
{
	Param([string]$RemoteComputerName)
	
	# original work by Kees Baggerman, 
	# Senior Technical Consultant @ Inter Access
	# k.baggerman@myvirtualvision.com
	# @kbaggerman on Twitter
	# http://blog.myvirtualvision.com

	#Get Computer info
	Write-Verbose "$(Get-Date): `t`tProcessing WMI Computer information"
	Write-Verbose "$(Get-Date): `t`t`tHardware information"
	WriteWordLine 3 0 "Computer Information"
	WriteWordLine 0 1 "General Computer"
	[bool]$GotComputerItems = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_computersystem
		$ComputerItems = $Results | Select Manufacturer, Model, Domain, @{N="TotalPhysicalRam"; E={[math]::round(($_.TotalPhysicalMemory / 1GB),0)}}
		$Results = $Null
	}
	
	Catch
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		$GotComputerItems = $False
		Write-Warning "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "Get-WmiObject win32_computersystem failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
	}
	
	If($GotComputerItems)
	{
		ForEach($Item in $ComputerItems)
		{
			WriteWordLine 0 2 "Manufacturer`t: " $Item.manufacturer
			WriteWordLine 0 2 "Model`t`t: " $Item.model
			WriteWordLine 0 2 "Domain`t`t: " $Item.domain
			WriteWordLine 0 2 "Total Ram`t: $($Item.totalphysicalram) GB"
			WriteWordLine 0 2 ""
		}
	}

	#Get Disk info
	Write-Verbose "$(Get-Date): `t`t`tDrive information"
	WriteWordLine 0 1 "Drive(s)"
	[bool]$GotDrives = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName Win32_LogicalDisk
		$drives = $Results | select caption, @{N="drivesize"; E={[math]::round(($_.size / 1GB),0)}}, 
		filesystem, @{N="drivefreespace"; E={[math]::round(($_.freespace / 1GB),0)}}, 
		volumename, drivetype, volumedirty, volumeserialnumber
		$Results = $Null
	}
	
	Catch
	{
		Write-Verbose "$(Get-Date): Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		$GotDrives = $False
		Write-Warning "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "Get-WmiObject Win32_LogicalDisk failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
	}
	
	If($GotDrives)
	{
		ForEach($drive in $drives)
		{
			If($drive.caption -ne "A:" -and $drive.caption -ne "B:")
			{
				WriteWordLine 0 2 "Caption`t`t: " $drive.caption
				WriteWordLine 0 2 "Size`t`t: $($drive.drivesize) GB"
				If(![String]::IsNullOrEmpty($drive.filesystem))
				{
					WriteWordLine 0 2 "File System`t: " $drive.filesystem
				}
				WriteWordLine 0 2 "Free Space`t: $($drive.drivefreespace) GB"
				If(![String]::IsNullOrEmpty($drive.volumename))
				{
					WriteWordLine 0 2 "Volume Name`t: " $drive.volumename
				}
				If(![String]::IsNullOrEmpty($drive.volumedirty))
				{
					WriteWordLine 0 2 "Volume is Dirty`t: " -nonewline
					If($drive.volumedirty)
					{
						WriteWordLine 0 0 "Yes"
					}
					Else
					{
						WriteWordLine 0 0 "No"
					}
				}
				If(![String]::IsNullOrEmpty($drive.volumeserialnumber))
				{
					WriteWordLine 0 2 "Volume Serial #`t: " $drive.volumeserialnumber
				}
				WriteWordLine 0 2 "Drive Type`t: " -nonewline
				Switch ($drive.drivetype)
				{
					0	{WriteWordLine 0 0 "Unknown"}
					1	{WriteWordLine 0 0 "No Root Directory"}
					2	{WriteWordLine 0 0 "Removable Disk"}
					3	{WriteWordLine 0 0 "Local Disk"}
					4	{WriteWordLine 0 0 "Network Drive"}
					5	{WriteWordLine 0 0 "Compact Disc"}
					6	{WriteWordLine 0 0 "RAM Disk"}
					Default {WriteWordLine 0 0 "Unknown"}
				}
				WriteWordLine 0 2 ""
			}
		}
	}

	#Get CPU's and stepping
	Write-Verbose "$(Get-Date): `t`t`tProcessor information"
	WriteWordLine 0 1 "Processor(s)"
	[bool]$GotProcessors = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_Processor
		$Processors = $Results | select availability, name, description, maxclockspeed, 
		l2cachesize, l3cachesize, numberofcores, numberoflogicalprocessors
		$Results = $Null
	}
	
	Catch
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		$GotProcessors = $False
		Write-Warning "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "Get-WmiObject win32_Processor failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
	}
	
	If($GotProcessors)
	{
		ForEach($processor in $processors)
		{
			WriteWordLine 0 2 "Name`t`t`t: " $processor.name
			WriteWordLine 0 2 "Description`t`t: " $processor.description
			WriteWordLine 0 2 "Max Clock Speed`t: $($processor.maxclockspeed) MHz"
			If($processor.l2cachesize -gt 0)
			{
				WriteWordLine 0 2 "L2 Cache Size`t`t: $($processor.l2cachesize) KB"
			}
			If($processor.l3cachesize -gt 0)
			{
				WriteWordLine 0 2 "L3 Cache Size`t`t: $($processor.l3cachesize) KB"
			}
			If($processor.numberofcores -gt 0)
			{
				WriteWordLine 0 2 "# of Cores`t`t: " $processor.numberofcores
			}
			If($processor.numberoflogicalprocessors -gt 0)
			{
				WriteWordLine 0 2 "# of Logical Procs`t: " $processor.numberoflogicalprocessors
			}
			WriteWordLine 0 2 "Availability`t`t: " -nonewline
			Switch ($processor.availability)
			{
				1	{WriteWordLine 0 0 "Other"}
				2	{WriteWordLine 0 0 "Unknown"}
				3	{WriteWordLine 0 0 "Running or Full Power"}
				4	{WriteWordLine 0 0 "Warning"}
				5	{WriteWordLine 0 0 "In Test"}
				6	{WriteWordLine 0 0 "Not Applicable"}
				7	{WriteWordLine 0 0 "Power Off"}
				8	{WriteWordLine 0 0 "Off Line"}
				9	{WriteWordLine 0 0 "Off Duty"}
				10	{WriteWordLine 0 0 "Degraded"}
				11	{WriteWordLine 0 0 "Not Installed"}
				12	{WriteWordLine 0 0 "Install Error"}
				13	{WriteWordLine 0 0 "Power Save - Unknown"}
				14	{WriteWordLine 0 0 "Power Save - Low Power Mode"}
				15	{WriteWordLine 0 0 "Power Save - Standby"}
				16	{WriteWordLine 0 0 "Power Cycle"}
				17	{WriteWordLine 0 0 "Power Save - Warning"}
				Default	{WriteWordLine 0 0 "Unknown"}
			}
			WriteWordLine 0 2 ""
		}
	}

	#Get Nics
	Write-Verbose "$(Get-Date): `t`t`tNIC information"
	WriteWordLine 0 1 "Network Interface(s)"
	[bool]$GotNics = $True
	
	Try
	{
		$Results = Get-WmiObject -computername $RemoteComputerName win32_networkadapterconfiguration 
		$Nics = $Results| where {$_.ipenabled -eq $True}
	}
	
	Catch
	{
		Write-Verbose "$(Get-Date): Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		$GotNics = $False
		Write-Warning "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "Get-WmiObject win32_networkadapterconfiguration failed for $($RemoteComputerName)"
		WriteWordLine 0 0 "On $($RemoteComputerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
	}
	
	If($GotNics)
	{
		ForEach($nic in $nics)
		{
			$ThisNic = Get-WmiObject -computername $RemoteComputerName win32_networkadapter | where {$_.index -eq $nic.index}
			If($ThisNic.Name -eq $nic.description)
			{
				WriteWordLine 0 2 "Name`t`t`t: " $ThisNic.Name
			}
			Else
			{
				WriteWordLine 0 2 "Name`t`t`t: " $ThisNic.Name
				WriteWordLine 0 2 "Description`t`t: " $nic.description
			}
			WriteWordLine 0 2 "Connection ID`t`t: " $ThisNic.NetConnectionID
			WriteWordLine 0 2 "Manufacturer`t`t: " $ThisNic.manufacturer
			WriteWordLine 0 2 "Availability`t`t: " -nonewline
			Switch ($ThisNic.availability)
			{
				1	{WriteWordLine 0 0 "Other"}
				2	{WriteWordLine 0 0 "Unknown"}
				3	{WriteWordLine 0 0 "Running or Full Power"}
				4	{WriteWordLine 0 0 "Warning"}
				5	{WriteWordLine 0 0 "In Test"}
				6	{WriteWordLine 0 0 "Not Applicable"}
				7	{WriteWordLine 0 0 "Power Off"}
				8	{WriteWordLine 0 0 "Off Line"}
				9	{WriteWordLine 0 0 "Off Duty"}
				10	{WriteWordLine 0 0 "Degraded"}
				11	{WriteWordLine 0 0 "Not Installed"}
				12	{WriteWordLine 0 0 "Install Error"}
				13	{WriteWordLine 0 0 "Power Save - Unknown"}
				14	{WriteWordLine 0 0 "Power Save - Low Power Mode"}
				15	{WriteWordLine 0 0 "Power Save - Standby"}
				16	{WriteWordLine 0 0 "Power Cycle"}
				17	{WriteWordLine 0 0 "Power Save - Warning"}
				Default	{WriteWordLine 0 0 "Unknown"}
			}
			WriteWordLine 0 2 "Physical Address`t: " $nic.macaddress
			WriteWordLine 0 2 "IP Address`t`t: " $nic.ipaddress
			WriteWordLine 0 2 "Default Gateway`t: " $nic.Defaultipgateway
			WriteWordLine 0 2 "Subnet Mask`t`t: " $nic.ipsubnet
			If($nic.dhcpenabled)
			{
				$DHCPLeaseObtainedDate = $nic.ConvertToDateTime($nic.dhcpleaseobtained)
				$DHCPLeaseExpiresDate = $nic.ConvertToDateTime($nic.dhcpleaseexpires)
				WriteWordLine 0 2 "DHCP Enabled`t`t: " $nic.dhcpenabled
				WriteWordLine 0 2 "DHCP Lease Obtained`t: " $dhcpleaseobtaineddate
				WriteWordLine 0 2 "DHCP Lease Expires`t: " $dhcpleaseexpiresdate
				WriteWordLine 0 2 "DHCP Server`t`t:" $nic.dhcpserver
			}
			If(![String]::IsNullOrEmpty($nic.dnsdomain))
			{
				WriteWordLine 0 2 "DNS Domain`t`t: " $nic.dnsdomain
			}
			If($nic.dnsdomainsuffixsearchorder -ne $Null -and $nic.dnsdomainsuffixsearchorder.length -gt 0)
			{
				[int]$x = 1
				WriteWordLine 0 2 "DNS Search Suffixes`t:" -nonewline
				ForEach($DNSDomain in $nic.dnsdomainsuffixsearchorder)
				{
					If($x -eq 1)
					{
						$x = 2
						WriteWordLine 0 0 " $($DNSDomain)"
					}
					Else
					{
						WriteWordLine 0 5 " $($DNSDomain)"
					}
				}
			}
			WriteWordLine 0 2 "DNS WINS Enabled`t: " -nonewline
			If($nic.dnsenabledforwinsresolution)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			If($nic.dnsserversearchorder -ne $Null -and $nic.dnsserversearchorder.length -gt 0)
			{
				[int]$x = 1
				WriteWordLine 0 2 "DNS Servers`t`t:" -nonewline
				ForEach($DNSServer in $nic.dnsserversearchorder)
				{
					If($x -eq 1)
					{
						$x = 2
						WriteWordLine 0 0 " $($DNSServer)"
					}
					Else
					{
						WriteWordLine 0 5 " $($DNSServer)"
					}
				}
			}
			WriteWordLine 0 2 "NetBIOS Setting`t`t: " -nonewline
			Switch ($nic.TcpipNetbiosOptions)
			{
				0	{WriteWordLine 0 0 "Use NetBIOS setting from DHCP Server"}
				1	{WriteWordLine 0 0 "Enable NetBIOS"}
				2	{WriteWordLine 0 0 "Disable NetBIOS"}
				Default	{WriteWordLine 0 0 "Unknown"}
			}
			WriteWordLine 0 2 "WINS:"
			WriteWordLine 0 3 "Enabled LMHosts`t: " -nonewline
			If($nic.winsenablelmhostslookup)
			{
				WriteWordLine 0 0 "Yes"
			}
			Else
			{
				WriteWordLine 0 0 "No"
			}
			If(![String]::IsNullOrEmpty($nic.winshostlookupfile))
			{
				WriteWordLine 0 3 "Host Lookup File`t: " $nic.winshostlookupfile
			}
			If(![String]::IsNullOrEmpty($nic.winsprimaryserver))
			{
				WriteWordLine 0 3 "Primary Server`t`t: " $nic.winsprimaryserver
			}
			If(![String]::IsNullOrEmpty($nic.winssecondaryserver))
			{
				WriteWordLine 0 3 "Secondary Server`t: " $nic.winssecondaryserver
			}
			If(![String]::IsNullOrEmpty($nic.winsscopeid))
			{
				WriteWordLine 0 3 "Scope ID`t`t: " $nic.winsscopeid
			}
		}
	}
	WriteWordLine 0 0 ""
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
	$var = ""
	$Tmp = '$InstalledApps | Where-Object {'
	$Exclusions = Get-Content "$($pwd.path)\SoftwareExclusions.txt" -EA 0
	If($Exclusions -ne $Null -and $Exclusions.Count -gt 0)
	{
		ForEach($Exclusion in $Exclusions) 
		{
			$Tmp += "(`$`_.DisplayName -notlike ""$($Exclusion)"") -and "
		}
		$var += $Tmp.Substring(0,($Tmp.Length - 6))
		$var += "} | Select-Object DisplayName | Sort DisplayName -unique"
	}
	return $var
}

Function CheckWordPrereq
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Word.Application) -eq $False)
	{
		Write-Host "This script directly outputs to Microsoft Word, please install Microsoft Word"
		exit
	}

	#find out our session (usually "1" except on TS/RDC or Citrix)
	$SessionID = (Get-Process -PID $PID).SessionId
	
	#Find out if winword is running in our session
	[bool]$wordrunning = ((Get-Process 'WinWord' -ea 0)|?{$_.SessionId -eq $SessionID}) -ne $Null
	If($wordrunning)
	{
		Write-Host "Please close all instances of Microsoft Word before running this report."
		exit
	}
}

Function CheckWord2007SaveAsPDFInstalled
{
	If((Test-Path  REGISTRY::HKEY_CLASSES_ROOT\Installer\Products\000021090B0090400000000000F01FEC) -eq $False)
	{
		Write-Host "Word 2007 is detected and the option to SaveAs PDF was selected but the Word 2007 SaveAs PDF add-in is not installed."
		Write-Host "The add-in can be downloaded from http://www.microsoft.com/en-us/download/details.aspx?id=9943"
		Write-Host "Install the SaveAs PDF add-in and rerun the script."
		Return $False
	}
	Return $True
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

Function MultiPortPolicyPriority
{
	Param([int]$PriorityValue = 3)

	Switch ($PriorityValue)
	{ 
		0 {"Very High"} 
		1 {"High"} 
		2 {"Medium"} 
		3 {"Low"} 
		Default {"Unknown Priority Value"}
	}
	Return $PriorityValue
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
	#$LoadedModules = Get-Module | Select Name
	#following line changed at the recommendation of @andyjmorgan
	$LoadedModules = Get-Module |% { $_.Name.ToString() }
	#bug reported on 21-JAN-2013 by @schose 
	#the following line did not work if the citrix.grouppolicy.commands.psm1 module
	#was manually loaded from a non Default folder
	#$ModuleFound = (!$LoadedModules -like "*$ModuleName*")
	
	[bool]$ModuleFound = ($LoadedModules -like "*$ModuleName*")
	If(!$ModuleFound) 
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

Function Check-NeededPSSnapins
{
	Param([parameter(Mandatory = $True)][alias("Snapin")][string[]]$Snapins)

	#Function specifics
	$MissingSnapins = @()
	[bool]$FoundMissingSnapin = $False
	$LoadedSnapins = @()
	$RegisteredSnapins = @()

	#Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
	$loadedSnapins += get-pssnapin | % {$_.name}
	$registeredSnapins += get-pssnapin -Registered | % {$_.name}

	ForEach($Snapin in $Snapins)
	{
		#check if the snapin is loaded
		If(!($LoadedSnapins -like $snapin))
		{
			#Check if the snapin is missing
			If(!($RegisteredSnapins -like $Snapin))
			{
				#set the flag if it's not already
				If(!($FoundMissingSnapin))
				{
					$FoundMissingSnapin = $True
				}
				#add the entry to the list
				$MissingSnapins += $Snapin
			}
			Else
			{
				#Snapin is registered, but not loaded, loading it now:
				Write-Host "Loading Windows PowerShell snap-in: $snapin"
				Add-PSSnapin -Name $snapin -EA 0
			}
		}
	}

	If($FoundMissingSnapin)
	{
		Write-Warning "Missing Windows PowerShell snap-ins Detected:"
		$missingSnapins | % {Write-Warning "($_)"}
		return $False
	}
	Else
	{
		Return $True
	}
}

Function WriteWordLine
#Function created by Ryan Revord
#@rsrevord on Twitter
#Function created to make output to Word easy in this script
{
	Param([int]$style=0, [int]$tabs = 0, [string]$name = '', [string]$value = '', [string]$newline = "'n", [Switch]$nonewline)
	[string]$output = ""
	#Build output style
	Switch ($style)
	{
		0 {$Selection.Style = $myHash.Word_NoSpacing}
		1 {$Selection.Style = $myHash.Word_Heading1}
		2 {$Selection.Style = $myHash.Word_Heading2}
		3 {$Selection.Style = $myHash.Word_Heading3}
		4 {$Selection.Style = $myHash.Word_Heading4}
		Default {$Selection.Style = $myHash.Word_NoSpacing}
	}
	#build # of tabs
	While($tabs -gt 0)
	{ 
		$output += "`t"; $tabs--; 
	}
		
	#output the rest of the parameters.
	$output += $name + $value
	$Selection.TypeText($output)
	
	#test for new WriteWordLine 0.
	If($nonewline)
	{
		# Do nothing.
	} 
	Else 
	{
		$Selection.TypeParagraph()
	}
}

Function _SetDocumentProperty 
{
	#jeff hicks
	Param([object]$Properties,[string]$Name,[string]$Value)
	#get the property object
	$prop = $properties | ForEach { 
		$propname=$_.GetType().InvokeMember("Name","GetProperty",$Null,$_,$Null)
		If($propname -eq $Name) 
		{
			Return $_
		}
	} #ForEach

	#set the value
	$Prop.GetType().InvokeMember("Value","SetProperty",$Null,$prop,$Value)
}

Function Get-PrinterModifiedSettings
{
	Param([string]$Value, [string]$xelement, [bool]$xtable)
	
	[string]$ReturnStr = ""

	Switch ($Value)
	{
		"copi" 
		{
			If($xtable)
			{
				$txt="Copies:"
			}
			Else
			{
				$txt="Copies`t`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"coll"
		{
			If($xtable)
			{
				$txt="Collate:"
			}
			Else
			{
				$txt="Collate`t`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"scal"
		{
			If($xtable)
			{
				$txt="Scale (%):"
			}
			Else
			{
				$txt="Scale (%)`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"colo"
		{
			If($xtable)
			{
				$txt="Color:"
			}
			Else
			{
				$txt="Color`t`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Monochrome"}
					2 {$tmp2 = "Color"}
					Default {$tmp2 = "Color could not be determined: $($xelement)"}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"prin"
		{
			If($xtable)
			{
				$txt="Print Quality:"
			}
			Else
			{
				$txt="Print Quality`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					-1 {$tmp2 = "150 dpi"}
					-2 {$tmp2 = "300 dpi"}
					-3 {$tmp2 = "600 dpi"}
					-4 {$tmp2 = "1200 dpi"}
					Default 
					{
						$tmp2 = "Custom...`tX resolution: $tmp1"
					}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"yres"
		{
			If($xtable)
			{
				$txt="Y resolution:"
			}
			Else
			{
				$txt="Y resolution`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"orie"
		{
			If($xtable)
			{
				$txt="Orientation:"
			}
			Else
			{
				$txt="Orientation`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					"portrait"  {$tmp2 = "Portrait"}
					"landscape" {$tmp2 = "Landscape"}
					Default {$tmp2 = "Orientation could not be determined: $($xelement)"}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"dupl"
		{
			If($xtable)
			{
				$txt="Duplex:"
			}
			Else
			{
				$txt="Duplex`t`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Simplex"}
					2 {$tmp2 = "Vertical"}
					3 {$tmp2 = "Horizontal"}
					Default {$tmp2 = "Duplex could not be determined: $($xelement)"}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"pape"
		{
			If($xtable)
			{
				$txt="Paper Size:"
			}
			Else
			{
				$txt="Paper Size`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1   {$tmp2 = "Letter"}
					2   {$tmp2 = "Letter Small"}
					3   {$tmp2 = "Tabloid"}
					4   {$tmp2 = "Ledger"}
					5   {$tmp2 = "Legal"}
					6   {$tmp2 = "Statement"}
					7   {$tmp2 = "Executive"}
					8   {$tmp2 = "A3"}
					9   {$tmp2 = "A4"}
					10  {$tmp2 = "A4 Small"}
					11  {$tmp2 = "A5"}
					12  {$tmp2 = "B4 (JIS)"}
					13  {$tmp2 = "B5 (JIS)"}
					14  {$tmp2 = "Folio"}
					15  {$tmp2 = "Quarto"}
					16  {$tmp2 = "10X14"}
					17  {$tmp2 = "11X17"}
					18  {$tmp2 = "Note"}
					19  {$tmp2 = "Envelope #9"}
					20  {$tmp2 = "Envelope #10"}
					21  {$tmp2 = "Envelope #11"}
					22  {$tmp2 = "Envelope #12"}
					23  {$tmp2 = "Envelope #14"}
					24  {$tmp2 = "C Size Sheet"}
					25  {$tmp2 = "D Size Sheet"}
					26  {$tmp2 = "E Size Sheet"}
					27  {$tmp2 = "Envelope DL"}
					28  {$tmp2 = "Envelope C5"}
					29  {$tmp2 = "Envelope C3"}
					30  {$tmp2 = "Envelope C4"}
					31  {$tmp2 = "Envelope C6"}
					32  {$tmp2 = "Envelope C65"}
					33  {$tmp2 = "Envelope B4"}
					34  {$tmp2 = "Envelope B5"}
					35  {$tmp2 = "Envelope B6"}
					36  {$tmp2 = "Envelope Italy"}
					37  {$tmp2 = "Envelope Monarch"}
					38  {$tmp2 = "Envelope Personal"}
					39  {$tmp2 = "US Std Fanfold"}
					40  {$tmp2 = "German Std Fanfold"}
					41  {$tmp2 = "German Legal Fanfold"}
					42  {$tmp2 = "B4 (ISO)"}
					43  {$tmp2 = "Japanese Postcard"}
					44  {$tmp2 = "9X11"}
					45  {$tmp2 = "10X11"}
					46  {$tmp2 = "15X11"}
					47  {$tmp2 = "Envelope Invite"}
					48  {$tmp2 = "Reserved - DO NOT USE"}
					49  {$tmp2 = "Reserved - DO NOT USE"}
					50  {$tmp2 = "Letter Extra"}
					51  {$tmp2 = "Legal Extra"}
					52  {$tmp2 = "Tabloid Extra"}
					53  {$tmp2 = "A4 Extra"}
					54  {$tmp2 = "Letter Transverse"}
					55  {$tmp2 = "A4 Transverse"}
					56  {$tmp2 = "Letter Extra Transverse"}
					57  {$tmp2 = "A Plus"}
					58  {$tmp2 = "B Plus"}
					59  {$tmp2 = "Letter Plus"}
					60  {$tmp2 = "A4 Plus"}
					61  {$tmp2 = "A5 Transverse"}
					62  {$tmp2 = "B5 (JIS) Transverse"}
					63  {$tmp2 = "A3 Extra"}
					64  {$tmp2 = "A5 Extra"}
					65  {$tmp2 = "B5 (ISO) Extra"}
					66  {$tmp2 = "A2"}
					67  {$tmp2 = "A3 Transverse"}
					68  {$tmp2 = "A3 Extra Transverse"}
					69  {$tmp2 = "Japanese Double Postcard"}
					70  {$tmp2 = "A6"}
					71  {$tmp2 = "Japanese Envelope Kaku #2"}
					72  {$tmp2 = "Japanese Envelope Kaku #3"}
					73  {$tmp2 = "Japanese Envelope Chou #3"}
					74  {$tmp2 = "Japanese Envelope Chou #4"}
					75  {$tmp2 = "Letter Rotated"}
					76  {$tmp2 = "A3 Rotated"}
					77  {$tmp2 = "A4 Rotated"}
					78  {$tmp2 = "A5 Rotated"}
					79  {$tmp2 = "B4 (JIS) Rotated"}
					80  {$tmp2 = "B5 (JIS) Rotated"}
					81  {$tmp2 = "Japanese Postcard Rotated"}
					82  {$tmp2 = "Double Japanese Postcard Rotated"}
					83  {$tmp2 = "A6 Rotated"}
					84  {$tmp2 = "Japanese Envelope Kaku #2 Rotated"}
					85  {$tmp2 = "Japanese Envelope Kaku #3 Rotated"}
					86  {$tmp2 = "Japanese Envelope Chou #3 Rotated"}
					87  {$tmp2 = "Japanese Envelope Chou #4 Rotated"}
					88  {$tmp2 = "B6 (JIS)"}
					89  {$tmp2 = "B6 (JIS) Rotated"}
					90  {$tmp2 = "12X11"}
					91  {$tmp2 = "Japanese Envelope You #4"}
					92  {$tmp2 = "Japanese Envelope You #4 Rotated"}
					93  {$tmp2 = "PRC 16K"}
					94  {$tmp2 = "PRC 32K"}
					95  {$tmp2 = "PRC 32K(Big)"}
					96  {$tmp2 = "PRC Envelope #1"}
					97  {$tmp2 = "PRC Envelope #2"}
					98  {$tmp2 = "PRC Envelope #3"}
					99  {$tmp2 = "PRC Envelope #4"}
					100 {$tmp2 = "PRC Envelope #5"}
					101 {$tmp2 = "PRC Envelope #6"}
					102 {$tmp2 = "PRC Envelope #7"}
					103 {$tmp2 = "PRC Envelope #8"}
					104 {$tmp2 = "PRC Envelope #9"}
					105 {$tmp2 = "PRC Envelope #10"}
					106 {$tmp2 = "PRC 16K Rotated"}
					107 {$tmp2 = "PRC 32K Rotated"}
					108 {$tmp2 = "PRC 32K(Big) Rotated"}
					109 {$tmp2 = "PRC Envelope #1 Rotated"}
					110 {$tmp2 = "PRC Envelope #2 Rotated"}
					111 {$tmp2 = "PRC Envelope #3 Rotated"}
					112 {$tmp2 = "PRC Envelope #4 Rotated"}
					113 {$tmp2 = "PRC Envelope #5 Rotated"}
					114 {$tmp2 = "PRC Envelope #6 Rotated"}
					115 {$tmp2 = "PRC Envelope #7 Rotated"}
					116 {$tmp2 = "PRC Envelope #8 Rotated"}
					117 {$tmp2 = "PRC Envelope #9 Rotated"}
					Default {$tmp2 = "Paper Size could not be determined: $($xelement)"}
				}
				$ReturnStr = "$txt $tmp2"
			}
		}
		"form"
		{
			If($xtable)
			{
				$txt="Form Name:"
			}
			Else
			{
				$txt="Form Name`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
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
			If($xtable)
			{
				$txt="TrueType:"
			}
			Else
			{
				$txt="TrueType`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp1 = $xelement.SubString($index + 1)
				Switch ($tmp1)
				{
					1 {$tmp2 = "Bitmap"}
					2 {$tmp2 = "Download"}
					3 {$tmp2 = "Substitute"}
					4 {$tmp2 = "Outline"}
					Default {$tmp2 = "TrueType could not be determined: $($xelement)"}
				}
			}
			$ReturnStr = "$txt $tmp2"
		}
		"mode" 
		{
			If($xtable)
			{
				$txt="Printer Model:"
			}
			Else
			{
				$txt="Printer Model`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				$ReturnStr = "$txt $tmp2"
			}
		}
		"loca" 
		{
			If($xtable)
			{
				$txt="Location:"
			}
			Else
			{
				$txt="Location`t:"
			}
			$index = $xelement.SubString(0).IndexOf('=')
			if($index -ge 0)
			{
				$tmp2 = $xelement.SubString($index + 1)
				If($tmp2.length -gt 0)
				{
					$ReturnStr = "$txt $tmp2"
				}
			}
		}
		Default {$ReturnStr = "Session printer setting could not be determined: $($xelement)"}
	}
	Return $ReturnStr
}

Function AbortScript
{
	$Word.quit()
	Write-Verbose "$(Get-Date): System Cleanup"
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
	Remove-Variable -Name word -Scope Global -EA 0
	[gc]::collect() 
	[gc]::WaitForPendingFinalizers()
	Write-Verbose "$(Get-Date): Script has been aborted"
	Exit
}

Function ProcessCitrixPolicies
{
	Param([string]$xDriveName)

	If($xDriveName -eq "")
	{
		$Policies = Get-CtxGroupPolicy -EA 0 | Sort-Object Type,Priority
	}
	Else
	{
		$Policies = Get-CtxGroupPolicy -DriveName $xDriveName -EA 0 | Sort-Object Type,Priority
	}
	If($?)
	{
		ForEach($Policy in $Policies)
		{
			Write-Verbose "$(Get-Date): `tStarted $($Policy.PolicyName)`t$($Policy.Type)"
			WriteWordLine 2 0 $Policy.PolicyName
			If($xDriveName -eq "")
			{
				WriteWordLine 0 1 "IMA Farm based policy"
				$Global:TotalIMAPolicies++
			}
			Else
			{
				WriteWordLine 0 1 "Active Directory based policy"
				$Global:TotalADPolicies++
			}

			WriteWordLine 0 1 "Type`t`t: " $Policy.Type
			
			If($Policy.Type -eq "Computer")
			{
				$Global:TotalComputerPolicies++
			}
			Else
			{
				$Global:TotalUserPolicies++
			}
			
			If(![String]::IsNullOrEmpty($Policy.Description))
			{
				WriteWordLine 0 1 "Description`t: " $Policy.Description
			}
			WriteWordLine 0 1 "Enabled`t`t: " $Policy.Enabled
			WriteWordLine 0 1 "Priority`t`t: " $Policy.Priority

			If($xDriveName -eq "")
			{
				$filters = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName -EA 0
			}
			Else
			{
				$filters = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName -DriveName $xDriveName -EA 0
			}

			If($?)
			{
				If(![String]::IsNullOrEmpty($filters))
				{
					WriteWordLine 0 1 "Filter(s)`t`t:"
					ForEach($Filter in $Filters)
					{
						WriteWordLine 0 2 "Filter name`t: " $filter.FilterName
						WriteWordLine 0 2 "Filter type`t: " -nonewline
						Switch($filter.FilterType)
						{
							"User"           {WriteWordLine 0 0 "User or Group"}
							"WorkerGroup"    {WriteWordLine 0 0 "Worker Group"}
							"OU"             {WriteWordLine 0 0 "Organization Unit"}
							"ClientName"     {WriteWordLine 0 0 "Client Name"}
							"ClientIP"       {WriteWordLine 0 0 "Client IP Address"}
							"BranchRepeater" {WriteWordLine 0 0 "Branch Repeater"}
							"AccessControl"  {WriteWordLine 0 0 "Access Control"}
							Default {WriteWordLine 0 3 "Policy Filter Type could not be determined: $($filter.FilterType)"}
						}
						WriteWordLine 0 2 "Filter enabled`t: " $filter.Enabled
						WriteWordLine 0 2 "Filter mode`t: " $filter.Mode
						If(![String]::IsNullOrEmpty($filter.FilterValue))
						{
							WriteWordLine 0 2 "Filter value`t: " $filter.FilterValue
						}
						WriteWordLine 0 2 ""
					}
				}
				Else
				{
					WriteWordLine 0 1 "Filter(s)`t`t: None"
				}
			}
			Else
			{
				WriteWordLine 0 1 "Unable to retrieve Filter settings"
			}

			If($xDriveName -eq "")
			{
				$Settings = Get-CtxGroupPolicyConfiguration -PolicyName $Policy.PolicyName -EA 0
			}
			Else
			{
				$Settings = Get-CtxGroupPolicyConfiguration -PolicyName $Policy.PolicyName -DriveName $xDriveName -EA 0
			}
			
			If($?)
			{
				ForEach($Setting in $Settings)
				{
					If($Setting.Type -eq "Computer")
					{
						Write-Verbose "$(Get-Date): `t`tComputer settings"
						Write-Verbose "$(Get-Date): `t`t`tICA"
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
							Switch ($Setting.AutoClientReconnectLogging.Value)
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
							
							Switch ($Setting.DisplayDegradePreference.Value)
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
							Switch ($Setting.MaximumColorDepth.Value)
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
							Switch ($Setting.IcaKeepAlives.Value)
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
							WriteWordLine 0 3 "CGP default port" -nonewline 
							WriteWordLine 0 1 "priority: High"
							[string]$Tmp = $Setting.MultiPortPolicy.Value
							If($Tmp.Length -gt 0)
							{
								[string]$cgpport1 = $Tmp.substring(0, $Tmp.indexof(";"))
								[string]$cgpport2 = $Tmp.substring($cgpport1.length + 1 , $Tmp.indexof(";"))
								[string]$cgpport3 = $Tmp.substring((($cgpport1.length + 1)+($cgpport2.length + 1)) , $Tmp.indexof(";"))
								[string]$cgpport1priority = multiportpolicypriority $cgpport1.substring($cgpport1.length -1, 1)
								[string]$cgpport2priority = multiportpolicypriority $cgpport2.substring($cgpport2.length -1, 1)
								[string]$cgpport3priority = multiportpolicypriority $cgpport3.substring($cgpport3.length -1, 1)
								$cgpport1 = $cgpport1.substring(0, $cgpport1.indexof(","))
								$cgpport2 = $cgpport2.substring(0, $cgpport2.indexof(","))
								$cgpport3 = $cgpport3.substring(0, $cgpport3.indexof(","))
								WriteWordLine 0 3 "CGP port1: " $cgpport1 -nonewline 
								WriteWordLine 0 1 "priority: " -nonewline
								Switch ($cgpport1priority[0])
								{
									"V"	{WriteWordLine 0 0 "Very High"}
									"M"	{WriteWordLine 0 0 "Medium"}
									"L"	{WriteWordLine 0 0 "Low"}
									Default	{WriteWordLine 0 0 "Unknown"}
								}
								WriteWordLine 0 3 "CGP port2: " $cgpport2 -nonewline
								WriteWordLine 0 1 "priority: " -nonewline
								Switch ($cgpport2priority[0])
								{
									"V"	{WriteWordLine 0 0 "Very High"}
									"M"	{WriteWordLine 0 0 "Medium"}
									"L"	{WriteWordLine 0 0 "Low"}
									Default	{WriteWordLine 0 0 "Unknown"}
								}
								WriteWordLine 0 3 "CGP port3: " $cgpport3 -nonewline
								WriteWordLine 0 1 "priority: " -nonewline
								Switch ($cgpport3priority[0])
								{
									"V"	{WriteWordLine 0 0 "Very High"}
									"M"	{WriteWordLine 0 0 "Medium"}
									"L"	{WriteWordLine 0 0 "Low"}
									Default	{WriteWordLine 0 0 "Unknown"}
								}
								$cgpport1 = $Null
								$cgpport2 = $Null
								$cgpport3 = $Null
								$cgpport1priority = $Null
								$cgpport2priority = $Null
								$cgpport3priority = $Null
							}
							$Tmp = $Null
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
						Write-Verbose "$(Get-Date): `t`t`tLicensing"
						If($Setting.LicenseServerHostName.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Licensing\License server host name: " $Setting.LicenseServerHostName.Value
						}
						If($Setting.LicenseServerPort.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Licensing\License server port: " $Setting.LicenseServerPort.Value
						}
						Write-Verbose "$(Get-Date): `t`t`tPower and Capacity Management"
						If($Setting.FarmName.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Power and Capacity Management\Farm name: " $Setting.FarmName.Value
						}
						If($Setting.WorkloadName.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Power and Capacity Management\Workload name: " $Setting.WorkloadName.Value
						}
						Write-Verbose "$(Get-Date): `t`t`tServer Settings"
						If($Setting.ConnectionAccessControl.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Server Settings\Connection access control: "
							Switch ($Setting.ConnectionAccessControl.Value)
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
							Switch ($Setting.ProductModel.Value)
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
								Write-Verbose "$(Get-Date): `t`t`t`tCreate Table for HMR Test $($test.name)"
								$TableRange = $doc.Application.Selection.Range
								[int]$Columns = 2
								[int]$Rows = $test.attributes.count - 1
								$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
								$table.Style = $myHash.Word_TableGrid
								$table.Borders.InsideLineStyle = 0
								$table.Borders.OutsideLineStyle = 0
								[int]$xRow = 1
								$Table.Cell($xRow,1).Range.Text = "Name"
								$Table.Cell($xRow,2).Range.Text = $test.name
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "File Location"
								$Table.Cell($xRow,2).Range.Text = $test.file
								If($test.HasAttribute("arguments"))
								{
									$xRow++
									$Table.Cell($xRow,1).Range.Text = "Arguments"
									$Table.Cell($xRow,2).Range.Text = $test.arguments
								}
								If(![String]::IsNullOrEmpty($test.Description))
								{
									$xRow++
									$Table.Cell($xRow,1).Range.Text = "Description"
									$Table.Cell($xRow,2).Range.Text = $test.description
								}
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "Interval"
								$Table.Cell($xRow,2).Range.Text = $test.interval
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "Time-out"
								$Table.Cell($xRow,2).Range.Text = $test.timeout
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "Threshold"
								$Table.Cell($xRow,2).Range.Text = $test.threshold
								$xRow++
								$Table.Cell($xRow,1).Range.Text = "Recovery Action"
								Switch ($test.RecoveryAction)
								{
									"AlertOnly"                     {$Table.Cell($xRow,2).Range.Text = "Alert Only"}
									"RemoveServerFromLoadBalancing" {$Table.Cell($xRow,2).Range.Text = "Remove Server from load balancing"}
									"RestartIma"                    {$Table.Cell($xRow,2).Range.Text = "Restart IMA"}
									"ShutdownIma"                   {$Table.Cell($xRow,2).Range.Text = "Shutdown IMA"}
									"RebootServer"                  {$Table.Cell($xRow,2).Range.Text = "Reboot Server"}
									Default {$Table.Cell($xRow,2).Range.Text = "Recovery Action could not be determined: $($test.RecoveryAction)"}
								}

								$Table.Rows.SetLeftIndent(108,1)
								$table.AutoFitBehavior(1)

								#return focus back to document
								Write-Verbose "$(Get-Date): `t`t`t`t`tReturn focus back to document"
								$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

								#move to the end of the current document
								Write-Verbose "$(Get-Date): `t`t`t`t`tMove to the end of the current document"
								$selection.EndKey($wdStory,$wdMove) | Out-Null
							}
							$XML = $Null
							$xRow = $Null
							$Columns = $Null
							$Row = $Null
						}
						If($Setting.MaximumServersOfflinePercent.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Server Settings\Health Monitoring and Recovery\"
							WriteWordLine 0 3 "Max % of servers with logon control: " $Setting.MaximumServersOfflinePercent.Value
						}
						If($Setting.CpuManagementServerLevel.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Server Settings\Memory/CPU\CPU management server level: "
							Switch ($Setting.CpuManagementServerLevel.Value)
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
							ForEach($element in $array)
							{
								WriteWordLine 0 3 $element
							}
						}
						If($Setting.MemoryOptimizationIntervalType.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Server Settings\Memory/CPU\Memory optimization interval: " -nonewline
							Switch ($Setting.MemoryOptimizationIntervalType.Value)
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
							$tmp = $Null
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
							ForEach($element in $array)
							{
								WriteWordLine 0 3 $element
							}
							$array = $Null
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
							Switch ($Setting.RebootDisableLogOnTime.Value)
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
							$Tmp = $Null
						}
						If($Setting.RebootScheduleTime.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot schedule time: " -nonewline
							$tmp = ConvertNumberToTime $Setting.RebootScheduleTime.Value 						
							WriteWordLine 0 0 $Tmp
							$Tmp = $Null
						}
						If($Setting.RebootWarningInterval.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Server Settings\Reboot Behavior\Reboot warning interval: "
							Switch ($Setting.RebootWarningInterval.Value)
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
							Switch ($Setting.RebootWarningStartTime.Value)
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
						Write-Verbose "$(Get-Date): `t`t`tVirtual IP"
						If($Setting.FilterAdapterAddresses.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Virtual IP\Virtual IP adapter address filtering: " $Setting.FilterAdapterAddresses.State
						}
						If($Setting.EnhancedCompatibilityPrograms.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Virtual IP\Virtual IP compatibility programs list: " 
							$array = $Setting.EnhancedCompatibilityPrograms.Values
							ForEach($element in $array)
							{
								WriteWordLine 0 3 $element
							}
							$array = $Null
						}
						If($Setting.EnhancedCompatibility.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Virtual IP\Virtual IP enhanced compatibility: " $Setting.EnhancedCompatibility.State
						}
						If($Setting.FilterAdapterAddressesPrograms.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Virtual IP\Virtual IP filter adapter addresses programs list: " 
							$array = $Setting.FilterAdapterAddressesPrograms.Values
							ForEach($element in $array)
							{
								WriteWordLine 0 3 $element
							}
							$array = $Null
						}
						If($Setting.VirtualLoopbackSupport.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Virtual IP\Virtual IP loopback support: " $Setting.VirtualLoopbackSupport.State
						}
						If($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Virtual IP\Virtual IP virtual loopback programs list: " 
							$array = $Setting.VirtualLoopbackPrograms.Values
							ForEach($element in $array)
							{
								WriteWordLine 0 3 $element
							}
							$array = $Null
						}
						Write-Verbose "$(Get-Date): `t`t`tXML Service"
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
						Write-Verbose "$(Get-Date): `t`tUser settings"
						Write-Verbose "$(Get-Date): `t`t`tICA"
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
						Write-Verbose "$(Get-Date): `t`t`tICA\Adobe Flash Delivery"
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
							$Values = $Null
						}
						If($Setting.FlashBackwardsCompatibility.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash backwards compatibility: " 
							WriteWordLine 0 3 $Setting.FlashBackwardsCompatibility.State
						}
						If($Setting.FlashDefaultBehavior.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash Default behavior: "
							Switch ($Setting.FlashDefaultBehavior.Value)
							{
								"Block"   {WriteWordLine 0 3 "Block Flash player"}
								"Disable" {WriteWordLine 0 3 "Disable Flash acceleration"}
								"Enable"  {WriteWordLine 0 3 "Enable Flash acceleration"}
								Default {WriteWordLine 0 3 "Flash Default behavior could not be determined: $($Setting.FlashDefaultBehavior.Value)"}
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
							$Values = $Null
						}
						If($Setting.FlashUrlCompatibilityList.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Flash Redirection\Flash URL compatibility list: " 
							$Values = $Setting.FlashUrlCompatibilityList.Values
							Write-Verbose "$(Get-Date): `t`t`t`tCreate table for Flash URL compatibility list"
							$TableRange = $doc.Application.Selection.Range
							[int]$Columns = 3
							[int]$Rows = $Values.count + 1
							Write-Verbose "$(Get-Date): `t`t`t`t`tAdd table to doc"
							$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
							$table.Style = $myHash.Word_TableGrid
							$table.Borders.InsideLineStyle = 1
							$table.Borders.OutsideLineStyle = 1
							[int]$xRow = 1
							Write-Verbose "$(Get-Date): `t`t`t`t`tFormat first row with column headings"
							$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
							$Table.Cell($xRow,1).Range.Font.Bold = $True
							$Table.Cell($xRow,1).Range.Text = "Action"
							$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
							$Table.Cell($xRow,2).Range.Font.Bold = $True
							$Table.Cell($xRow,2).Range.Text = "URL Pattern"
							$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
							$Table.Cell($xRow,3).Range.Font.Bold = $True
							$Table.Cell($xRow,3).Range.Text = "Flash Instance"
							ForEach($Value in $Values)
							{
								$Items = $Value.Split(' ')
								$xRow++
								Write-Verbose "$(Get-Date): `t`t`t`t`t`tProcessing row for $($Value)"
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
									$FlashInstance = ""
								}
								$Table.Cell($xRow,1).Range.Text = $Action
								$Table.Cell($xRow,2).Range.Text = $Url
								$Table.Cell($xRow,3).Range.Text = $FlashInstance
							}

							$Table.Rows.SetLeftIndent(108,1)
							$table.AutoFitBehavior(1)

							#return focus back to document
							Write-Verbose "$(Get-Date): `t`t`t`t`t`tReturn focus back to document"
							$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

							#move to the end of the current document
							Write-Verbose "$(Get-Date): `t`t`t`t`t`tMove to the end of the current document"
							$selection.EndKey($wdStory,$wdMove) | Out-Null
							$Values = $Null
							$Action = $Null
							$Url = $Null
							$FlashInstance = $Null
							$Spc = $Null
						}
						If($Setting.AllowSpeedFlash.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Adobe Flash Delivery\Legacy Server Side Optimizations\"
							WriteWordLine 0 3 "Flash quality adjustment: "
							Switch ($Setting.AllowSpeedFlash.Value)
							{
								"NoOptimization"      {WriteWordLine 0 3 "Do not optimize Flash animation options"}
								"AllConnections"      {WriteWordLine 0 3 "Optimize Flash animation options for all connections"}
								"RestrictedBandwidth" {WriteWordLine 0 3 "Optimize Flash animation options for low bandwidth connections only"}
								Default {WriteWordLine 0 3 "Flash quality adjustment could not be determined: $($Setting.AllowSpeedFlash.Value)"}
							}
						}
						Write-Verbose "$(Get-Date): `t`t`tICA\Audio"
						If($Setting.AudioPlugNPlay.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Audio\Audio Plug N Play: " $Setting.AudioPlugNPlay.State
						}
						If($Setting.AudioQuality.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Audio\Audio quality: "
							Switch ($Setting.AudioQuality.Value)
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
						Write-Verbose "$(Get-Date): `t`t`tICA\Bandwidth"
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
						Write-Verbose "$(Get-Date): `t`t`tICA\Desktop UI"
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
						Write-Verbose "$(Get-Date): `t`t`tICA\File Redirection"
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
						Write-Verbose "$(Get-Date): `t`t`tICA Multi-Stream Connections"
						If($Setting.MultiStream.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Multi-Stream Connections\Multi-Stream: " $Setting.MultiStream.State
						}
						Write-Verbose "$(Get-Date): `t`t`tICA\Port Redirection"
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
						Write-Verbose "$(Get-Date): `t`t`tICA\Printing"
						If($Setting.ClientPrinterRedirection.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Client printer redirection: " $Setting.ClientPrinterRedirection.State
						}
						If($Setting.DefaultClientPrinter.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Default printer - Choose client's Default printer: " 
							Switch ($Setting.DefaultClientPrinter.Value)
							{
								"ClientDefault" {WriteWordLine 0 3 "Set Default printer to the client's main printer"}
								"DoNotAdjust"   {WriteWordLine 0 3 "Do not adjust the user's Default printer"}
								Default {WriteWordLine 0 0 "Default printer could not be determined: $($Setting.DefaultClientPrinter.Value)"}
							}
						}
						If($Setting.AutoCreationEventLogPreference.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Printer auto-creation event log preference: " 
							Switch ($Setting.AutoCreationEventLogPreference.Value)
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
							ForEach($printer in $valArray)
							{
								$prArray = $printer.Split(',')
								ForEach($element in $prArray)
								{
									if($element.SubString(0, 2) -eq "\\")
									{
										$index = $element.SubString(2).IndexOf('\')
										if($index -ge 0)
										{
											$server = $element.SubString(0, $index + 2)
											$share  = $element.SubString($index + 3)
											WriteWordLine 0 3 "Server`t`t: $server"
											WriteWordLine 0 3 "Shared Name`t: $share"
										}
										$index = $Null
									}
									Else
									{
										$tmp = $element.SubString(0, 4)
										$PrtString = Get-PrinterModifiedSettings $tmp $element $False
										If(![String]::IsNullOrEmpty($PrtString))
										{
											WriteWordLine 0 3 $PrtString
										}
										$tmp = $Null
										$PrtString = $Null
									}
								}
								WriteWordLine 0 0 ""
							}
							$valArray = $Null
							$prArray = $Null
						}
						If($Setting.WaitForPrintersToBeCreated.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Wait for printers to be created (desktop): " $Setting.WaitForPrintersToBeCreated.State
						}
						If($Setting.ClientPrinterAutoCreation.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Client Printers\Auto-create client printers: "
							Switch ($Setting.ClientPrinterAutoCreation.Value)
							{
								"DoNotAutoCreate"    {WriteWordLine 0 3 "Do not auto-create client printers"}
								"DefaultPrinterOnly" {WriteWordLine 0 3 "Auto-create the client's Default printer only"}
								"LocalPrintersOnly"  {WriteWordLine 0 3 "Auto-create local (non-network) client printers only"}
								"AllPrinters"        {WriteWordLine 0 3 "Auto-create all client printers"}
								Default {WriteWordLine 0 3 "Auto-create client printers could not be determined: $($Setting.ClientPrinterAutoCreation.Value)"}
							}
						}
						If($Setting.GenericUniversalPrinterAutoCreation.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Client Printers\Auto-create generic universal printer: " $Setting.GenericUniversalPrinterAutoCreation.State
						}
						If($Setting.ClientPrinterNames.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Client Printers\Client printer names: " 
							Switch ($Setting.ClientPrinterNames.Value)
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
							Write-Verbose "$(Get-Date): `t`t`t`tCreate table for printer drive mapping"
							$TableRange = $doc.Application.Selection.Range
							[int]$Columns = 4
							[int]$Rows = $array.count + 1
							Write-Verbose "$(Get-Date): `t`t`t`t`tAdd table to doc"
							$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
							$table.Style = $myHash.Word_TableGrid
							$table.Borders.InsideLineStyle = 1
							$table.Borders.OutsideLineStyle = 1
							[int]$xRow = 1
							Write-Verbose "$(Get-Date): `t`t`t`t`tFormat first row with column headings"
							$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
							$Table.Cell($xRow,1).Range.Font.Bold = $True
							$Table.Cell($xRow,1).Range.Text = "Driver Name"
							$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
							$Table.Cell($xRow,2).Range.Font.Bold = $True
							$Table.Cell($xRow,2).Range.Text = "Action"
							$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
							$Table.Cell($xRow,3).Range.Font.Bold = $True
							$Table.Cell($xRow,3).Range.Text = "Settings"
							$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
							$Table.Cell($xRow,4).Range.Font.Bold = $True
							$Table.Cell($xRow,4).Range.Text = "Server Driver"
							ForEach($element in $array)
							{
								#WriteWordLine 0 3 $element
								$Items = $element.Split(',')
								$xRow++
								Write-Verbose "$(Get-Date): `t`t`t`t`t`tProcessing row for $($Items[0])"
								$DriverName = $Items[0]
								$Action = $Items[1]
								If($Action -match 'Replace=')
								{
									$ServerDriver = $Action.substring($Action.indexof("=")+1)
									$Action = "Replace"
								}
								Else
								{
									$ServerDriver = ""
									If($Action -eq "Allow")
									{
										$Action = "Allow"
									}
									ElseIf($Action -eq "Deny")
									{
										$Action = "Do not create"
									}
									ElseIf($Action -eq "UPD_Only")
									{
										$Action = "Create with universal driver"
									}
								}
								If($Items.count -gt 2)
								{
									$PrtSettings = ""
									[int]$BeginAt = 2
									[int]$EndAt = $Items.count
									for ($i=$BeginAt;$i -lt $EndAt; $i++) 
									{
										$tmp = $Items[$i].SubString(0, 4)
										$tmp1 = Get-PrinterModifiedSettings $tmp $Items[$i] $True
										If(![String]::IsNullOrEmpty($tmp1))
										{
											$PrtSettings += ($tmp1 + "`n")
										}
									}
								}
								Else
								{
									$PrtSettings = "Unmodified"
								}
								$Table.Cell($xRow,1).Range.Text = $DriverName
								$Table.Cell($xRow,2).Range.Text = $Action
								$Table.Cell($xRow,3).Range.Text = $PrtSettings
								$Table.Cell($xRow,4).Range.Text = $ServerDriver
							}
							$array = $Null
							$Table.Rows.SetLeftIndent(108,1)
							$table.AutoFitBehavior(1)

							#return focus back to document
							Write-Verbose "$(Get-Date): `t`t`t`t`tReturn focus back to document"
							$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

							#move to the end of the current document
							Write-Verbose "$(Get-Date): `t`t`t`t`tMove to the end of the current document"
							$selection.EndKey($wdStory,$wdMove) | Out-Null
							WriteWordLine 0 0 ""
						}
						If($Setting.PrinterPropertiesRetention.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Client Printers\Printer properties retention: " 
							Switch ($Setting.PrinterPropertiesRetention.Value)
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
							$TmpArray = $Setting.UniversalDriverPriority.Value.Split(';')
							ForEach($Thing in $TmpArray)
							{
								WriteWordLine 0 3 $Thing
							}
							$TmpArray = $Null
						}
						If($Setting.UniversalPrintDriverUsage.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Drivers\Universal print driver usage: " 
							Switch ($Setting.UniversalPrintDriverUsage.Value)
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
							Switch ($Setting.EMFProcessingMode.Value)
							{
								"ReprocessEMFsForPrinter" {WriteWordLine 0 3 "Reprocess EMFs for printer"}
								"SpoolDirectlyToPrinter"  {WriteWordLine 0 3 "Spool directly to printer"}
								Default {WriteWordLine 0 3 "Universal printing EMF processing mode could not be determined: $($Setting.EMFProcessingMode.Value)"}
							}
						}
						If($Setting.ImageCompressionLimit.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing image compression limit: " 
							Switch ($Setting.ImageCompressionLimit.Value)
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
							WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing optimization defaults: "
							$TmpArray = $Setting.UPDCompressionDefaults.Value.Split(';')
							ForEach($Thing in $TmpArray)
							{
								$TestLabel = $Thing.substring(0, $Thing.indexof("="))
								$TestSetting = $Thing.substring($Thing.indexof("=")+1)
								$TxtLabel = ""
								$TxtSetting = "ABC"
								Switch($TestLabel)
								{
									"ImageCompression"
									{
										$TxtLabel = "Desired image quality:"
										Switch($TestSetting)
										{
											"StandardQuality"	{$TxtSetting = "Standard quality"}
											"BestQuality"	{$TxtSetting = "Best quality (lossless compression)"}
											"HighQuality"	{$TxtSetting = "High quality"}
											"ReducedQuality"	{$TxtSetting = "Reduced quality (maximum compression)"}
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
								WriteWordLine 0 3 "$TxtLabel $TxtSetting"
							}
							$TmpArray = $Null
						}
						If($Setting.UniversalPrintingPreviewPreference.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Printing\Universal Printing\Universal printing preview preference: " 
							Switch ($Setting.UniversalPrintingPreviewPreference.Value)
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
							Switch ($Setting.DPILimit.Value)
							{
								"Draft"            {WriteWordLine 0 3 "Draft (150 DPI)"}
								"LowResolution"    {WriteWordLine 0 3 "Low Resolution (300 DPI)"}
								"MediumResolution" {WriteWordLine 0 3 "Medium Resolution (600 DPI)"}
								"HighResolution"   {WriteWordLine 0 3 "High Resolution (1200 DPI)"}
								"Unlimited "       {WriteWordLine 0 3 "No Limit"}
								Default {WriteWordLine 0 3 "Universal printing print quality limit could not be determined: $($Setting.DPILimit.Value)"}
							}
						}
						Write-Verbose "$(Get-Date): `t`t`tICA\Security"
						If($Setting.MinimumEncryptionLevel.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Security\SecureICA minimum encryption level: " 
							Switch ($Setting.MinimumEncryptionLevel.Value)
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
						Write-Verbose "$(Get-Date): `t`t`tICA\Session limits"
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
						Write-Verbose "$(Get-Date): `t`t`tICA\Shadowing"
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
							ForEach($element in $array)
							{
								$x = $element.indexof("/",8)
								$tmp = $element.substring(8,$x-8)
								WriteWordLine 0 3 $tmp
							}
							$array = $Null
							$x = $Null
							$tmp = $Null
						}
						If($Setting.ShadowDenyList.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Shadowing\Users who cannot shadow other users: " 
							$array = $Setting.ShadowDenyList.Values
							ForEach($element in $array)
							{
								$x = $element.indexof("/",8)
								$tmp = $element.substring(8,$x-8)
								WriteWordLine 0 3 $tmp
							}
							$array = $Null
							$x = $Null
							$tmp = $Null
						}
						Write-Verbose "$(Get-Date): `t`t`tICA\Time Zone Control"
						If($Setting.LocalTimeEstimation.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Time Zone Control\Estimate local time for legacy clients: " $Setting.LocalTimeEstimation.State
						}
						If($Setting.SessionTimeZone.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Time Zone Control\Use local time of client: " 
							Switch ($Setting.SessionTimeZone.Value)
							{
								"UseServerTimeZone" {WriteWordLine 0 3 "Use server time zone"}
								"UseClientTimeZone" {WriteWordLine 0 3 "Use client time zone"}
								Default {WriteWordLine 0 3 "Use local time of client could not be determined: $($Setting.SessionTimeZone.Value)"}
							}
						}
						Write-Verbose "$(Get-Date): `t`t`tICA\TWAIN devices"
						If($Setting.TwainRedirection.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\TWAIN devices\Client TWAIN device redirection: " $Setting.TwainRedirection.State
						}
						If($Setting.TwainCompressionLevel.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\TWAIN devices\TWAIN compression level: " 
							Switch ($Setting.TwainCompressionLevel.Value)
							{
								"None"   {WriteWordLine 0 3 "None"}
								"Low"    {WriteWordLine 0 3 "Low"}
								"Medium" {WriteWordLine 0 3 "Medium"}
								"High"   {WriteWordLine 0 3 "High"}
								Default {WriteWordLine 0 3 "TWAIN compression level could not be determined: $($Setting.TwainCompressionLevel.Value)"}
							}
						}
						Write-Verbose "$(Get-Date): `t`t`tICA\USB devices"
						If($Setting.UsbDeviceRedirection.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\USB devices\Client USB device redirection: " $Setting.UsbDeviceRedirection.State
						}
						If($Setting.UsbDeviceRedirectionRules.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\USB devices\Client USB device redirection rules: " 
							$array = $Setting.UsbDeviceRedirectionRules.Values
							ForEach($element in $array)
							{
								WriteWordLine 0 3 $element
							}
							$array = $Null
						}
						If($Setting.UsbPlugAndPlayRedirection.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\USB devices\Client USB Plug and Play device redirection: " $Setting.UsbPlugAndPlayRedirection.State
						}
						Write-Verbose "$(Get-Date): `t`t`tICA\Visual Display"
						If($Setting.FramesPerSecond.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Visual Display\Max Frames Per Second (fps): " $Setting.FramesPerSecond.Value
						}
						If($Setting.ProgressiveCompressionLevel.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "ICA\Visual Display\Moving Images\Progressive compression level: " -nonewline
							Switch ($Setting.ProgressiveCompressionLevel.Value)
							{
								"UltraHigh" {WriteWordLine 0 0 "Ultra high"}
								"VeryHigh"  {WriteWordLine 0 0 "Very high"}
								"High"      {WriteWordLine 0 0 "High"}
								"Normal"    {WriteWordLine 0 0 "Normal"}
								"Low"       {WriteWordLine 0 0 "Low"}
								"None"      {WriteWordLine 0 0 "None"}
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
							Switch ($Setting.LossyCompressionLevel.Value)
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
						Write-Verbose "$(Get-Date): `t`t`tServer Session Settings"
						If($Setting.SessionImportance.State -ne "NotConfigured")
						{
							WriteWordLine 0 2 "Server Session Settings\Session importance: " 
							Switch ($Setting.SessionImportance.Value)
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
			$Filter = $Null
			$Settings = $Null
			Write-Verbose "$(Get-Date): `t`tFinished $($Policy.PolicyName)`t$($Policy.Type)"
			Write-Verbose "$(Get-Date): "
		}
	}
	Else 
	{
		Write-Warning "Citrix Policy information could not be retrieved."
	}
	$Policies = $Null
	If($xDriveName -ne "")
	{
		Write-Verbose "$(Get-Date): `tRemoving ADGpoDrv PSDrive"
		Remove-PSDrive ADGpoDrv -EA 0
		Write-Verbose "$(Get-Date): "
	}
}

Function GetCtxGPOsInAD
{
	#thanks to the Citrix Engineering Team for pointers and for Michael B. Smith for creating the function
	#updated 07-Nov-13 to work in a Windows Workgroup environment
	$root = [ADSI]"LDAP://RootDSE"
	If([String]::IsNullOrEmpty($root.PSBase.Name))
	{
		Write-Verbose "$(Get-Date): Not in an Active Directory environment"
		$root = $null
		$xArray = @()
	}
	Else
	{
		$domainNC = $root.defaultNamingContext.ToString()
		$root = $null
		$xArray = @()

		$domain = $domainNC.Replace( 'DC=', '' ).Replace( ',', '.' )
		$sysvolFiles = dir -Recurse ( '\\' + $domain  + '\sysvol\' + $domain + '\Policies' )
		foreach( $file in $sysvolFiles )
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
						$xArray += $gpObject.DisplayName	### name of the group policy object
					}
				}
			}
		}
	}
	Return ,$xArray | Sort
}

#Script begins

$script:startTime = Get-Date

If(!(Check-NeededPSSnapins "Citrix.Common.Commands","Citrix.XenApp.Commands"))
{
    #We're missing Citrix Snapins that we need
    Write-Error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Are you sure you are running this script on a XenApp 6.5 Server? Script will now close."
    Exit
}

CheckWordPreReq

#if software inventory is specified then verify SoftwareExclusions.txt exists
If($Software)
{
	If(!(Test-Path "$($pwd.path)\SoftwareExclusions.txt"))
	{
		Write-Error "Software inventory requested but $($pwd.path)\SoftwareExclusions.txt does not exist.  Script cannot continue."
		Exit
	}
	
	#file does exist but can we access it?
	$x = Get-Content "$($pwd.path)\SoftwareExclusions.txt" -EA 0
	If(!($?))
	{
		Write-Error "There was an error accessing or reading $($pwd.path)\SoftwareExclusions.txt.  Script cannot continue."
		Exit
	}
	$x = $Null
}

[bool]$Remoting = $False
$RemoteXAServer = Get-XADefaultComputerName -EA 0
If(![String]::IsNullOrEmpty($RemoteXAServer))
{
	$Remoting = $True
}

If($Remoting)
{
	Write-Verbose "$(Get-Date): Remoting is enabled to XenApp server $RemoteXAServer"
}
Else
{
	Write-Verbose "$(Get-Date): Remoting is not being used"
	
	#now need to make sure the script is not being run on a session-only host
	$ServerName = (Get-Childitem env:computername).value
	$Server = Get-XAServer -ServerName $ServerName -EA 0
	If($Server.ElectionPreference -eq "WorkerMode")
	{
		Write-Warning "This script cannot be run on a Session-only Host Server if Remoting is not enabled."
		Write-Warning "Use Set-XADefaultComputerName XA65ControllerServerName or run the script on a controller."
		Write-Error "Script cannot continue.  See messages above."
		Exit
	}
}

# Get farm information
Write-Verbose "$(Get-Date): Getting Farm data"
$farm = Get-XAFarm -EA 0

If($?)
{
	Write-Verbose "$(Get-Date): Verify farm version"
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
	[string]$FarmName = $farm.FarmName
	[string]$Title = "Inventory Report for the $($FarmName) Farm"
	[string]$filename1 = "$($pwd.path)\$($farm.FarmName).docx"
	If($PDF)
	{
		[string]$filename2 = "$($pwd.path)\$($farm.FarmName).pdf"
	}
} 
Else 
{
	Write-Warning "Farm information could not be retrieved"
	If($Remoting)
	{
		Write-Error "A remote connection to $RemoteXAServer could not be established.  Script cannot continue."
	}
	Else
	{
		Write-Error "Farm information could not be retrieved.  Script cannot continue."
	}
	Exit
}
$farm = $Null

Write-Verbose "$(Get-Date): Setting up Word"

# Setup word for output
Write-Verbose "$(Get-Date): Create Word comObject.  If you are not running Word 2007, ignore the next message."
$Word = New-Object -comobject "Word.Application" -EA 0

If(!$? -or $Word -eq $Null)
{
	Write-Warning "The Word object could not be created.  You may need to repair your Word installation."
	Write-Error "The Word object could not be created.  You may need to repair your Word installation.  Script cannot continue."
	Exit
}

[int]$WordVersion = [int]$Word.Version
If($WordVersion -eq $wdWord2013)
{
	$WordProduct = "Word 2013"
}
ElseIf($WordVersion -eq $wdWord2010)
{
	$WordProduct = "Word 2010"
}
ElseIf($WordVersion -eq $wdWord2007)
{
	$WordProduct = "Word 2007"
}
Else
{
	Write-Error "You are running an untested or unsupported version of Microsoft Word.  Script will end.  Please send info on your version of Word to webster@carlwebster.com"
	AbortScript
}

Write-Verbose "$(Get-Date): Running Microsoft $WordProduct"

If($PDF -and $WordVersion -eq $wdWord2007)
{
	Write-Verbose "$(Get-Date): Verify the Word 2007 Save As PDF add-in is installed"
	If(CheckWord2007SaveAsPDFInstalled)
	{
		Write-Verbose "$(Get-Date): The Word 2007 Save As PDF add-in is installed"
	}
	Else
	{
		AbortScript
	}
}

Write-Verbose "$(Get-Date): Validate company name"
#only validate CompanyName if the field is blank
If([String]::IsNullOrEmpty($CompanyName))
{
	$CompanyName = ValidateCompanyName
	If([String]::IsNullOrEmpty($CompanyName))
	{
		Write-Warning "Company Name cannot be blank."
		Write-Warning "Check HKCU:\Software\Microsoft\Office\Common\UserInfo for Company or CompanyName value."
		Write-Error "Script cannot continue.  See messages above."
		AbortScript
	}
}

Write-Verbose "$(Get-Date): Check Default Cover Page for language specific version"
[bool]$CPChanged = $False
Switch ($PSUICulture.Substring(0,3))
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
				If($WordVersion -eq $wdWord2013)
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
}

If($CPChanged)
{
	Write-Verbose "$(Get-Date): Changed Default Cover Page from Sideline to $($CoverPage)"
}

Write-Verbose "$(Get-Date): Validate cover page"
[bool]$ValidCP = ValidateCoverPage $WordVersion $CoverPage
If(!$ValidCP)
{
	Write-Error "For $WordProduct, $CoverPage is not a valid Cover Page option.  Script cannot continue."
	AbortScript
}

Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Company Name : $CompanyName"
Write-Verbose "$(Get-Date): Cover Page   : $CoverPage"
Write-Verbose "$(Get-Date): User Name    : $UserName"
Write-Verbose "$(Get-Date): Save As PDF  : $PDF"
Write-Verbose "$(Get-Date): HW Inventory : $Hardware"
Write-Verbose "$(Get-Date): SW Inventory : $Software"
Write-Verbose "$(Get-Date): Farm Name    : $FarmName"
Write-Verbose "$(Get-Date): Title        : $Title"
Write-Verbose "$(Get-Date): Filename1    : $filename1"
If($PDF)
{
	Write-Verbose "$(Get-Date): Filename2    : $filename2"
}
Write-Verbose "$(Get-Date): OS Detected  : $RunningOS"
Write-Verbose "$(Get-Date): PSUICulture  : $PSUICulture"
Write-Verbose "$(Get-Date): PSCulture    : $PSCulture"
Write-Verbose "$(Get-Date): Word version : $WordProduct"
Write-Verbose "$(Get-Date): Word language: $($Word.Language)"
Write-Verbose "$(Get-Date): PoSH version : $($Host.Version)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): Script start : $($Script:StartTime)"
Write-Verbose "$(Get-Date): "
Write-Verbose "$(Get-Date): "

$Word.Visible = $False

#http://jdhitsolutions.com/blog/2012/05/san-diego-2012-powershell-deep-dive-slides-and-demos/
#using Jeff's Demo-WordReport.ps1 file for examples
#down to $configlog = $False is from Jeff Hicks
Write-Verbose "$(Get-Date): Load Word Templates"

[bool]$CoverPagesExist = $False
[bool]$BuildingBlocksExist = $False

$word.Templates.LoadBuildingBlocks()
If($WordVersion -eq $wdWord2007)
{
	$BuildingBlocks = $word.Templates | Where {$_.name -eq "Building Blocks.dotx"}
}
Else
{
	$BuildingBlocks = $word.Templates | Where {$_.name -eq "Built-In Building Blocks.dotx"}
}

Write-Verbose "$(Get-Date): Attempt to load cover page $($CoverPage)"
$part = $Null

If($BuildingBlocks -ne $Null)
{
	$BuildingBlocksExist = $True

	Try 
		{$part = $BuildingBlocks.BuildingBlockEntries.Item($CoverPage)}

	Catch
		{$part = $Null}

	If($part -ne $Null)
	{
		$CoverPagesExist = $True
	}
}

#cannot continue if cover page does not exist
If(!$CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Cover Pages are not installed or the Cover Page $($CoverPage) does not exist."
	Write-Error "Cover Pages are not installed or the Cover Page $($CoverPage) does not exist.  Script cannot continue."
	Write-Verbose "$(Get-Date): Closing Word"
	AbortScript
}

Write-Verbose "$(Get-Date): Create empty word doc"
$Doc = $Word.Documents.Add()
If($Doc -eq $Null)
{
	Write-Verbose "$(Get-Date): "
	Write-Error "An empty Word document could not be created.  Script cannot continue."
	AbortScript
}

$Selection = $Word.Selection
If($Selection -eq $Null)
{
	Write-Verbose "$(Get-Date): "
	Write-Error "An unknown error happened selecting the entire Word document for default formatting options.  Script cannot continue."
	AbortScript
}

#set Default tab stops to 1/2 inch (this line is not from Jeff Hicks)
#36 = .50"
$Word.ActiveDocument.DefaultTabStop = 36

#Disable Spell and Grammar Check to resolve issue and improve performance (from Pat Coughlin)
Write-Verbose "$(Get-Date): Disable grammar and spell checking"
$Word.Options.CheckGrammarAsYouType = $False
$Word.Options.CheckSpellingAsYouType = $False

If($BuildingBlocksExist)
{
	#insert new page, getting ready for table of contents
	Write-Verbose "$(Get-Date): Insert new page, getting ready for table of contents"
	$part.Insert($selection.Range,$True) | out-null
	$selection.InsertNewPage()

	#table of contents
	Write-Verbose "$(Get-Date): Table of Contents - $($myHash.Word_TableOfContents)"
	$toc = $BuildingBlocks.BuildingBlockEntries.Item($myHash.Word_TableOfContents)
	If($toc -eq $Null)
	{
		Write-Verbose "$(Get-Date): "
		Write-Verbose "$(Get-Date): Table of Content - $($myHash.Word_TableOfContents) could not be retrieved."
		Write-Warning "This report will not have a Table of Contents."
	}
	Else
	{
		$toc.insert($selection.Range,$True) | out-null
	}
}
Else
{
	Write-Verbose "$(Get-Date): Table of Contents are not installed."
	Write-Warning "Table of Contents are not installed so this report will not have a Table of Contents."
}

#set the footer
Write-Verbose "$(Get-Date): Set the footer"
[string]$footertext = "Report created by $username"

#get the footer
Write-Verbose "$(Get-Date): Get the footer and format font"
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekPrimaryFooter
#get the footer and format font
$footers = $doc.Sections.Last.Footers
ForEach($footer in $footers) 
{
	If($footer.exists) 
	{
		$footer.range.Font.name = "Calibri"
		$footer.range.Font.size = 8
		$footer.range.Font.Italic = $True
		$footer.range.Font.Bold = $True
	}
} #end ForEach
Write-Verbose "$(Get-Date): Footer text"
$selection.HeaderFooter.Range.Text = $footerText

#add page numbering
Write-Verbose "$(Get-Date): Add page numbering"
$selection.HeaderFooter.PageNumbers.Add($wdAlignPageNumberRight) | Out-Null

#return focus to main document
Write-Verbose "$(Get-Date): Return focus to main document"
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
Write-Verbose "$(Get-Date): Move to the end of the current document"
Write-Verbose "$(Get-Date):"
$selection.EndKey($wdStory,$wdMove) | Out-Null
#end of Jeff Hicks 

Write-Verbose "$(Get-Date): Processing Configuration Logging"
[bool]$ConfigLog = $False
$ConfigurationLogging = Get-XAConfigurationLog -EA 0

If($?)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Configuration Logging"
	If($ConfigurationLogging.LoggingEnabled) 
	{
		$ConfigLog = $True
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
		$Tmp = $Null
	}
	Else 
	{
		WriteWordLine 0 1 "Configuration Logging is disabled."
	}
}
Else 
{
	Write-Warning  "Configuration Logging could not be retrieved"
}
$ConfigurationLogging = $Null
Write-Verbose "$(Get-Date): Finished Configuration Logging"
Write-Verbose "$(Get-Date): "

Write-Verbose "$(Get-Date): Processing Administrators"
Write-Verbose "$(Get-Date): `tSetting summary variables"
[int]$TotalFullAdmins = 0
[int]$TotalViewAdmins = 0
[int]$TotalCustomAdmins = 0

Write-Verbose "$(Get-Date): `tRetrieving Administrators"
$Administrators = Get-XAAdministrator -EA 0 | Sort-Object AdministratorName

If($?)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Administrators:"
	ForEach($Administrator in $Administrators)
	{
		Write-Verbose "$(Get-Date): `t`tProcessing administrator $($Administrator.AdministratorName)"
		WriteWordLine 2 0 $Administrator.AdministratorName
		WriteWordLine 0 1 "Administrator type: " -nonewline
		Switch ($Administrator.AdministratorType)
		{
			"Unknown"  {WriteWordLine 0 0 "Unknown"}
			"Full"     {WriteWordLine 0 0 "Full Administration"; $TotalFullAdmins++}
			"ViewOnly" {WriteWordLine 0 0 "View Only"; $TotalViewAdmins++}
			"Custom"   {WriteWordLine 0 0 "Custom"; $TotalCustomAdmins++}
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
		If($Administrator.AdministratorType -eq "Custom") 
		{
			WriteWordLine 0 1 "Farm Privileges:"
			ForEach($farmprivilege in $Administrator.FarmPrivileges) 
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing farm privilege $farmprivilege"
				Switch ($farmprivilege)
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
	
			Write-Verbose "$(Get-Date): `t`t`tProcessing folder privileges"
			WriteWordLine 0 1 "Folder Privileges:"
			ForEach($folderprivilege in $Administrator.FolderPrivileges) 
			{
				#The Citrix PoSH cmdlet only returns data for three folders:
				#Servers
				#WorkerGroups
				#Applications
				
				Write-Verbose "$(Get-Date): `t`t`t`tProcessing folder permissions for $($FolderPrivilege.FolderPath)"
				WriteWordLine 0 2 $FolderPrivilege.FolderPath
				ForEach($FolderPermission in $FolderPrivilege.FolderPrivileges)
				{
					Switch ($folderpermission)
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
	Write-Warning "Administrator information could not be retrieved"
}
$Administrators = $Null
Write-Verbose "$(Get-Date): Finished Processing Administrators"
Write-Verbose "$(Get-Date): "

Write-Verbose "$(Get-Date): Processing Applications"

Write-Verbose "$(Get-Date): `tSetting summary variables"
[int]$TotalPublishedApps = 0
[int]$TotalPublishedContent = 0
[int]$TotalPublishedDesktops = 0
[int]$TotalStreamedApps = 0
$SessionSharingItems = @()

Write-Verbose "$(Get-Date): `tRetrieving Applications"
$Applications = Get-XAApplication -EA 0 | Sort-Object FolderPath, DisplayName

If($? -and $Applications -ne $Null)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Applications:"

	ForEach($Application in $Applications)
	{
		Write-Verbose "$(Get-Date): `t`tProcessing application $($Application.BrowserName)"
		
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
			
			Write-Verbose "$(Get-Date): `t`t`tGather session sharing info for Appendix A"
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
			$SessionSharingItems += $obj
		}
		$AppServerInfoResults = $False
		$AppServerInfo = Get-XAApplicationReport -BrowserName $Application.BrowserName -EA 0
		If($?)
		{
			$AppServerInfoResults = $True
		}
		[bool]$streamedapp = $False
		If($Application.ApplicationType -Contains "streamedtoclient" -or $Application.ApplicationType -Contains "streamedtoserver")
		{
			$streamedapp = $True
		}
		#name properties
		WriteWordLine 2 0 $Application.DisplayName
		WriteWordLine 0 1 "Application name`t`t: " $Application.BrowserName
		WriteWordLine 0 1 "Disable application`t`t: " -NoNewLine
		#weird, if application is enabled, it is disabled!
		If($Application.Enabled) 
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

		If(![String]::IsNullOrEmpty($Application.Description))
		{
			WriteWordLine 0 1 "Application description`t`t: " $Application.Description
		}
		
		#type properties
		WriteWordLine 0 1 "Application Type`t`t: " -nonewline
		Switch ($Application.ApplicationType)
		{
			"Unknown"                            {WriteWordLine 0 0 "Unknown"}
			"ServerInstalled"                    {WriteWordLine 0 0 "Installed application"; $TotalPublishedApps++}
			"ServerDesktop"                      {WriteWordLine 0 0 "Server desktop"; $TotalPublishedDesktops++}
			"Content"                            {WriteWordLine 0 0 "Content"; $TotalPublishedContent++}
			"StreamedToServer"                   {WriteWordLine 0 0 "Streamed to server"; $TotalStreamedApps++}
			"StreamedToClient"                   {WriteWordLine 0 0 "Streamed to client"; $TotalStreamedApps++}
			"StreamedToClientOrInstalled"        {WriteWordLine 0 0 "Streamed if possible, otherwise accessed from server as Installed application"; $TotalStreamedApps++}
			"StreamedToClientOrStreamedToServer" {WriteWordLine 0 0 "Streamed if possible, otherwise Streamed to server"; $TotalStreamedApps++}
			Default {WriteWordLine 0 0 "Application Type could not be determined: $($Application.ApplicationType)"}
		}
		If(![String]::IsNullOrEmpty($Application.FolderPath))
		{
			WriteWordLine 0 1 "Folder path`t`t`t: " $Application.FolderPath
		}
		If(![String]::IsNullOrEmpty($Application.ContentAddress))
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
			If(![String]::IsNullOrEmpty($Application.ProfileProgramArguments))
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
				Switch ($Application.CachingOption)
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
			If(![String]::IsNullOrEmpty($Application.CommandLineExecutable))
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
			If(![String]::IsNullOrEmpty($Application.WorkingDirectory))
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
				If(![String]::IsNullOrEmpty($AppServerInfo.ServerNames))
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
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			[int]$Rows = $Application.AccessSessionConditions.count + 1
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 1
			$table.Borders.OutsideLineStyle = 1
			[int]$xRow = 1
			Write-Verbose "$(Get-Date): `t`t`t`tFormat first row with column headings"
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Farm Name"
			$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,2).Range.Font.Bold = $True
			$Table.Cell($xRow,2).Range.Text = "Filter"
			ForEach($AccessCondition in $Application.AccessSessionConditions)
			{
				[string]$Tmp = $AccessCondition
				[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
				[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
				$xRow++
				Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing row for Access Condition $($Tmp)"
				$Table.Cell($xRow,1).Range.Text = $AGFarm
				$Table.Cell($xRow,2).Range.Text = $AGFilter
			}

			Write-Verbose "$(Get-Date): `t`t`t`tMove table to the right"
			$Table.Rows.SetLeftIndent(72,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			Write-Verbose "$(Get-Date): `t`t`t`tReturn focus back to document"
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			Write-Verbose "$(Get-Date): `t`t`t`tMove to the end of the current document"
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			$tmp = $Null
			$AGFarm = $Null
			$AGFilter = $Null
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
	
		If($Application.MultipleInstancesPerUserAllowed) 
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
			Switch ($Application.CpuPriorityLevel)
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
		Switch ($Application.AudioType)
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
			Switch ($Application.EncryptionLevel)
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
		If($Application.WaitOnPrinterCreation) 
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
			Switch ($Application.ColorDepth)
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
	$AppServerInfo = $Null
	}
}
ElseIf($Applications -eq $Null)
{
	Write-Verbose "$(Get-Date): There are no Applications published"
}
Else 
{
	Write-Warning "Application information could not be retrieved."
}
$Applications = $Null
Write-Verbose "$(Get-Date): Finished Processing Applications"
Write-Verbose "$(Get-Date): "

[int]$TotalConfigLogItems = 0

Write-Verbose "$(Get-Date): Processing Configuration Logging/History Report"
If($ConfigLog)
{
	#history AKA Configuration Logging report
	#only process if $ConfigLog = $True and XA65ConfigLog.udl file exists
	#build connection string
	#User ID is account that has access permission for the configuration logging database
	#Initial Catalog is the name of the Configuration Logging SQL Database
	#bug fixed by Esther Barthel
	If(Test-Path "$($pwd.path)\XA65ConfigLog.udl")
	{
		$ConnectionString = Get-Content "$($pwd.path)\XA65ConfigLog.udl" | select-object -last 1
		$ConfigLogReport = get-CtxConfigurationLogReport -connectionstring $ConnectionString -EA 0

		If($? -and $ConfigLogReport)
		{
			Write-Verbose "$(Get-Date): `tProcessing $($ConfigLogReport.Count) history items"
			$selection.InsertNewPage()
			WriteWordLine 1 0 "History:"
			ForEach($ConfigLogItem in $ConfigLogReport)
			{
				$TotalConfigLogItems++
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
		$ConnectionString = $Null
		$ConfigLogReport = $Null
	}
	Else 
	{
		WriteWordLine 1 0 "Configuration Logging is enabled but the XA65ConfigLog.udl file was not found"
	}
}

Write-Verbose "$(Get-Date): Finished Processing Configuration Logging/History Report"
Write-Verbose "$(Get-Date): "

#load balancing policies
Write-Verbose "$(Get-Date): Processing Load Balancing Policies"
Write-Verbose "$(Get-Date): `tSetting summary variables"
[int]$TotalLBPolicies = 0

Write-Verbose "$(Get-Date): `tRetrieving Load Balancing Policies"
$LoadBalancingPolicies = Get-XALoadBalancingPolicy -EA 0 | Sort-Object PolicyName

If($? -and $LoadBalancingPolicies -ne $Null)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Load Balancing Policies:"
	ForEach($LoadBalancingPolicy in $LoadBalancingPolicies)
	{
		$TotalLBPolicies++
		Write-Verbose "$(Get-Date): `t`tProcessing Load Balancing Policy $($LoadBalancingPolicy.PolicyName)"
		$LoadBalancingPolicyConfiguration = Get-XALoadBalancingPolicyConfiguration -PolicyName $LoadBalancingPolicy.PolicyName -EA 0
		$LoadBalancingPolicyFilter = Get-XALoadBalancingPolicyFilter -PolicyName $LoadBalancingPolicy.PolicyName -EA 0
	
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
				If($LoadBalancingPolicyFilter.AllowOtherConnections)
				{
					WriteWordLine 0 2 "Any connection"
				} 
				Else
				{
					WriteWordLine 0 2 "Any connection that meets any of the following filters"
					If($LoadBalancingPolicyFilter.AccessSessionConditions)
					{
						Write-Verbose "$(Get-Date): `t`t`tCreate table for Load Balancing Policy Access Session Condition"
						$TableRange = $doc.Application.Selection.Range
						[int]$Columns = 2
						[int]$Rows = $LoadBalancingPolicyFilter.AccessSessionConditions.count + 1
						$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
						$table.Style = $myHash.Word_TableGrid
						$table.Borders.InsideLineStyle = 1
						$table.Borders.OutsideLineStyle = 1
						[int]$xRow = 1
						Write-Verbose "$(Get-Date): `t`t`t`tFormat first row with column headings"
						$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,1).Range.Font.Bold = $True
						$Table.Cell($xRow,1).Range.Text = "Farm Name"
						$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
						$Table.Cell($xRow,2).Range.Font.Bold = $True
						$Table.Cell($xRow,2).Range.Text = "Filter"
						ForEach($AccessSessionCondition in $LoadBalancingPolicyFilter.AccessSessionConditions)
						{
							[string]$Tmp = $AccessSessionCondition
							[string]$AGFarm = $Tmp.substring(0, $Tmp.indexof(":"))
							[string]$AGFilter = $Tmp.substring($Tmp.indexof(":")+1)
							$xRow++
							Write-Verbose "$(Get-Date): `t`t`t`t`tProcessing row for Access Session Condition $($Tmp)"
							$Table.Cell($xRow,1).Range.Text = $AGFarm
							$Table.Cell($xRow,2).Range.Text = $AGFilter
						}

						Write-Verbose "$(Get-Date): `t`t`t`tMove table to the right"
						$Table.Rows.SetLeftIndent(72,1)
						$table.AutoFitBehavior(1)

						#return focus back to document
						Write-Verbose "$(Get-Date): `t`t`t`tReturn focus back to document"
						$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

						#move to the end of the current document
						Write-Verbose "$(Get-Date): `t`t`t`tMove to the end of the current document"
						$selection.EndKey($wdStory,$wdMove) | Out-Null
						$tmp = $Null
						$AGFarm = $Null
						$AGFilter = $Null
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
				Write-Verbose "$(Get-Date): `t`t`tCreate table for Load Balancing Policy Worker Group Filter"
				$TableRange = $doc.Application.Selection.Range
				[int]$Columns = 2
				[int]$Rows = $LoadBalancingPolicyConfiguration.WorkerGroupPreferences.count + 1
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$table.Style = $myHash.Word_TableGrid
				$table.Borders.InsideLineStyle = 1
				$table.Borders.OutsideLineStyle = 1
				[int]$xRow = 1
				Write-Verbose "$(Get-Date): `t`t`t`tFormat first row with column headings"
				$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,1).Range.Font.Bold = $True
				$Table.Cell($xRow,1).Range.Text = "Worker Group"
				$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
				$Table.Cell($xRow,2).Range.Font.Bold = $True
				$Table.Cell($xRow,2).Range.Text = "Priority"
				ForEach($WorkerGroupPreference in $LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
				{
					[string]$Tmp = $WorkerGroupPreference
					[string]$WGName = $Tmp.substring($Tmp.indexof("=")+1)
					[string]$WGPriority = $Tmp.substring($Tmp.indexof(":")+1, (($Tmp.indexof("=")-1)-$Tmp.indexof(":")))
					$xRow++
					Write-Verbose "$(Get-Date): `t`t`tProcessing row for Worker Group Filter $($Tmp)"
					$Table.Cell($xRow,1).Range.Text = $WGName
					$Table.Cell($xRow,2).Range.Text = $WGPriority
				}

				Write-Verbose "$(Get-Date): `t`t`t`tMove table to the right"
				$Table.Rows.SetLeftIndent(72,1)
				$table.AutoFitBehavior(1)

				#return focus back to document
				Write-Verbose "$(Get-Date): `t`t`t`tReturn focus back to document"
				$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				Write-Verbose "$(Get-Date): `t`t`t`tMove to the end of the current document"
				$selection.EndKey($wdStory,$wdMove) | Out-Null
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
				"Unknown"                {WriteWordLine 0 0 "Unknown"}
				"ForceServerAccess"      {WriteWordLine 0 0 "Do not allow applications to stream to the client"}
				"ForcedStreamedDelivery" {WriteWordLine 0 0 "Force applications to stream to the client"}
				Default {WriteWordLine 0 0 "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"}
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
	}
}
Elseif($LoadBalancingPolicies -eq $Null)
{
	Write-Verbose "$(Get-Date): There are no Load balancing policies created"
}
Else 
{
	Write-Warning "Load balancing policy information could not be retrieved.  "
}
$LoadBalancingPolicies = $Null
Write-Verbose "$(Get-Date): Finished Processing Load Balancing Policies"
Write-Verbose "$(Get-Date): "

#load evaluators
Write-Verbose "$(Get-Date): Processing Load Evaluators"
Write-Verbose "$(Get-Date): `tSetting summary variables"
[int]$TotalLoadEvaluators = 0

Write-Verbose "$(Get-Date): `tRetrieving Load Evaluators"
$LoadEvaluators = Get-XALoadEvaluator -EA 0 | Sort-Object LoadEvaluatorName

If($?)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Load Evaluators:"
	ForEach($LoadEvaluator in $LoadEvaluators)
	{
		$TotalLoadEvaluators++
		Write-Verbose "$(Get-Date): `t`tProcessing Load Evaluator $($LoadEvaluator.LoadEvaluatorName)"
		WriteWordLine 2 0 $LoadEvaluator.LoadEvaluatorName
		If(![String]::IsNullOrEmpty($LoadEvaluator.Description))
		{
			WriteWordLine 0 1 "Description: " $LoadEvaluator.Description
		}
		
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
			WriteWordLine 0 2 "Report full load when the # of context Switches per second is > than: " $LoadEvaluator.ContextSwitches[1]
			WriteWordLine 0 2 "Report no load when the # of context Switches per second is <= to: " $LoadEvaluator.ContextSwitches[0]
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
			Switch ($LoadEvaluator.LoadThrottling)
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
	Write-Warning "Load Evaluator information could not be retrieved"
}
$LoadEvaluators = $Null
Write-Verbose "$(Get-Date): Finished Processing Load Evaluators"
Write-Verbose "$(Get-Date): "

#servers
Write-Verbose "$(Get-Date): Processing Servers"
Write-Verbose "$(Get-Date): `tSetting summary variables"
[int]$TotalControllers = 0
[int]$TotalWorkers = 0
$ServerItems = @()

Write-Verbose "$(Get-Date): `tRetrieving Servers"
$servers = Get-XAServer -EA 0 | Sort-Object FolderPath, ServerName

If($?)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Servers:"
	ForEach($server in $servers)
	{
		Write-Verbose "$(Get-Date): `t`tProcessing server $($server.ServerName)"

		[bool]$SvrOnline = $False
		Write-Verbose "$(Get-Date): `t`t`tTesting to see if $($server.ServerName) is online and reachable"
		If(Test-Connection -ComputerName $server.servername -quiet -EA 0)
		{
			$SvrOnline = $True
			If($Hardware -and $Software)
			{
				Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online.  Hardware inventory, Software Inventory, Citrix Services and Hotfix areas will be processed."
			}
			ElseIf($Hardware -and !($Software))
			{
				Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online.  Hardware inventory, Citrix Services and Hotfix areas will be processed."
			}
			ElseIf(!($Hardware) -and $Software)
			{
				Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online.  Software Inventory, Citrix Services and Hotfix areas will be processed."
			}
			Else
			{
				Write-Verbose "$(Get-Date): `t`t`t`t$($server.ServerName) is online.  Citrix Services and Hotfix areas will be processed."
			}
		}
		
		#create array for appendix B
		Write-Verbose "$(Get-Date): `t`t`tGather server info for Appendix B"
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
		
		$ServerItems += $obj

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
		Switch ($Server.LogOnMode)
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
		Switch ($server.ElectionPreference)
		{
			"Unknown"           {WriteWordLine 0 0 "Unknown"}
			"MostPreferred"     {WriteWordLine 0 0 "Most Preferred"; $TotalControllers++}
			"Preferred"         {WriteWordLine 0 0 "Preferred"; $TotalControllers++}
			"DefaultPreference" {WriteWordLine 0 0 "Default Preference"; $TotalControllers++}
			"NotPreferred"      {WriteWordLine 0 0 "Not Preferred"; $TotalControllers++}
			"WorkerMode"        {WriteWordLine 0 0 "Worker Mode"; $TotalWorkers++}
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

		If($SvrOnline -and $Hardware)
		{
			GetComputerWMIInfo $server.ServerName
		}
		
		#applications published to server
		$Applications = Get-XAApplication -ServerName $server.ServerName -EA 0 | Sort-Object FolderPath, DisplayName
		If($? -and $Applications)
		{
			WriteWordLine 0 1 "Published applications:"
			Write-Verbose "$(Get-Date): `t`tProcessing published applications for server $($server.ServerName)"
			Write-Verbose "$(Get-Date): `t`tCreate Word Table for server's published applications"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			
			If($Applications -is [Array])
			{
				[int]$Rows = $Applications.count + 1
			}
			Else
			{
				[int]$Rows = 2
			}

			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 1
			$table.Borders.OutsideLineStyle = 1
			[int]$xRow = 1
			Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Display name"
			$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,2).Range.Font.Bold = $True
			$Table.Cell($xRow,2).Range.Text = "Folder path"
			ForEach($app in $Applications)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing published application $($app.DisplayName)"
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $app.DisplayName
				$Table.Cell($xRow,2).Range.Text = $app.FolderPath
			}
			Write-Verbose "$(Get-Date): `t`tMove table of published applications to the right"
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			Write-Verbose "$(Get-Date): `t`tReturn focus back to document"
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			Write-Verbose "$(Get-Date): `t`tMove to the end of the current document"
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			WriteWordLine 0 0 ""
		}

		#get list of applications installed on server
		# original work by Shaun Ritchie
		# modified by Jeff Wouters
		# modified by Webster
		# fixed, as usual, by Michael B. Smith
		If($SvrOnline -and $Software)
		{
			$InstalledApps = @()
			$JustApps = @()

			#Define the variable to hold the location of Currently Installed Programs
			$UninstallKey1="SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 

			#Create an instance of the Registry Object and open the HKLM base key
			$reg=[microsoft.win32.registrykey]::OpenRemoteBaseKey('LocalMachine',$Server.ServerName) 

			#Drill down into the Uninstall key using the OpenSubKey Method
			$regkey1=$reg.OpenSubKey($UninstallKey1) 

			#Retrieve an array of string that contain all the subkey names
			If($regkey1 -ne $Null)
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
						$InstalledApps += $obj
					}
				}
			}			

			$UninstallKey2="SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall" 
			$regkey2=$reg.OpenSubKey($UninstallKey2)
			If($regkey2 -ne $Null)
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
						$InstalledApps += $obj
					}
				}
			}

			$InstalledApps = $InstalledApps | Sort DisplayName

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
			
			$JustApps = $TempApps | Select DisplayName | Sort DisplayName -unique

			WriteWordLine 0 1 "Installed applications:"
			Write-Verbose "$(Get-Date): `t`tProcessing installed applications for server $($server.ServerName)"
			Write-Verbose "$(Get-Date): `t`tCreate Word Table for server's installed applications"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 1
			[int]$Rows = $JustApps.count + 1

			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 1
			$table.Borders.OutsideLineStyle = 1
			[int]$xRow = 1
			Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Application name"
			ForEach($app in $JustApps)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing installed application $($app.DisplayName)"
				$xRow++
				$Table.Cell($xRow,1).Range.Text = $app.DisplayName
			}
			Write-Verbose "$(Get-Date): `t`tMove table of installed applications to the right"
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			Write-Verbose "$(Get-Date): `t`tReturn focus back to document"
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			Write-Verbose "$(Get-Date): `t`tMove to the end of the current document"
			$selection.EndKey($wdStory,$wdMove) | Out-Null
			WriteWordLine 0 0 ""
		}
		
		#list citrix services
		If($SvrOnline)
		{
			Write-Verbose "$(Get-Date): `t`tProcessing Citrix services for server $($server.ServerName) by calling Get-Service"
			$services = get-service -ComputerName $server.ServerName -EA 0 | where-object {$_.DisplayName -like "*Citrix*"} | Sort-Object DisplayName
			WriteWordLine 0 1 "Citrix Services"
			Write-Verbose "$(Get-Date): `t`tCreate Word Table for Citrix services"
			$TableRange = $doc.Application.Selection.Range
			[int]$Columns = 2
			[int]$Rows = $services.count + 1
			Write-Verbose "$(Get-Date): `t`tAdd Citrix services table to doc"
			$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
			$table.Style = $myHash.Word_TableGrid
			$table.Borders.InsideLineStyle = 1
			$table.Borders.OutsideLineStyle = 1
			[int]$xRow = 1
			Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
			$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,1).Range.Font.Bold = $True
			$Table.Cell($xRow,1).Range.Text = "Display Name"
			$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
			$Table.Cell($xRow,2).Range.Font.Bold = $True
			$Table.Cell($xRow,2).Range.Text = "Status"
			ForEach($Service in $Services)
			{
				Write-Verbose "$(Get-Date): `t`t`tProcessing Citrix service $($Service.DisplayName)"
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

			Write-Verbose "$(Get-Date): `t`tMove table of Citrix services to the right"
			$Table.Rows.SetLeftIndent(36,1)
			$table.AutoFitBehavior(1)

			#return focus back to document
			Write-Verbose "$(Get-Date): `t`tReturn focus back to document"
			$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

			#move to the end of the current document
			Write-Verbose "$(Get-Date): `t`tMove to the end of the current document"
			$selection.EndKey($wdStory,$wdMove) | Out-Null

			#Citrix hotfixes installed
			Write-Verbose "$(Get-Date): `t`tGet list of Citrix hotfixes installed using Get-XAServerHotfix"
			$hotfixes = Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | Sort-Object HotfixName
			If($? -and $hotfixes)
			{
				[int]$Rows = 1
				$Single_Row = (Get-Member -Type Property -Name Length -InputObject $hotfixes -EA 0) -eq $Null
				If(-not $Single_Row)
				{
					$Rows = $Hotfixes.length
				}
				$Rows++
				
				Write-Verbose "$(Get-Date): `t`tNumber of hotfixes is $($Rows-1)"
				$HotfixArray = @()
				[bool]$HRP2Installed = $False
				WriteWordLine 0 0 ""
				WriteWordLine 0 1 "Citrix Installed Hotfixes:"
				Write-Verbose "$(Get-Date): `t`tCreate Word Table for Citrix Hotfixes"
				$TableRange = $doc.Application.Selection.Range
				[int]$Columns = 5
				Write-Verbose "$(Get-Date): `t`tAdd Citrix installed hotfix table to doc"
				$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
				$table.Style = $myHash.Word_TableGrid
				$table.Borders.InsideLineStyle = 1
				$table.Borders.OutsideLineStyle = 1
				[int]$xRow = 1
				Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
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
					Write-Verbose "$(Get-Date): `t`t`tProcessing Citrix hotfix $($hotfix.HotfixName)"
					$xRow++
					$HotfixArray += $hotfix.HotfixName
					If($hotfix.HotfixName -eq "XA650W2K8R2X64R02")
					{
						$HRP2Installed = $True
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
				Write-Verbose "$(Get-Date): `t`tMove table of Citrix installed hotfixes to the right"
				$Table.Rows.SetLeftIndent(36,1)
				$table.AutoFitBehavior(1)

				#return focus back to document
				Write-Verbose "$(Get-Date): `t`tReturn focus back to document"
				$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

				#move to the end of the current document
				Write-Verbose "$(Get-Date): `t`tMove to the end of the current document"
				$selection.EndKey($wdStory,$wdMove) | Out-Null
				WriteWordLine 0 0 ""

				#compare Citrix hotfixes to recommended Citrix hotfixes from CTX129229
				#hotfix lists are from CTX129229 dated 16-JUL-2013
				Write-Verbose "$(Get-Date): `t`tCompare Citrix hotfixes to recommended Citrix hotfixes from CTX129229"
				# as of the 16-JUL-2013 update, there are recommended hotfixes for pre R02 and none for post R02
				Write-Verbose "$(Get-Date): `t`tProcessing Citrix hotfix list for server $($server.ServerName)"
				If($HRP2Installed)
				{
					$RecommendedList = @()
				}
				Else
				{
					$RecommendedList = @("XA650W2K8R2X64001","XA650W2K8R2X64011","XA650W2K8R2X64019","XA650W2K8R2X64025","XA650R01W2K8R2X64061", "XA650W2K8R2X64R02")
				}
				If($RecommendedList.count -gt 0)
				{
					WriteWordLine 0 1 "Citrix Recommended Hotfixes:"
					Write-Verbose "$(Get-Date): `t`tCreate Word Table for Citrix Hotfixes"
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					[int]$Rows = $RecommendedList.count + 1
					Write-Verbose "$(Get-Date): `t`tAdd Citrix recommended hotfix table to doc"
					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$table.Style = $myHash.Word_TableGrid
					$table.Borders.InsideLineStyle = 1
					$table.Borders.OutsideLineStyle = 1
					[int]$xRow = 1
					Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
					$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Citrix Hotfix"
					$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Status"
					ForEach($element in $RecommendedList)
					{
						Write-Verbose "$(Get-Date): `t`t`tProcessing Recommended Citrix hotfix $($element)"
						$xRow++
						$Table.Cell($xRow,1).Range.Text = $element
						If(!($HotfixArray -contains $element))
						{
							#missing a recommended Citrix hotfix
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
					Write-Verbose "$(Get-Date): `t`tMove table of Citrix hotfixes to the right"
					$Table.Rows.SetLeftIndent(36,1)
					$table.AutoFitBehavior(1)

					#return focus back to document
					Write-Verbose "$(Get-Date): `t`tReturn focus back to document"
					$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					Write-Verbose "$(Get-Date): `t`tMove to the end of the current document"
					$selection.EndKey($wdStory,$wdMove) | Out-Null
					WriteWordLine 0 0 ""
				}
				#build list of installed Microsoft hotfixes
				Write-Verbose "$(Get-Date): `t`tProcessing Microsoft hotfixes for server $($server.ServerName)"
				[bool]$GotMSHotfixes = $True
				
				Try
				{
					$results = Get-HotFix -computername $Server.ServerName 
					$MSInstalledHotfixes = $results | select-object -Expand HotFixID | Sort-Object HotFixID
					$results = $Null
				}
				
				Catch
				{
					Write-Verbose "$(Get-Date): Get-HotFix failed for $($server.ServerName)"
					$GotMSHotfixes = $False
					Write-Warning "Get-HotFix failed for $($server.ServerName)"
					WriteWordLine 0 0 "Get-HotFix failed for $($server.ServerName)"
					WriteWordLine 0 0 "On $($server.ServerName) you may need to run winmgmt /verifyrepository and winmgmt /salvagerepository"
				}
				
				If($GotMSHotfixes)
				{
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
										"KB2661332", "KB2731847", "KB2748302", "KB2778831", 
										"KB917607", "KB975777", "KB979530", "KB980663", "KB983460")
					}
					
					WriteWordLine 0 1 "Microsoft Recommended Hotfixes:"
					Write-Verbose "$(Get-Date): `t`tCreate Word Table for Microsoft Hotfixes"
					$TableRange = $doc.Application.Selection.Range
					[int]$Columns = 2
					[int]$Rows = $RecommendedList.count + 1
					Write-Verbose "$(Get-Date): `t`tAdd Microsoft hotfix table to doc"
					$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
					$table.Style = $myHash.Word_TableGrid
					$table.Borders.InsideLineStyle = 1
					$table.Borders.OutsideLineStyle = 1
					[int]$xRow = 1
					Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
					$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,1).Range.Font.Bold = $True
					$Table.Cell($xRow,1).Range.Text = "Microsoft Hotfix"
					$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
					$Table.Cell($xRow,2).Range.Font.Bold = $True
					$Table.Cell($xRow,2).Range.Text = "Status"

					$results = @{}
					ForEach($hotfix in $RecommendedList)
					{
						$xRow++
						Write-Verbose "$(Get-Date): `t`t`tProcessing Microsoft hotfix $($hotfix)"
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
					Write-Verbose "$(Get-Date): `t`tMove table of Microsoft hotfixes to the right"
					$Table.Rows.SetLeftIndent(36,1)
					$table.AutoFitBehavior(1)

					#return focus back to document
					Write-Verbose "$(Get-Date): `t`tReturn focus back to document"
					$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

					#move to the end of the current document
					Write-Verbose "$(Get-Date): `t`tMove to the end of the current document"
					$selection.EndKey($wdStory,$wdMove) | Out-Null
					WriteWordLine 0 1 "Not all missing Microsoft hotfixes may be needed for this server"
				}
			}
		}
		Else
		{
			Write-Verbose "$(Get-Date): `t`t$($server.ServerName) is offline or unreachable.  Citrix Services and Hotfix areas skipped."
			WriteWordLine 0 0 "Server $($server.ServerName) was offline or unreachable at "(get-date).ToString()
			WriteWordLine 0 0 "The Citrix Services and Hotfix areas were skipped."
		}
		WriteWordLine 0 0 "" 
		Write-Verbose "$(Get-Date): `tFinished Processing server $($server.ServerName)"
		Write-Verbose "$(Get-Date): "
	}
}
Else 
{
	Write-Warning "Server information could not be retrieved"
}
$servers = $Null
Write-Verbose "$(Get-Date): Finished Processing Servers"
Write-Verbose "$(Get-Date): "

#worker groups
Write-Verbose "$(Get-Date): Processing Worker Groups"
Write-Verbose "$(Get-Date): `tSetting summary variables"
[int]$TotalWGByServerName = 0
[int]$TotalWGByServerGroup = 0
[int]$TotalWGByOU = 0

Write-Verbose "$(Get-Date): `tRetrieving Worker Groups"
$WorkerGroups = Get-XAWorkerGroup -EA 0 | Sort-Object WorkerGroupName

If($? -and $WorkerGroups -ne $Null)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Worker Groups:"
	ForEach($WorkerGroup in $WorkerGroups)
	{
		Write-Verbose "$(Get-Date): `t`tProcessing Worker Group $($WorkerGroup.WorkerGroupName)"
		WriteWordLine 2 0 $WorkerGroup.WorkerGroupName
		If(![String]::IsNullOrEmpty($WorkerGroup.Description))
		{
			WriteWordLine 0 1 "Description: " $WorkerGroup.Description
		}
		WriteWordLine 0 1 "Folder Path: " $WorkerGroup.FolderPath
		If($WorkerGroup.ServerNames)
		{
			$TotalWGByServerName++
			WriteWordLine 0 1 "Farm Servers:"
			$TempArray = $WorkerGroup.ServerNames | Sort-Object
			ForEach($ServerName in $TempArray)
			{
				WriteWordLine 0 2 $ServerName
			}
			$TempArray = $Null
		}
		If($WorkerGroup.ServerGroups)
		{
			$TotalWGByServerGroup++
			WriteWordLine 0 1 "Server Group Accounts:"
			$TempArray = $WorkerGroup.ServerGroups | Sort-Object
			ForEach($ServerGroup in $TempArray)
			{
				WriteWordLine 0 2 $ServerGroup
			}
			$TempArray = $Null
		}
		If($WorkerGroup.OUs)
		{
			$TotalWGByOU++
			WriteWordLine 0 1 "Organizational Units:"
			$TempArray = $WorkerGroup.OUs | Sort-Object
			ForEach($OU in $TempArray)
			{
				WriteWordLine 0 2 $OU
			}
			$TempArray = $Null
		}
		#applications published to worker group
		$Applications = Get-XAApplication -WorkerGroup $WorkerGroup.WorkerGroupName -EA 0 | Sort-Object FolderPath, DisplayName
		If($? -and $Applications)
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
ElseIf($WorkerGroups -eq $Null)
{

	Write-Verbose "$(Get-Date): There are no Worker Groups created"
}
Else 
{
	Write-Warning "Worker Group information could not be retrieved"
}
$WorkerGroups = $Null
Write-Verbose "$(Get-Date): Finished Processing Worker Groups"
Write-Verbose "$(Get-Date): "

#zones
Write-Verbose "$(Get-Date): Processing Zones"
Write-Verbose "$(Get-Date): `tSetting summary variables"
[int]$TotalZones = 0

Write-Verbose "$(Get-Date): `tRetrieving Zones"
$Zones = Get-XAZone -EA 0 | Sort-Object ZoneName
If($?)
{
	$selection.InsertNewPage()
	WriteWordLine 1 0 "Zones:"
	ForEach($Zone in $Zones)
	{
		$TotalZones++
		Write-Verbose "$(Get-Date): `t`tProcessing Zone $($Zone.ZoneName)"
		WriteWordLine 2 0 $Zone.ZoneName
		WriteWordLine 0 1 "Current Data Collector: " $Zone.DataCollector
		$Servers = Get-XAServer -ZoneName $Zone.ZoneName -EA 0 | Sort-Object ElectionPreference, ServerName
		If($?)
		{		
			WriteWordLine 0 1 "Servers in Zone"
	
			ForEach($Server in $Servers)
			{
				WriteWordLine 0 2 "Server Name and Preference: " $server.ServerName -NoNewLine
				WriteWordLine 0 0  " - " -nonewline
				Switch ($server.ElectionPreference)
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
	Write-Warning "Zone information could not be retrieved"
}
$Servers = $Null
$Zones = $Null
Write-Verbose "$(Get-Date): Finished Processing Zones"
Write-Verbose "$(Get-Date): "

[int]$Global:TotalComputerPolicies = 0
[int]$Global:TotalUserPolicies = 0
[int]$Global:TotalIMAPolicies = 0
[int]$Global:TotalADPolicies = 0
[int]$Global:TotalADPoliciesNotProcessed = 0
$ADPoliciesNotProcessed = @()

#if remoting is enabled, the citrix.grouppolicy.commands module does not work with remoting so skip it
If($Remoting)
{
	Write-Warning "Remoting is enabled."
	Write-Warning "The Citrix.GroupPolicy.Commands module does not work with Remoting."
	Write-Warning "Citrix Policy documentation will not take place."
}
Else
{
	#make sure Citrix.GroupPolicy.Commands module is loaded
	If(!(Check-LoadedModule "Citrix.GroupPolicy.Commands"))
	{
		Write-Warning "The Citrix Group Policy module Citrix.GroupPolicy.Commands.psm1 could not be loaded `nPlease see the Prerequisites section in the ReadMe file (https://www.dropbox.com/s/glq4u2p5xte8s6g/XA6_Inventory_V4_ReadMe.rtf). `nCitrix Policy documentation will not take place"
		Write-Verbose "$(Get-Date): "
	}
	Else
	{
		$selection.InsertNewPage()
		WriteWordLine 1 0 "Policies:"
		Write-Verbose "$(Get-Date): Processing Citrix IMA Policies"
		Write-Verbose "$(Get-Date): `tRetrieving IMA Farm Policies"
		ProcessCitrixPolicies	
		Write-Verbose "$(Get-Date): Finished Processing Citrix IMA Policies"
		Write-Verbose "$(Get-Date): "
		
		#thanks to the Citrix Engineering Team for helping me solve processing Citrix AD based Policies
		Write-Verbose "$(Get-Date): See if there are any Citrix AD based policies to process"
		$CtxGPOArray = @()
		$CtxGPOArray = GetCtxGPOsInAD
		If($CtxGPOArray -is [Array] -and $CtxGPOArray.Count -gt 0)
		{
			Write-Verbose "$(Get-Date): There are $($CtxGPOArray.Count) Citrix AD based policies to process"
			
			ForEach($CtxGPO in $CtxGPOArray)
			{
				Write-Verbose "$(Get-Date): Creating ADGpoDrv PSDrive"
				New-PSDrive -Name ADGpoDrv -PSProvider CitrixGroupPolicy -Root \ -DomainGpo $($CtxGPO) -Scope "Global" | out-null
				If(Get-PSDrive ADGpoDrv -EA 0)
				{
					Write-Verbose "$(Get-Date): Processing Citrix AD Policy $($CtxGPO)"
				
					Write-Verbose "$(Get-Date): `tRetrieving AD Policy $($CtxGPO)"
					ProcessCitrixPolicies "ADGpoDrv"
					Write-Verbose "$(Get-Date): Finished Processing Citrix AD Policy $($CtxGPO)"
					Write-Verbose "$(Get-Date): "
				}
				Else
				{
					$ADPoliciesNotProcessed += $CtxGPO
					$Global:TotalADPoliciesNotProcessed++
					Write-Warning "$($CtxGPO) is not readable by this XenApp 6.5 server"
					Write-Warning "$($CtxGPO)  was probably created by an updated Citrix Group Policy Provider"
				}
			}
		
			If($Global:TotalADPoliciesNotProcessed -gt 0)
			{
				Write-Verbose "$(Get-Date): Processing list of Citrix AD Policies not processed"
				$ADPoliciesNotProcessed = $ADPoliciesNotProcessed | Sort -unique
				WriteWordLine 0 0 ""
				WriteWordLine 2 0 "Active Directory Citrix policies that could not be processed:"
				ForEach($Policy in $ADPoliciesNotProcessed)
				{
					Write-Verbose "$(Get-Date): `t Processing skipped Citrix AD policy $($Policy)"
					WriteWordLine 0 1 $Policy
				}
				Write-Verbose "$(Get-Date): Finished processing list of Citrix AD Policies not processed"
				WriteWordLine 0 0 ""
			}

			Write-Verbose "$(Get-Date): Finished Processing Citrix AD Policies"
			Write-Verbose "$(Get-Date): "
		}
		Else
		{
			Write-Verbose "$(Get-Date): There are no Citrix AD based policies to process"
			Write-Verbose "$(Get-Date): "
		}
		Write-Verbose "$(Get-Date): Finished Processing Citrix Policies"
		Write-Verbose "$(Get-Date): "
	}
}

Write-Verbose "$(Get-Date): Create Appendix A Session Sharing Items"
$selection.InsertNewPage()
WriteWordLine 1 0 "Appendix A - Session Sharing Items from CTX159159"
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
#	In addition, a XenApp server can have Session Sharing disable in a registry key
#	To disable session sharing, the following registry key must be present.
#	This information has been added to the Server Appendix B section
#
#	Add the following value to disable this feature (this value does not exist by default):
#	HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Citrix\Wfshell\TWI\:
#	Type: REG_DWORD
#	Value: SeamlessFlags = 1

$TableRange = $doc.Application.Selection.Range
[int]$Columns = 5
[int]$Rows = $SessionSharingItems.count + 1
Write-Verbose "$(Get-Date): `tAdd Session Sharing Items table to doc"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = $myHash.Word_TableGrid
$table.Borders.InsideLineStyle = 1
$table.Borders.OutsideLineStyle = 1
[int]$xRow = 1
Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Application Name"
$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,2).Range.Font.Bold = $True
$Table.Cell($xRow,2).Range.Text = "Maximum color quality"
$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,3).Range.Font.Bold = $True
$Table.Cell($xRow,3).Range.Text = "Session window size"
$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,4).Range.Font.Bold = $True
$Table.Cell($xRow,4).Range.Text = "Access Control Filters"
$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,5).Range.Font.Bold = $True
$Table.Cell($xRow,5).Range.Text = "Encryption"
ForEach($Item in $SessionSharingItems)
{
	$xRow++
	Write-Verbose "$(Get-Date): `t`t`tProcessing row for application $($Item.ApplicationName)"
	$Table.Cell($xRow,1).Range.Text = $Item.ApplicationName
	$Table.Cell($xRow,2).Range.Text = $Item.MaximumColorQuality
	$Table.Cell($xRow,3).Range.Text = $Item.SessionWindowSize
	$Table.Cell($xRow,4).Range.Text = $Item.AccessControlFilters
	$Table.Cell($xRow,5).Range.Text = $Item.Encryption
}

$table.AutoFitBehavior(1)

#return focus back to document
Write-Verbose "$(Get-Date): `t`tReturn focus back to document"
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
Write-Verbose "$(Get-Date): `t`tMove to the end of the current document"
$selection.EndKey($wdStory,$wdMove) | Out-Null
Write-Verbose "$(Get-Date): Finished Create Appendix A - Session Sharing Items"
Write-Verbose "$(Get-Date): "


Write-Verbose "$(Get-Date): Create Appendix B Server Major Items"
$selection.InsertNewPage()
WriteWordLine 1 0 "Appendix B - Server Major Items"
$TableRange = $doc.Application.Selection.Range
[int]$Columns = 7
[int]$Rows = $ServerItems.count + 1
Write-Verbose "$(Get-Date): `tAdd Major Server Items table to doc"
$Table = $doc.Tables.Add($TableRange, $Rows, $Columns)
$table.Style = $myHash.Word_TableGrid
$table.Borders.InsideLineStyle = 1
$table.Borders.OutsideLineStyle = 1
[int]$xRow = 1
Write-Verbose "$(Get-Date): `t`tFormat first row with column headings"
$Table.Cell($xRow,1).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,1).Range.Font.Bold = $True
$Table.Cell($xRow,1).Range.Text = "Server Name"
$Table.Cell($xRow,2).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,2).Range.Font.Bold = $True
$Table.Cell($xRow,2).Range.Text = "Zone Name"
$Table.Cell($xRow,3).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,3).Range.Font.Bold = $True
$Table.Cell($xRow,3).Range.Text = "OS Version"
$Table.Cell($xRow,4).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,4).Range.Font.Bold = $True
$Table.Cell($xRow,4).Range.Text = "Citrix Version"
$Table.Cell($xRow,5).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,5).Range.Font.Bold = $True
$Table.Cell($xRow,5).Range.Text = "Product Edition"
$Table.Cell($xRow,6).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,6).Range.Font.Bold = $True
$Table.Cell($xRow,6).Range.Text = "License Server"
$Table.Cell($xRow,7).Shading.BackgroundPatternColor = $wdColorGray15
$Table.Cell($xRow,7).Range.Font.Bold = $True
$Table.Cell($xRow,7).Range.Text = "Session Sharing"
ForEach($ServerItem in $ServerItems)
{
	$xRow++
	Write-Verbose "$(Get-Date): `t`t`tProcessing row for server $($ServerItem.ServerName)"
	$Table.Cell($xRow,1).Range.Text = $ServerItem.ServerName
	$Table.Cell($xRow,2).Range.Text = $ServerItem.ZoneName
	$Table.Cell($xRow,3).Range.Text = $ServerItem.OSVersion
	$Table.Cell($xRow,4).Range.Text = $ServerItem.CitrixVersion
	$Table.Cell($xRow,5).Range.Text = $ServerItem.ProductEdition
	$Table.Cell($xRow,6).Range.Text = $ServerItem.LicenseServer
	$Table.Cell($xRow,7).Range.Text = $ServerItem.SessionSharing
}

$table.AutoFitBehavior(1)

#return focus back to document
Write-Verbose "$(Get-Date): `t`tReturn focus back to document"
$doc.ActiveWindow.ActivePane.view.SeekView = $wdSeekMainDocument

#move to the end of the current document
Write-Verbose "$(Get-Date): `t`tMove to the end of the current document"
$selection.EndKey($wdStory,$wdMove) | Out-Null
Write-Verbose "$(Get-Date): Finished Create Appendix B - Server Major Items"
Write-Verbose "$(Get-Date): "

#summary page
Write-Verbose "$(Get-Date): Create Summary Page"
$selection.InsertNewPage()
WriteWordLine 1 0 "Summary Page"
Write-Verbose "$(Get-Date): `tAdd administrator summary info"
WriteWordLine 0 0 "Administrators"
WriteWordLine 0 1 "Total Full Administrators`t: " $TotalFullAdmins
WriteWordLine 0 1 "Total View Administrators`t: " $TotalViewAdmins
WriteWordLine 0 1 "Total Custom Administrators`t: " $TotalCustomAdmins
WriteWordLine 0 2 "Total Administrators`t: " ($TotalFullAdmins + $TotalViewAdmins + $TotalCustomAdmins)
WriteWordLine 0 0 ""
Write-Verbose "$(Get-Date): `tAdd application summary info"
WriteWordLine 0 0 "Applications"
WriteWordLine 0 1 "Total Published Applications`t: " $TotalPublishedApps
WriteWordLine 0 1 "Total Published Content`t`t: " $TotalPublishedContent
WriteWordLine 0 1 "Total Published Desktops`t: " $TotalPublishedDesktops
WriteWordLine 0 1 "Total Streamed Applications`t: " $TotalStreamedApps
WriteWordLine 0 2 "Total Applications`t: " ($TotalPublishedApps + $TotalPublishedContent + $TotalPublishedDesktops + $TotalStreamedApps)
WriteWordLine 0 0 ""
Write-Verbose "$(Get-Date): `tAdd configuration logging summary info"
WriteWordLine 0 0 "Configuration Logging"
WriteWordLine 0 1 "Total Config Log Items`t`t: " $TotalConfigLogItems 
WriteWordLine 0 0 ""
Write-Verbose "$(Get-Date): `tAdd load balancing policies summary info"
WriteWordLine 0 0 "Load Balancing Policies"
WriteWordLine 0 1 "Total Load Balancing Policies`t: " $TotalLBPolicies
WriteWordLine 0 0 ""
Write-Verbose "$(Get-Date): `tAdd load evaluator summary info"
WriteWordLine 0 0 "Load Evaluators"
WriteWordLine 0 1 "Total Load Evaluators`t`t: " $TotalLoadEvaluators
WriteWordLine 0 0 ""
Write-Verbose "$(Get-Date): `tAdd server summary info"
WriteWordLine 0 0 "Servers"
WriteWordLine 0 1 "Total Controllers`t`t: " $TotalControllers
WriteWordLine 0 1 "Total Workers`t`t`t: " $TotalWorkers
WriteWordLine 0 2 "Total Servers`t`t: " ($TotalControllers + $TotalWorkers)
WriteWordLine 0 0 ""
Write-Verbose "$(Get-Date): `tAdd worker group summary info"
WriteWordLine 0 0 "Worker Groups"
WriteWordLine 0 1 "Total WGs by Server Name`t: " $TotalWGByServerName
WriteWordLine 0 1 "Total WGs by Server Group`t: " $TotalWGByServerGroup
WriteWordLine 0 1 "Total WGs by AD Container`t: " $TotalWGByOU
WriteWordLine 0 2 "Total Worker Groups`t: " ($TotalWGByServerName + $TotalWGByServerGroup + $TotalWGByOU)
WriteWordLine 0 0 ""
Write-Verbose "$(Get-Date): `tAdd zone summary info"
WriteWordLine 0 0 "Zones"
WriteWordLine 0 1 "Total Zones`t`t`t: " $TotalZones
WriteWordLine 0 0 ""
Write-Verbose "$(Get-Date): `tAdd policy summary info"
WriteWordLine 0 0 "Policies"
WriteWordLine 0 1 "Total Computer Policies`t`t: " $Global:TotalComputerPolicies
WriteWordLine 0 1 "Total User Policies`t`t: " $Global:TotalUserPolicies
WriteWordLine 0 2 "Total Policies`t`t: " ($Global:TotalComputerPolicies + $Global:TotalUserPolicies)
WriteWordLine 0 0 ""
WriteWordLine 0 1 "IMA Policies`t`t`t: " $Global:TotalIMAPolicies
WriteWordLine 0 1 "Citrix AD Policies Processed`t: $($Global:TotalADPolicies)`t(AD Policies can contain multiple Citrix policies)"
WriteWordLine 0 1 "Citrix AD Policies not Processed`t: " $Global:TotalADPoliciesNotProcessed
Write-Verbose "$(Get-Date): Finished Create Summary Page"
Write-Verbose "$(Get-Date): "

Write-Verbose "$(Get-Date): Finishing up Word document"
#end of document processing
#Update document properties

If($CoverPagesExist)
{
	Write-Verbose "$(Get-Date): Set Cover Page Properties"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Company" $CompanyName
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Title" $title
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Subject" "XenApp 6.5 Farm Inventory"
	_SetDocumentProperty $doc.BuiltInDocumentProperties "Author" $username

	#Get the Coverpage XML part
	$cp = $doc.CustomXMLParts | where {$_.NamespaceURI -match "coverPageProps$"}

	#get the abstract XML part
	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "Abstract"}
	#set the text
	[string]$abstract = "Citrix XenApp 6.5 Inventory for $CompanyName"
	$ab.Text = $abstract

	$ab = $cp.documentelement.ChildNodes | Where {$_.basename -eq "PublishDate"}
	#set the text
	[string]$abstract = (Get-Date -Format d).ToString()
	$ab.Text = $abstract

	Write-Verbose "$(Get-Date): Update the Table of Contents"
	#update the Table of Contents
	$doc.TablesOfContents.item(1).Update()
	$cp = $Null
	$ab = $Null
	$abstract = $Null
}

Write-Verbose "$(Get-Date): Save and Close document and Shutdown Word"
If($WordVersion -eq $wdWord2007)
{
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
	}
	Else
	{
		Write-Verbose "$(Get-Date): Saving DOCX file"
	}
	Write-Verbose "$(Get-Date): Running Word 2007 and detected operating system $($RunningOS)"
	If($RunningOS.Contains("Server 2008 R2"))
	{
		$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type] 
		$doc.SaveAs($filename1, $SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$SaveFormat = $wdSaveFormatPDF
			$doc.SaveAs($filename2, $SaveFormat)
		}
	}
	Else
	{
		#works for Server 2008 and Windows 7
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
		$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
		If($PDF)
		{
			Write-Verbose "$(Get-Date): Now saving as PDF"
			$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
			$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
		}
	}
}
Else
{
	#the $saveFormat below passes StrictMode 2
	#I found this at the following two links
	#http://blogs.technet.com/b/bshukla/archive/2011/09/27/3347395.aspx
	#http://msdn.microsoft.com/en-us/library/microsoft.office.interop.word.wdsaveformat(v=office.14).aspx
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Saving as DOCX file first before saving to PDF"
	}
	Else
	{
		Write-Verbose "$(Get-Date): Saving DOCX file"
	}
	$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatDocumentDefault")
	$doc.SaveAs([REF]$filename1, [ref]$SaveFormat)
	If($PDF)
	{
		Write-Verbose "$(Get-Date): Now saving as PDF"
		$saveFormat = [Enum]::Parse([Microsoft.Office.Interop.Word.WdSaveFormat], "wdFormatPDF")
		$doc.SaveAs([REF]$filename2, [ref]$saveFormat)
	}
}

Write-Verbose "$(Get-Date): Closing Word"
$doc.Close()
$Word.Quit()
If($PDF)
{
	Write-Verbose "$(Get-Date): Deleting $($filename1) since only $($filename2) is needed"
	Remove-Item $filename1 -EA 0
}
Write-Verbose "$(Get-Date): System Cleanup"
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Word) | out-null
Remove-Variable -Name word -Scope Global -EA 0
$SaveFormat = $Null
[gc]::collect() 
[gc]::WaitForPendingFinalizers()
Write-Verbose "$(Get-Date): Script has completed"
Write-Verbose "$(Get-Date): "

If($PDF)
{
	Write-Verbose "$(Get-Date): $($filename2) is ready for use"
}
Else
{
	Write-Verbose "$(Get-Date): $($filename1) is ready for use"
}
Write-Verbose "$(Get-Date): "

#http://poshtips.com/measuring-elapsed-time-in-powershell/
Write-Verbose "$(Get-Date): Script started: $($Script:StartTime)"
Write-Verbose "$(Get-Date): Script ended: $(Get-Date)"
$runtime = $(Get-Date) - $Script:StartTime
$Str = [string]::format("{0} days, {1} hours, {2} minutes, {3}.{4} seconds", `
        $runtime.Days, `
        $runtime.Hours, `
        $runtime.Minutes, `
        $runtime.Seconds,
        $runtime.Milliseconds)
Write-Verbose "$(Get-Date): Elapsed time: $($Str)"
$runtime = $Null
$Str = $Null