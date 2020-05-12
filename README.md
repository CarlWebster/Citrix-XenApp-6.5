# XenApp-6.5
Creates an inventory of a Citrix XenApp 6.5 Farm

	Creates an inventory of a Citrix XenApp 6.5 Farm using Microsoft PowerShell, Word,
	PDF, plain text, or HTML.
	
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
