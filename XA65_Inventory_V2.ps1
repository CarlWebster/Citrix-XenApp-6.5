#Original Script created 8/17/2010 by Michael Bogobowicz, Citrix Systems.
#To contact, please message @mcbogo on Twitter
#This script is designed to be run on a XenApp 6.5 server

#Modifications by Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#modified from the original script for XenApp 6.5
#originally released to the Citrix community on October 7, 2011
#update October 9, 2011: fixed the formatting of the Health Monitoring & Recovery policy setting
#update January 9 through 18, 2013:
#	Updated output text to match what is shown in AppCenter
#	Added function and logic to load citrix.grouppolicy.commands module
#	Removed items that never returned data
#	Changed some text labels to shorten the length
#	Policies are now sorted by Type and Priority
#	Figured out how to retrieve all the settings for the Session Printer policy setting
#	Fixed policy filters not working
#	Fixed date display for Reboot schedule start date
#	Fixed time display for:
#		Memory optimization schedule
#		Reboot schedule time
#	Fixed missing policy entries for:
#		Memory optimization exclusion list
#		Offline app users
#		Virtual IP compatibility programs list
#		Virtual IP filter adapter addresses programs list
#		Virtual IP virtual loopback programs list
#		Flash server-side content fetching whitelist
#		Printer driver mapping and compatibility
#		Users who can shadow other users
#		Users who cannot shadow other users
#		Client USB device redirection rules
#update January 17, 2013:
#	updated Function Check-LoadedModule with an improvement suggested by @andyjmorgan
#	added by @andyjmorgan, checking for required Citrix PoSH snap-ins
#	added by @andyjmorgan, changed text output when the citrix.grouppolicy.commands module does not exist
#update January 21, 2013
#	bug reported and fixed by @schose in Function Check-LoadedModule

Function line
#function created by Michael B. Smith, Exchange MVP
#@essentialexchange on Twitter
#http://TheEssentialExchange.com
{
	Param( [int]$tabs = 0, [string]$name = ’’, [string]$value = ’’, [string]$newline = “`n”, [switch]$nonewline )

	While( $tabs –gt 0 ) { $global:output += “`t”; $tabs--; }

	If( $nonewline )
	{
		$global:output += $name + $value
	}
	Else
	{
		$global:output += $name + $value + $newline
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
	line 0 "$($hour):$($tempminute) $($AMorPM)"
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
		$module = Import-Module -Name $ModuleName –PassThru –EA 0
		If( $module –and $? )
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
    
    #Creates arrays of strings, rather than objects, we're passing strings so this will be more robust.
    $loadedSnapins += get-pssnapin | % {$_.name}
    $registeredSnapins += get-pssnapin -Registered | % {$_.name}
    
    
    foreach ($Snapin in $Snapins){
        #check if the snapin is loaded
        if (!($LoadedSnapins -like $snapin)){

            #Check if the snapin is missing
            if (!($RegisteredSnapins -like $Snapin)){

                #set the flag if it's not already
                if (!($FoundMissingSnapin)){
                    $FoundMissingSnapin = $True
                }
                
                #add the entry to the list
                $MissingSnapins += $Snapin
            }#End Registered If 
            
            Else{
                #Snapin is registered, but not loaded, loading it now:
                Write-Host "Loading Windows PowerShell snap-in: $snapin"
                Add-PSSnapin -Name $snapin
            }
            
        }#End Loaded If
        #Snapin is registered and loaded
        else{write-debug "Windows PowerShell snap-in: $snapin - Already Loaded"}
    }#End For
    
    if ($FoundMissingSnapin){
        write-warning "Missing Windows PowerShell snap-ins Detected:"
        $missingSnapins | % {write-warning "($_)"}
        return $False
    }#End If
    
    Else{
        Return $true
    }#End Else
    
}#End Function

#Script begins
$global:output = ""

if (!(Check-NeededPSSnapins "Citrix.Common.Commands","Citrix.Common.GroupPolicy","Citrix.XenApp.Commands")){
    #We're missing Citrix Snapins that we need
    write-error "Missing Citrix PowerShell Snap-ins Detected, check the console above for more information. Are you sure you are running this script on a XenApp 6.5 Server? Script will now close."
    break
}

# Get farm information
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
	line 0 "Farm: "$farm.FarmName
} 
Else 
{
	line 0 "Farm information could not be retrieved"
}
Write-Output $global:output
$farm = $null
$global:output = $null

$global:ConfigLog = $False
$ConfigurationLogging = Get-XAConfigurationLog -EA 0

If( $? )
{
	If ($ConfigurationLogging.LoggingEnabled ) 
	{
		$global:ConfigLog = $True
		line 0 ""
		line 0 "Configuration Logging is enabled."
		line 1 "Allow changes to the farm when logging database is disconnected: " $ConfigurationLogging.ChangesWhileDisconnectedAllowed
		line 1 "Require administrator to enter credentials before clearing the log: " $ConfigurationLogging.CredentialsOnClearLogRequired
		line 1 "Database type: " $ConfigurationLogging.DatabaseType
		line 1 "Authentication mode: " $ConfigurationLogging.AuthenticationMode
		line 1 "Connection string: " 
		$Tmp = "`t`t" + $ConfigurationLogging.ConnectionString.replace(";","`n`t`t`t")
		line 1 $Tmp -NoNewline
		line 0 ""
		line 1 "User name: " $ConfigurationLogging.UserName
		$Tmp = $null
	}
	Else 
	{
		line 0 ""
		line 0 "Configuration Logging is disabled."
	}
}
Else 
{
	line 0 "Configuration Logging could not be retrieved"
}
Write-Output $global:output
$ConfigurationLogging = $null
$global:output = $null

$Administrators = Get-XAAdministrator -EA 0 | sort-object AdministratorName

If( $? )
{
	line 0 ""
	line 0 "Administrators:"
	ForEach($Administrator in $Administrators)
	{
		line 0 ""
		line 1 "Administrator name: "$Administrator.AdministratorName
		line 1 "Administrator type: " -nonewline
		switch ($Administrator.AdministratorType)
		{
			"Unknown"  {line 0 "Unknown"}
			"Full"     {line 0 "Full Administration"}
			"ViewOnly" {line 0 "View Only"}
			"Custom"   {line 0 "Custom"}
			Default    {line 0 "Administrator type could not be determined: $($Administrator.AdministratorType)"}
		}
		line 1 "Administrator account is " -NoNewLine
		If($Administrator.Enabled)
		{
			line 0 "Enabled" 
		} 
		Else
		{
			line 0 "Disabled" 
		}
		If ($Administrator.AdministratorType -eq "Custom") 
		{
			line 1 "Farm Privileges:"
			ForEach($farmprivilege in $Administrator.FarmPrivileges) 
			{
				switch ($farmprivilege)
				{
					"Unknown"                   {line 2 "Unknown"}
					"ViewFarm"                  {line 2 "View farm management"}
					"EditZone"                  {line 2 "Edit zones"}
					"EditConfigurationLog"      {line 2 "Configure logging for the farm"}
					"EditFarmOther"             {line 2 "Edit all other farm settings"}
					"ViewAdmins"                {line 2 "View Citrix administrators"}
					"LogOnConsole"              {line 2 "Log on to console"}
					"LogOnWIConsole"            {line 2 "Logon on to Web Interface console"}
					"ViewLoadEvaluators"        {line 2 "View load evaluators"}
					"AssignLoadEvaluators"      {line 2 "Assign load evaluators"}
					"EditLoadEvaluators"        {line 2 "Edit load evaluators"}
					"ViewLoadBalancingPolicies" {line 2 "View load balancing policies"}
					"EditLoadBalancingPolicies" {line 2 "Edit load balancing policies"}
					"ViewPrinterDrivers"        {line 2 "View printer drivers"}
					"ReplicatePrinterDrivers"   {line 2 "Replicate printer drivers"}
					Default {line 2 "Farm privileges could not be determined: $($farmprivilege)"}
				}
			}
	
			line 1 "Folder Privileges:"
			ForEach($folderprivilege in $Administrator.FolderPrivileges) 
			{
				#The Citrix PoSH cmdlet only returns data for three folders:
				#Servers
				#WorkerGroups
				#Applications
				
				line 2 $FolderPrivilege.FolderPath
				ForEach($FolderPermission in $FolderPrivilege.FolderPrivileges)
				{
					switch ($folderpermission)
					{
						"Unknown"                          {line 3 "Unknown"}
						"ViewApplications"                 {line 3 "View applications"}
						"EditApplications"                 {line 3 "Edit applications"}
						"TerminateProcessApplication"      {line 3 "Terminate process that is created as a result of launching a published application"}
						"AssignApplicationsToServers"      {line 3 "Assign applications to servers"}
						"ViewServers"                      {line 3 "View servers"}
						"EditOtherServerSettings"          {line 3 "Edit other server settings"}
						"RemoveServer"                     {line 3 "Remove a bad server from farm"}
						"TerminateProcess"                 {line 3 "Terminate processes on a server"}
						"ViewSessions"                     {line 3 "View ICA/RDP sessions"}
						"ConnectSessions"                  {line 3 "Connect sessions"}
						"DisconnectSessions"               {line 3 "Disconnect sessions"}
						"LogOffSessions"                   {line 3 "Log off sessions"}
						"ResetSessions"                    {line 3 "Reset sessions"}
						"SendMessages"                     {line 3 "Send messages to sessions"}
						"ViewWorkerGroups"                 {line 3 "View worker groups"}
						"AssignApplicationsToWorkerGroups" {line 3 "Assign applications to worker groups"}
						Default {line 3 "Folder permission could not be determined: $($folderpermissions)"}
					}
				}
			}
		}		
	
	Write-Output $global:output
	$global:output = $null
	}
}
Else 
{
	line 0 "Administrator information could not be retrieved"
	Write-Output $global:output
}

$Administrators = $null
$global:outout = $null

$Applications = Get-XAApplication -EA 0 | sort-object FolderPath, DisplayName

If( $? -and $Applications)
{
	line 0 ""
	line 0 "Applications:"
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
		line 0 ""
		line 1 "Display name: " $Application.DisplayName
		line 2 "Application name         : " $Application.BrowserName
		line 2 "Disable application      : " -NoNewLine
		#weird, if application is enabled, it is disabled!
		If ($Application.Enabled) 
		{
			line 0 "No"
		} 
		Else
		{
			line 0 "Yes"
			line 2 "Hide disabled application: " -nonewline
			If($Application.HideWhenDisabled)
			{
				line 0 "Yes"
			}
			Else
			{
				line 0 "No"
			}
		}

		If(![String]::IsNullOrEmpty( $Application.Description))
		{
			line 2 "Application description  : " $Application.Description
		}
		
		#type properties
		line 2 "Application Type         : " -nonewline
		switch ($Application.ApplicationType)
		{
			"Unknown"                            {line 0 "Unknown"}
			"ServerInstalled"                    {line 0 "Installed application"}
			"ServerDesktop"                      {line 0 "Server desktop"}
			"Content"                            {line 0 "Content"}
			"StreamedToServer"                   {line 0 "Streamed to server"}
			"StreamedToClient"                   {line 0 "Streamed to client"}
			"StreamedToClientOrInstalled"        {line 0 "Streamed if possible, otherwise accessed from server as Installed application"}
			"StreamedToClientOrStreamedToServer" {line 0 "Streamed if possible, otherwise Streamed to server"}
			Default {line 0 "Application Type could not be determined: $($Application.ApplicationType)"}
		}
		If(![String]::IsNullOrEmpty( $Application.FolderPath))
		{
			line 2 "Folder path              : " $Application.FolderPath
		}
		If(![String]::IsNullOrEmpty( $Application.ContentAddress))
		{
			line 2 "Content Address          : " $Application.ContentAddress
		}
	
		#if a streamed app
		If($streamedapp)
		{
			line 2 "Citrix streaming app profile address         : " 
			line 3 $Application.ProfileLocation
			line 2 "App to launch from Citrix stream app profile : " 
			line 3 $Application.ProfileProgramName
			If(![String]::IsNullOrEmpty( $Application.ProfileProgramArguments))
			{
				line 2 "Extra command line parameters                : " 
				line 3 $Application.ProfileProgramArguments
			}
			#if streamed, Offline access properties
			If($Application.OfflineAccessAllowed)
			{
				line 2 "Enable offline access                        : " -nonewline
				If($Application.OfflineAccessAllowed)
				{
					line 0 "Yes"
				}
				Else
				{
					line 0 "No"
				}
			}
			If($Application.CachingOption)
			{
				line 2 "Cache preference                             : " -nonewline
				switch ($Application.CachingOption)
				{
					"Unknown"   {line 0 "Unknown"}
					"PreLaunch" {line 0 "Cache application prior to launching"}
					"AtLaunch"  {line 0 "Cache application during launch"}
					Default {line 0 "Application Cache prefeence could not be determined: $($Application.CachingOption)"}
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
					line 2 "Command line             : " $Application.CommandLineExecutable
				}
				Else
				{
					line 2 "Command line             : " 
					line 3 $Application.CommandLineExecutable
				}
			}
			If(![String]::IsNullOrEmpty( $Application.WorkingDirectory))
			{
				If($Application.WorkingDirectory.Length -lt 40)
				{
					line 2 "Working directory        : " $Application.WorkingDirectory
				}
				Else
				{
					line 2 "Working directory        : " 
					line 3 $Application.WorkingDirectory
				}
			}
			
			#servers properties
			If($AppServerInfoResults)
			{
				If(![String]::IsNullOrEmpty( $AppServerInfo.ServerNames))
				{
					line 2 "Servers:"
					ForEach($servername in $AppServerInfo.ServerNames)
					{
						line 3 $servername
					}
				}
				If(![String]::IsNullOrEmpty($AppServerInfo.WorkerGroupNames))
				{
					line 2 "Workergroups:"
					ForEach($workergroup in $AppServerInfo.WorkerGroupNames)
					{
						line 3 $workergroup
					}
				}
			}
			Else
			{
				line 3 "Unable to retrieve a list of Servers or Worker Groups for this application"
			}
		}
	
		#users properties
		If($Application.AnonymousConnectionsAllowed)
		{
			line 2 "Allow anonymous users    : " $Application.AnonymousConnectionsAllowed
		}
		Else
		{
			If($AppServerInfoResults)
			{
				line 2 "Users:"
				ForEach($user in $AppServerInfo.Accounts)
				{
					line 3 $user
				}
			}
			Else
			{
				line 3 "Unable to retrieve a list of Users for this application"
			}
		}	

		#shortcut presentation properties
		#application icon is ignored
		If(![String]::IsNullOrEmpty($Application.ClientFolder))
		{
			If($Application.ClientFolder.Length -lt 30)
			{
				line 2 "Client application folder                    : " $Application.ClientFolder
			}
			Else
			{
				line 2 "Client application folder                    : " 
				line 3 $Application.ClientFolder
			}
		}
		If($Application.AddToClientStartMenu)
		{
			line 2 "Add to client's start menu"
			If($Application.StartMenuFolder)
			{
				line 3 "Start menu folder: " $Application.StartMenuFolder
			}
		}
		If($Application.AddToClientDesktop)
		{
			line 2 "Add shortcut to the client's desktop "
		}
	
		#access control properties
		If($Application.ConnectionsThroughAccessGatewayAllowed)
		{
			line 2 "Allow connections made through AGAE          : " -nonewline
			If($Application.ConnectionsThroughAccessGatewayAllowed)
			{
				line 0 "Yes"
			} 
			Else
			{
				line 0 "No"
			}
		}
		If($Application.OtherConnectionsAllowed)
		{
			line 2 "Any connection                               : " -nonewline
			If($Application.OtherConnectionsAllowed)
			{
				line 0 "Yes"
			} 
			Else
			{
				line 0 "No"
			}
		}
		If($Application.AccessSessionConditionsEnabled)
		{
			line 2 "Any connection that meets any of the following filters: " $Application.AccessSessionConditionsEnabled
			line 2 "Access Gateway Filters:"
			ForEach($filter in $Application.AccessSessionConditions)
			{
				line 3 $filter
			}
		}
	
		#content redirection properties
		If($AppServerInfoResults)
		{
			If($AppServerInfo.FileTypes)
			{
				line 2 "File type associations:"
				ForEach($filetype in $AppServerInfo.FileTypes)
				{
					line 3 $filetype
				}
			}
			Else
			{
				line 2 "File Type Associations for this application  : None"
			}
		}
		Else
		{
			line 2 "Unable to retrieve the list of FTAs for this application"
		}
	
		#if streamed app, Alternate profiles
		If($streamedapp)
		{
			If($Application.AlternateProfiles)
			{
				line 2 "Primary application profile location         : " $Application.AlternateProfiles
			}
		
			#if streamed app, User privileges properties
			If($Application.RunAsLeastPrivilegedUser)
			{
				line 2 "Run app as a least-privileged user account   : " $Application.RunAsLeastPrivilegedUser
			}
		}
	
		#limits properties
		line 2 "Limit instances allowed to run in server farm: " -NoNewLine

		If($Application.InstanceLimit -eq -1)
		{
			line 0 "No limit set"
		}
		Else
		{
			line 0 $Application.InstanceLimit
		}
	
		line 2 "Allow only 1 instance of app for each user   : " -NoNewLine
	
		If ($Application.MultipleInstancesPerUserAllowed) 
		{
			line 0 "No"
		} 
		Else
		{
			line 0 "Yes"
		}
	
		If($Application.CpuPriorityLevel)
		{
			line 2 "Application importance                       : " -nonewline
			switch ($Application.CpuPriorityLevel)
			{
				"Unknown"     {line 0 "Unknown"}
				"BelowNormal" {line 0 "Below Normal"}
				"Low"         {line 0 "Low"}
				"Normal"      {line 0 "Normal"}
				"AboveNormal" {line 0 "Above Normal"}
				"High"        {line 0 "High"}
				Default {line 0 "Application importance could not be determined: $($Application.CpuPriorityLevel)"}
			}
		}
		
		#client options properties
		line 2 "Enable legacy audio                          : " -nonewline
		switch ($Application.AudioType)
		{
			"Unknown" {line 0 "Unknown"}
			"None"    {line 0 "Not Enabled"}
			"Basic"   {line 0 "Enabled"}
			Default {line 0 "Enable legacy audio could not be determined: $($Application.AudioType)"}
		}
		line 2 "Minimum requirement                          : " -nonewline
		If($Application.AudioRequired)
		{
			line 0 "Enabled"
		}
		Else
		{
			line 0 "Disabled"
		}
		If($Application.SslConnectionEnable)
		{
			line 2 "Enable SSL and TLS protocols                 : " -nonewline
			If($Application.SslConnectionEnabled)
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}
		}
		If($Application.EncryptionLevel)
		{
			line 2 "Encryption                                   : " -nonewline
			switch ($Application.EncryptionLevel)
			{
				"Unknown" {line 0 "Unknown"}
				"Basic"   {line 0 "Basic"}
				"LogOn"   {line 0 "128-Bit Login Only (RC-5)"}
				"Bits40"  {line 0 "40-Bit (RC-5)"}
				"Bits56"  {line 0 "56-Bit (RC-5)"}
				"Bits128" {line 0 "128-Bit (RC-5)"}
				Default {line 0 "Encryption could not be determined: $($Application.EncryptionLevel)"}
			}
		}
		If($Application.EncryptionRequired)
		{
			line 2 "Minimum requirement                          : " -nonewline
			If($Application.EncryptionRequired)
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}
		}
	
		line 2 "Start app w/o waiting for printer creation   : " -NoNewLine
		#another weird one, if True then this is Disabled
		If ($Application.WaitOnPrinterCreation) 
		{
			line 0 "Disabled"
		} 
		Else
		{
			line 0 "Enabled"
		}
		
		#appearance properties
		If($Application.WindowType)
		{
			line 2 "Session window size                          : " $Application.WindowType
		}
		If($Application.ColorDepth)
		{
			line 2 "Maximum color quality                        : " -nonewline
			switch ($Application.ColorDepth)
			{
				"Unknown"     {line 0 "Unknown color depth"}
				"Colors8Bit"  {line 0 "256-color (8-bit)"}
				"Colors16Bit" {line 0 "Better Speed (16-bit)"}
				"Colors32Bit" {line 0 "Better Appearance (32-bit)"}
				Default {line 0 "Maximum color quality could not be determined: $($Application.ColorDepth)"}
			}
		}
		If($Application.TitleBarHidden)
		{
			line 2 "Hide application title bar                   : " -nonewline
			If($Application.TitleBarHidden)
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}
		}
		If($Application.MaximizedOnStartup)
		{
			line 2 "Maximize application at startup              : " -nonewline
			If($Application.MaximizedOnStartup)
			{
				line 0 "Enabled"
			}
			Else
			{
				line 0 "Disabled"
			}
		}
	
	Write-Output $global:output
	$global:output = $null
	$AppServerInfo = $null
	}
}
Else 
{
	line 0 "Application information could not be retrieved"
}

$Applications = $null
$global:output = $null

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
			line 0 ""
			line 0 "History:"
			ForEach($ConfigLogItem in $ConfigLogReport)
			{
				line 0 ""
				Line 1 "Date              : " $ConfigLogItem.Date
				Line 1 "Account           : " $ConfigLogItem.Account
				Line 1 "Change description: " $ConfigLogItem.Description
				Line 1 "Type of change    : " $ConfigLogItem.TaskType
				Line 1 "Type of item      : " $ConfigLogItem.ItemType
				Line 1 "Name of item      : " $ConfigLogItem.ItemName
			}
			Write-Output $global:output
			$global:output = $null
		} 
		Else 
		{
			line 0 "History information could not be retrieved"
		}
		Write-Output $global:output
		$ConnectionString = $null
		$ConfigLogReport = $null
		$global:output = $null
	}
	Else 
	{
		line 0 "XA65ConfigLog.udl file was not found"
	}
}

#load balancing policies
$LoadBalancingPolicies = Get-XALoadBalancingPolicy -EA 0 | sort-object PolicyName

If( $? -and $LoadBalancingPolicies)
{
	line 0 ""
	line 0 "Load Balancing Policies:"
	ForEach($LoadBalancingPolicy in $LoadBalancingPolicies)
	{
		$LoadBalancingPolicyConfiguration = Get-XALoadBalancingPolicyConfiguration -PolicyName $LoadBalancingPolicy.PolicyName
		$LoadBalancingPolicyFilter = Get-XALoadBalancingPolicyFilter -PolicyName $LoadBalancingPolicy.PolicyName 
	
		line 1 "Name: " $LoadBalancingPolicy.PolicyName
		line 2 "Description: " $LoadBalancingPolicy.Description
		line 2 "Enabled    : " -nonewline
		If($LoadBalancingPolicy.Enabled)
		{
			line 0 "Yes"
		}
		Else
		{
			line 0 "No"
		}
		line 2 "Priority   : " $LoadBalancingPolicy.Priority
	
		line 2 "Filter based on Access Control: " -nonewline
		If($LoadBalancingPolicyFilter.AccessControlEnabled)
		{
			line 0 "Yes"
		}
		Else
		{
			line 0 "No"
		}
		If($LoadBalancingPolicyFilter.AccessControlEnabled)
		{
			line 2 "Apply to connections made through Access Gateway: " -nonewline
			If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
			{
				line 0 "Yes"
			}
			Else
			{
				line 0 "No"
			}
			If($LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway)
			{
				If($LoadBalancingPolicyFilter.AllowOtherConnections)
				{
					line 3 "Any connection"
				} 
				Else
				{
					line 3 "Any connection that meets any of the following filters"
					If($LoadBalancingPolicyFilter.AccessSessionConditions)
					{
						ForEach($AccessSessionCondition in $LoadBalancingPolicyFilter.AccessSessionConditions)
						{
							line 4 $AccessSessionCondition
						}
					}
				}
			}
		}
	
		If($LoadBalancingPolicyFilter.ClientIPAddressEnabled)
		{
			line 2 "Filter based on client IP address"
			If($LoadBalancingPolicyFilter.ApplyToAllClientIPAddresses)
			{
				line 3 "Apply to all client IP addresses"
			} 
			Else
			{
				If($LoadBalancingPolicyFilter.AllowedIPAddresses)
				{
					ForEach($AllowedIPAddress in $LoadBalancingPolicyFilter.AllowedIPAddresses)
					{
						line 3 "Client IP Address Matched: " $AllowedIPAddress
					}
				}
				If($LoadBalancingPolicyFilter.DeniedIPAddresses)
				{
					ForEach($DeniedIPAddress in $LoadBalancingPolicyFilter.DeniedIPAddresses)
					{
						line 3 "Client IP Address Ignored: " $DeniedIPAddress
					}
				}
			}
		}
		If($LoadBalancingPolicyFilter.ClientNameEnabled)
		{
			line 2 "Filter based on client name"
			If($LoadBalancingPolicyFilter.ApplyToAllClientNames)
			{
				line 3 "Apply to all client names"
			} 
			Else
			{
				If($LoadBalancingPolicyFilter.AllowedClientNames)
				{
					ForEach($AllowedClientName in $LoadBalancingPolicyFilter.AllowedClientNames)
					{
						line 3 "Client Name Matched: " $AllowedClientName
					}
				}
				If($LoadBalancingPolicyFilter.DeniedClientNames)
				{
					ForEach($DeniedClientName in $LoadBalancingPolicyFilter.DeniedClientNames)
					{
						line 3 "Client Name Ignored: " $DeniedClientName
					}
				}
			}
		}
		If($LoadBalancingPolicyFilter.AccountEnabled)
		{
			line 2 "Filter based on user"
			line 3 "Apply to anonymous users: " -nonewline
			If($LoadBalancingPolicyFilter.ApplyToAnonymousAccounts)
			{
				line 0 "Yes"
			}
			Else
			{
				line 0 "No"
			}
			If($LoadBalancingPolicyFilter.ApplyToAllExplicitAccounts)
			{
				line 3 "Apply to all explicit (non-anonymous) users"
			} 
			Else
			{
				If($LoadBalancingPolicyFilter.AllowedAccounts)
				{
					ForEach($AllowedAccount in $LoadBalancingPolicyFilter.AllowedAccounts)
					{
						line 3 "User Matched: " $AllowedAccount
					}
				}
				If($LoadBalancingPolicyFilter.DeniedAccounts)
				{
					ForEach($DeniedAccount in $LoadBalancingPolicyFilter.DeniedAccounts)
					{
						line 3 "User Ignored: " $DeniedAccount
					}
				}
			}
		}
		If($LoadBalancingPolicyConfiguration.WorkerGroupPreferenceAndFailoverState)
		{
			line 2 "Configure application connection preference based on worker group"
			If($LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
			{
				ForEach($WorkerGroupPreference in $LoadBalancingPolicyConfiguration.WorkerGroupPreferences)
				{
					line 3 "Worker Group: " $WorkerGroupPreference
				}
			}
		}
		If($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Enabled")
		{
			line 2 "Set the delivery protocols for applications streamed to client"
			line 3 "" -nonewline
			switch ($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)
			{
				"Unknown"                {line 0 "Unknown"}
				"ForceServerAccess"      {line 0 "Do not allow applications to stream to the client"}
				"ForcedStreamedDelivery" {line 0 "Force applications to stream to the client"}
				Default {line 0 "Delivery protocol could not be determined: $($LoadBalancingPolicyConfiguration.StreamingDeliveryOption)"}
			}
		}
		Elseif($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState -eq "Disabled")
		{
			#In the GUI, if "Set the delivery protocols for applications streamed to client" IS selected AND 
			#"Allow applications to stream to the client or run on a Terminal Server (default)" IS selected
			#then "Set the delivery protocols for applications streamed to client" is set to Disabled
			line 2 "Set the delivery protocols for applications streamed to client"
			line 3 "Allow applications to stream to the client or run on a Terminal Server (default)"
		}
		Else
		{
			line 2 "Streamed App Delivery is not configured"
		}
	
		Write-Output $global:output
		$global:output = $null
		$LoadBalancingPolicyConfiguration = $null
		$LoadBalancingPolicyFilter = $null
	}
}
Else 
{
	line 0 "Load balancing policy information could not be retrieved"
}
$LoadBalancingPolicies = $null
$global:output = $null

#load evaluators
$LoadEvaluators = Get-XALoadEvaluator -EA 0 | sort-object LoadEvaluatorName

If( $? )
{
	line 0 ""
	line 0 "Load Evaluators:"
	ForEach($LoadEvaluator in $LoadEvaluators)
	{
		line 1 "Name: " $LoadEvaluator.LoadEvaluatorName
		line 2 "Description: " $LoadEvaluator.Description
		
		If($LoadEvaluator.IsBuiltIn)
		{
			line 2 "Built-in Load Evaluator"
		} 
		Else 
		{
			line 2 "User created load evaluator"
		}
	
		If($LoadEvaluator.ApplicationUserLoadEnabled)
		{
			line 2 "Application User Load Settings"
			line 3 "Report full load when the number of users for this application equals: " $LoadEvaluator.ApplicationUserLoad
			line 3 "Application: " $LoadEvaluator.ApplicationBrowserName
		}
	
		If($LoadEvaluator.ContextSwitchesEnabled)
		{
			line 2 "Context Switches Settings"
			line 3 "Report full load when the number of context switches per second is > than: " $LoadEvaluator.ContextSwitches[1]
			line 3 "Report no load when the number of context switches per second is <= to   : " $LoadEvaluator.ContextSwitches[0]
		}
	
		If($LoadEvaluator.CpuUtilizationEnabled)
		{
			line 2 "CPU Utilization Settings"
			line 3 "Report full load when the processor utilization % is > than: " $LoadEvaluator.CpuUtilization[1]
			line 3 "Report no load when the processor utilization % is <= to   : " $LoadEvaluator.CpuUtilization[0]
		}
	
		If($LoadEvaluator.DiskDataIOEnabled)
		{
			line 2 "Disk Data I/O Settings"
			line 3 "Report full load when the total disk I/O in kbps is > than        : " $LoadEvaluator.DiskDataIO[1]
			line 3 "Report no load when the total disk I/O in kbps per second is <= to: " $LoadEvaluator.DiskDataIO[0]
		}
	
		If($LoadEvaluator.DiskOperationsEnabled)
		{
			line 2 "Disk Operations Settings"
			line 3 "Report full load when the total number of read & write operations per second is > than: " $LoadEvaluator.DiskOperations[1]
			line 3 "Report no load when the total number of read & write operations per second is <= to   : " $LoadEvaluator.DiskOperations[0]
		}
	
		If($LoadEvaluator.IPRangesEnabled)
		{
			line 2 "IP Range Settings"
			If($LoadEvaluator.IPRangesAllowed)
			{
				line 3 "Allow " -NoNewLine
			} 
			Else 
			{
				line 3 "Deny " -NoNewLine
			}
			line 0 "client connections from the listed IP Ranges"
			ForEach($IPRange in $LoadEvaluator.IPRanges)
			{
				line 4 "IP Address Ranges: " $IPRange
			}
		}
	
		If($LoadEvaluator.LoadThrottlingEnabled)
		{
			line 2 "Load Throttling Settings"
			line 3 "Impact of logons on load: " -nonewline
			switch ($LoadEvaluator.LoadThrottling)
			{
				"Unknown"    {line 0 "Unknown"}
				"Extreme"    {line 0 "Extreme"}
				"High"       {line 0 "High (Default)"}
				"MediumHigh" {line 0 "Medium High"}
				"Medium"     {line 0 "Medium"}
				"MediumLow"  {line 0 "Medium Low"}
				Default {line 0 "Impact of logons on load could not be determined: $($LoadEvaluator.LoadThrottling)"}
			}
		}
	
		If($LoadEvaluator.MemoryUsageEnabled)
		{
			line 2 "Memory Usage Settings"
			line 3 "Report full load when the memory usage is > than: " $LoadEvaluator.MemoryUsage[1]
			line 3 "Report no load when the memory usage is <= to   : " $LoadEvaluator.MemoryUsage[0]
		}
	
		If($LoadEvaluator.PageFaultsEnabled)
		{
			line 2 "Page Faults Settings"
			line 3 "Report full load when the number of page faults per second is > than: " $LoadEvaluator.PageFaults[1]
			line 3 "Report no load when the number of page faults per second is <= to   : " $LoadEvaluator.PageFaults[0]
		}
	
		If($LoadEvaluator.PageSwapsEnabled)
		{
			line 2 "Page Swaps Settings"
			line 3 "Report full load when the number of page swaps per second is > than: " $LoadEvaluator.PageSwaps[1]
			line 3 "Report no load when the number of page swaps per second is <= to   : " $LoadEvaluator.PageSwaps[0]
		}
	
		If($LoadEvaluator.ScheduleEnabled)
		{
			line 2 "Scheduling Settings"
			line 3 "Sunday Schedule   : " $LoadEvaluator.SundaySchedule
			line 3 "Monday Schedule   : " $LoadEvaluator.MondaySchedule
			line 3 "Tuesday Schedule  : " $LoadEvaluator.TuesdaySchedule
			line 3 "Wednesday Schedule: " $LoadEvaluator.WednesdaySchedule
			line 3 "Thursday Schedule : " $LoadEvaluator.ThursdaySchedule
			line 3 "Friday Schedule   : " $LoadEvaluator.FridaySchedule
			line 3 "Saturday Schedule : " $LoadEvaluator.SaturdaySchedule
		}
	
		If($LoadEvaluator.ServerUserLoadEnabled)
		{
			line 2 "Server User Load Settings"
			line 3 "Report full load when the number of server users equals: " $LoadEvaluator.ServerUserLoad
		}
	
		line 0 ""
		Write-Output $global:output
		$global:output = $null
	}
}
Else 
{
	line 0 "Load Evaluator information could not be retrieved"
}
$LoadEvaluators = $null
$global:output = $null

#servers
$servers = Get-XAServer -EA 0 | sort-object FolderPath, ServerName

If( $? )
{
	line 0 ""
	line 0 "Servers:"
	ForEach($server in $servers)
	{
		line 1 "Name: " $server.ServerName
		line 2 "Product                  : " $server.CitrixProductName
		line 2 "Edition                  : " $server.CitrixEdition
		line 2 "Version                  : " $server.CitrixVersion
		line 2 "Service Pack             : " $server.CitrixServicePack
		line 2 "IP Address               : " $server.IPAddresses
		line 2 "Logons                   : " -NoNewLine
		If($server.LogOnsEnabled)
		{
			line 0 "Enabled"
		} 
		Else 
		{
			line 0 "Disabled"
		}
		line 2 "Logon Control Mode       : " -nonewline
		switch ($Server.LogOnMode)
		{
			"Unknown"                       {line 0 "Unknown"}
			"AllowLogOns"                   {line 0 "Allow logons and reconnections"}
			"ProhibitNewLogOnsUntilRestart" {line 0 "Prohibit logons until server restart"}
			"ProhibitNewLogOns "            {line 0 "Prohibit logons only"}
			"ProhibitLogOns "               {line 0 "Prohibit logons and reconnections"}
			Default {line 0 "Logon control mode could not be determined: $($Server.LogOnMode)"}
		}

		line 2 "Product Installation Date: " $server.CitrixInstallDate
		line 2 "Operating System Version : " $server.OSVersion -NoNewLine
		line 0 " " $server.OSServicePack
		line 2 "Zone                     : " $server.ZoneName
		line 2 "Election Preference      : " -nonewline
		switch ($server.ElectionPreference)
		{
			"Unknown"           {line 0 "Unknown"}
			"MostPreferred"     {line 0 "Most Preferred"}
			"Preferred"         {line 0 "Preferred"}
			"DefaultPreference" {line 0 "Default Preference"}
			"NotPreferred"      {line 0 "Not Preferred"}
			"WorkerMode"        {line 0 "Worker Mode"}
			Default {line 0 "Server election preference could not be determined: $($server.ElectionPreference)"}
		}
		line 2 "Folder                   : " $server.FolderPath
		line 2 "Product Installation Path: " $server.CitrixInstallPath
		If($server.LicenseServerName)
		{
			line 2 "License Server Name      : " $server.LicenseServerName
			line 2 "License Server Port      : " $server.LicenseServerPortNumber
		}
		If($server.ICAPortNumber -gt 0)
		{
			line 2 "ICA Port Number          : " $server.ICAPortNumber
		}
		
		#applications published to server
		$Applications = Get-XAApplication -ServerName $server.ServerName -EA 0 | sort-object FolderPath, DisplayName
		If( $? -and $Applications )
		{
			line 2 "Published applications:"
			ForEach($app in $Applications)
			{
				line 3 "Display name: " $app.DisplayName
				line 3 "Folder path : " $app.FolderPath
				line 0 ""
			}
		}
		#Citrix hotfixes installed
		$hotfixes = Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | sort-object HotfixName
		If( $? -and $hotfixes )
		{
			line 2 "Citrix Hotfixes:"
			ForEach($hotfix in $hotfixes)
			{
				line 3 "Hotfix           : " $hotfix.HotfixName
				line 3 "Installed by     : " $hotfix.InstalledBy
				line 3 "Installed date   : " $hotfix.InstalledOn
				line 3 "Hotfix type      : " $hotfix.HotfixType
				line 3 "Valid            : " $hotfix.Valid
				line 3 "Hotfixes replaced: "
				ForEach($Replaced in $hotfix.HotfixesReplaced)
				{
					line 4 $Replaced
				}
				line 0 ""
			}
		}
		line 0 "" 
		Write-Output $global:output
		$global:output = $null
	}
}
Else 
{
	line 0 "Server information could not be retrieved"
}
$servers = $null
$global:output = $null

#worker groups
$WorkerGroups = Get-XAWorkerGroup -EA 0 | sort-object WorkerGroupName

If( $? -and $WorkerGroups)
{
	line 0 ""
	line 0 "Worker Groups:"
	ForEach($WorkerGroup in $WorkerGroups)
	{
		line 0 ""
		line 1 "Name: " $WorkerGroup.WorkerGroupName
		line 2 "Description: " $WorkerGroup.Description
		line 2 "Folder Path: " $WorkerGroup.FolderPath
		If($WorkerGroup.ServerNames)
		{
			line 2 "Farm Servers:"
			$TempArray = $WorkerGroup.ServerNames | Sort-Object
			ForEach($ServerName in $TempArray)
			{
				line 3 $ServerName
			}
			$TempArray = $null
		}
		If($WorkerGroup.ServerGroups)
		{
			line 2 "Server Group Accounts:"
			$TempArray = $WorkerGroup.ServerGroups | Sort-Object
			ForEach($ServerGroup in $TempArray)
			{
				line 3 $ServerGroup
			}
			$TempArray = $null
		}
		If($WorkerGroup.OUs)
		{
			line 2 "Organizational Units:"
			$TempArray = $WorkerGroup.OUs | Sort-Object
			ForEach($OU in $TempArray)
			{
				line 3 $OU
			}
			$TempArray = $null
		}
		#applications published to worker group
		$Applications = Get-XAApplication -WorkerGroup $WorkerGroup.WorkerGroupName -EA 0 | sort-object FolderPath, DisplayName
		If( $? -and $Applications )
		{
			line 2 "Published applications:"
			ForEach($app in $Applications)
			{
				line 0 ""
				line 3 "Display name: " $app.DisplayName
				line 3 "Folder path : " $app.FolderPath
			}
		}

		Write-Output $global:output
		$global:output = $null
	}
}
Else 
{
	line 0 "Worker Group information could not be retrieved"
}
$WorkerGroups = $null
$global:output = $null

#zones
$Zones = Get-XAZone -EA 0 | sort-object ZoneName
If( $? )
{
	line 0 ""
	line 0 "Zones:"
	ForEach($Zone in $Zones)
	{
		line 1 "Zone Name: " $Zone.ZoneName
		line 2 "Current Data Collector: " $Zone.DataCollector
		$Servers = Get-XAServer -ZoneName $Zone.ZoneName -EA 0 | sort-object ElectionPreference, ServerName
		If( $? )
		{		
			line 2 "Servers in Zone"
	
			ForEach($Server in $Servers)
			{
				line 3 "Server Name and Preference: " $server.ServerName -NoNewLine
				line 0  " - " -nonewline
				switch ($server.ElectionPreference)
				{
					"Unknown"           {line 0 "Unknown"}
					"MostPreferred"     {line 0 "Most Preferred"}
					"Preferred"         {line 0 "Preferred"}
					"DefaultPreference" {line 0 "Default Preference"}
					"NotPreferred"      {line 0 "Not Preferred"}
					"WorkerMode"        {line 0 "Worker Mode"}
					Default {line 0 "Zone preference could not be determined: $($server.ElectionPreference)"}
				}
			}
		}
		Else
		{
			line 2 "Unable to enumerate servers in the zone"
		}
		Write-Output $global:output
		$global:output = $null
		$Servers = $Null
	}
}
Else 
{
	line 0 "Zone information could not be retrieved"
}
$Servers = $null
$Zones = $null
$global:output = $null

#make sure Citrix.GroupPolicy.Commands module is loaded
If(!(Check-LoadedModule "Citrix.GroupPolicy.Commands"))
{
	write-warning "The Citrix Group Policy module Citrix.GroupPolicy.Commands.psm1 does not exist (http://support.citrix.com/article/CTX128625), Citrix Policy documentation will not take place."
}
else
{
	Echo "Please wait while Citrix Policies are retrieved..."
	$Policies = Get-CtxGroupPolicy -EA 0 | sort-object Type,Priority
	If( $? )
	{
		line 0 ""
		line 0 "Policies:"
		ForEach($Policy in $Policies)
		{
			line 1 "Policy Name   : " $Policy.PolicyName
			line 2 "Type          : " $Policy.Type
			line 2 "Description   : " $Policy.Description
			line 2 "Enabled       : " $Policy.Enabled
			line 2 "Priority      : " $Policy.Priority

			$filters = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName -EA 0

			If( $? )
			{
				If(![String]::IsNullOrEmpty($filters))
				{
					line 2 "Filter(s):"
					ForEach($Filter in $Filters)
					{
						Line 3 "Filter name   : " $filter.FilterName
						Line 3 "Filter type   : " $filter.FilterType
						Line 3 "Filter enabled: " $filter.Enabled
						Line 3 "Filter mode   : " $filter.Mode
						Line 3 "Filter value  : " $filter.FilterValue
						Line 3 ""
					}
				}
				Else
				{
					line 2 "No filter information"
				}
			}
			Else
			{
				Line 2 "Unable to retrieve Filter settings"
			}

			$Settings = Get-CtxGroupPolicyConfiguration -PolicyName $Policy.PolicyName -EA 0
			If( $? )
			{
				ForEach($Setting in $Settings)
				{
					If($Setting.Type -eq "Computer")
					{
						line 2 "Computer settings:"
						If($Setting.IcaListenerTimeout.State -ne "NotConfigured")
						{
							line 3 "ICA\ICA listener connection timeout - Value (milliseconds): " $Setting.IcaListenerTimeout.Value
						}
						If($Setting.IcaListenerPortNumber.State -ne "NotConfigured")
						{
							line 3 "ICA\ICA listener port number - Value: " $Setting.IcaListenerPortNumber.Value
						}
						If($Setting.AutoClientReconnect.State -ne "NotConfigured")
						{
							line 3 "ICA\Auto Client Reconnect\Auto client reconnect: " $Setting.AutoClientReconnect.State
						}
						If($Setting.AutoClientReconnectLogging.State -ne "NotConfigured")
						{
							line 3 "ICA\Auto Client Reconnect\Auto client reconnect logging: "
							switch ($Setting.AutoClientReconnectLogging.Value)
							{
								"DoNotLogAutoReconnectEvents" {line 4 "Do Not Log auto-reconnect events"}
								"LogAutoReconnectEvents"      {line 4 "Log auto-reconnect events"}
								Default {line 4 "Auto client reconnect logging could not be determined: $($Setting.AutoClientReconnectLogging.Value)"}
							}
						}
						If($Setting.IcaRoundTripCalculation.State -ne "NotConfigured")
						{
							line 3 "ICA\End User Monitoring\ICA round trip calculation: " $Setting.IcaRoundTripCalculation.State
						}
						If($Setting.IcaRoundTripCalculationInterval.State -ne "NotConfigured")
						{
							line 3 "ICA\End User Monitoring\ICA round trip calculation interval - Value (seconds): " $Setting.IcaRoundTripCalculationInterval.Value
						}
						If($Setting.IcaRoundTripCalculationWhenIdle.State -ne "NotConfigured")
						{
							line 3 "ICA\End User Monitoring\ICA round trip calculations for idle connections: " $Setting.IcaRoundTripCalculationWhenIdle.State
						}
						If($Setting.DisplayMemoryLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Graphics\Display memory limit - Value (KB): " $Setting.DisplayMemoryLimit.Value
						}
						If($Setting.DisplayDegradePreference.State -ne "NotConfigured")
						{
							line 3 "ICA\Graphics\Display mode degrade preference: "
							
							switch ($Setting.DisplayDegradePreference.Value)
							{
								"ColorDepth" {line 4 "Degrade color depth first"}
								"Resolution" {line 4 "Degrade resolution first"}
								Default {line 4 "Display mode degrade preference could not be determined: $($Setting.DisplayDegradePreference.Value)"}
							}
						}
						If($Setting.DynamicPreview.State -ne "NotConfigured")
						{
							line 3 "ICA\Graphics\Dynamic Windows Preview: " $Setting.DynamicPreview.State
						}
						If($Setting.ImageCaching.State -ne "NotConfigured")
						{
							line 3 "ICA\Graphics\Image caching: " $Setting.ImageCaching.State
						}
						If($Setting.MaximumColorDepth.State -ne "NotConfigured")
						{
							line 3 "ICA\Graphics\Maximum allowed color depth: "
							switch ($Setting.MaximumColorDepth.Value)
							{
								"BitsPerPixel8"  {line 4 "8 Bits Per Pixel"}
								"BitsPerPixel15" {line 4 "15 Bits Per Pixel"}
								"BitsPerPixel16" {line 4 "16 Bits Per Pixel"}
								"BitsPerPixel24" {line 4 "24 Bits Per Pixel"}
								"BitsPerPixel32" {line 4 "32 Bits Per Pixel"}
								Default {line 4 "Maximum allowed color depth could not be determined: $($Setting.MaximumColorDepth.Value)"}
							}
						}
						If($Setting.DisplayDegradeUserNotification.State -ne "NotConfigured")
						{
							line 3 "ICA\Graphics\Notify user when display mode is degraded: " $Setting.DisplayDegradeUserNotification.State
						}
						If($Setting.QueueingAndTossing.State -ne "NotConfigured")
						{
							line 3 "ICA\Graphics\Queueing and tossing: " $Setting.QueueingAndTossing.State
						}
						If($Setting.PersistentCache.State -ne "NotConfigured")
						{
							line 3 "ICA\Graphics\Caching\Persistent Cache Threshold - Value (Kbps): " $Setting.PersistentCache.Value
						}
						If($Setting.IcaKeepAliveTimeout.State -ne "NotConfigured")
						{
							line 3 "ICA\Keep ALive\ICA keep alive timeout - Value (seconds): " $Setting.IcaKeepAliveTimeout.Value
						}
						If($Setting.IcaKeepAlives.State -ne "NotConfigured")
						{
							line 3 "ICA\Keep ALive\ICA keep alives - Value: "
							switch ($Setting.IcaKeepAlives.Value)
							{
								"DoNotSendKeepAlives" {line 4 "Do not send ICA keep alive messages"}
								"SendKeepAlives"      {line 4 "Send ICA keep alive messages"}
								Default {line 4 "ICA keep alives could not be determined: $($Setting.IcaKeepAlives.Value)"}
							}
						}
						If($Setting.MultimediaConferencing.State -ne "NotConfigured")
						{
							line 3 "ICA\Multimedia\Multimedia conferencing: " $Setting.MultimediaConferencing.State
						}
						If($Setting.MultimediaAcceleration.State -ne "NotConfigured")
						{
							line 3 "ICA\Multimedia\Windows Media Redirection: " $Setting.MultimediaAcceleration.State
						}
						If($Setting.MultimediaAccelerationDefaultBufferSize.State -ne "NotConfigured")
						{
							line 3 "ICA\Multimedia\Windows Media Redirection Buffer Size - Value (seconds): " $Setting.MultimediaAccelerationDefaultBufferSize.Value
						}
						If($Setting.MultimediaAccelerationUseDefaultBufferSize.State -ne "NotConfigured")
						{
							line 3 "ICA\Multimedia\Windows Media Redirection Buffer Size Use: " $Setting.MultimediaAccelerationUseDefaultBufferSize.State
						}
						If($Setting.MultiPortPolicy.State -ne "NotConfigured")
						{
							line 3 "ICA\MultiStream Connections\Multi-Port Policy: " 
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
							line 4 "CGP port1: " $cgpport1 -nonewline 
							line 1 "priority: " $cgpport1priority[0]
							line 4 "CGP port2: " $cgpport2 -nonewline
							line 1 "priority: " $cgpport2priority[0]
							line 4 "CGP port3: " $cgpport3 -nonewline
							line 1 "priority: " $cgpport3priority[0]
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
							line 3 "ICA\MultiStream Connections\Multi-Stream: " $Setting.MultiStreamPolicy.State
						}
						If($Setting.PromptForPassword.State -ne "NotConfigured")
						{
							line 3 "ICA\Security\Prompt for password: " $Setting.PromptForPassword.State
						}
						If($Setting.IdleTimerInterval.State -ne "NotConfigured")
						{
							line 3 "ICA\Server Limits\Server idle timer interval - Value (milliseconds): " $Setting.IdleTimerInterval.Value
						}
						If($Setting.SessionReliabilityConnections.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Reliability\Session reliability connections: " $Setting.SessionReliabilityConnections.State
						}
						If($Setting.SessionReliabilityPort.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Reliability\Session reliability port number - Value: " $Setting.SessionReliabilityPort.Value
						}
						If($Setting.SessionReliabilityTimeout.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Reliability\Session reliability timeout - Value (seconds): " $Setting.SessionReliabilityTimeout.Value
						}
						If($Setting.Shadowing.State -ne "NotConfigured")
						{
							line 3 "ICA\Shadowing\Shadowing: " $Setting.Shadowing.State
						}
						If($Setting.LicenseServerHostName.State -ne "NotConfigured")
						{
							line 3 "Licensing\License server host name - Value: " $Setting.LicenseServerHostName.Value
						}
						If($Setting.LicenseServerPort.State -ne "NotConfigured")
						{
							line 3 "Licensing\License server port - Value: " $Setting.LicenseServerPort.Value
						}
						If($Setting.FarmName.State -ne "NotConfigured")
						{
							line 3 "Power and Capacity Management\Farm name - Value: " $Setting.FarmName.Value
						}
						If($Setting.WorkloadName.State -ne "NotConfigured")
						{
							line 3 "Power and Capacity Management\Workload name - Value: " $Setting.WorkloadName.Value
						}
						If($Setting.ConnectionAccessControl.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Connection access control - Value: "
							switch ($Setting.ConnectionAccessControl.Value)
							{
								"AllowAny"                     {line 4 "Any connections"}
								"AllowTicketedConnectionsOnly" {line 4 "Citrix Access Gateway, Citrix Receiver, and Web Interface connections only"}
								"AllowAccessGatewayOnly"       {line 4 "Citrix Access Gateway connections only"}
								Default {line 4 "Connection access control could not be determined: $($Setting.ConnectionAccessControl.Value)"}
							}
						}
						If($Setting.DnsAddressResolution.State -ne "NotConfigured")
						{
							line 3 "Server Settings\DNS address resolution: " $Setting.DnsAddressResolution.State
						}
						If($Setting.FullIconCaching.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Full icon caching: " $Setting.FullIconCaching.State
						}
						If($Setting.LoadEvaluator.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Load Evaluator Name - Load evaluator: " $Setting.LoadEvaluator.Value
						}
						If($Setting.ProductEdition.State -ne "NotConfigured")
						{
							line 3 "Server Settings\XenApp product edition - Value: " $Setting.ProductEdition.Value
						}
						If($Setting.ProductModel.State -ne "NotConfigured")
						{
							line 3 "Server Settings\XenApp product model - Value: " -nonewline
							switch ($Setting.ProductModel.Value)
							{
								"XenAppCCU"                  {line 0 "XenApp"}
								"XenDesktopConcurrentServer" {line 0 "XenDesktop Concurrent"}
								"XenDesktopUserDevice"       {line 0 "XenDesktop User Device"}
								Default {line 0 "XenApp product model could not be determined: $($Setting.ProductModel.Value)"}
							}
						}
						If($Setting.UserSessionLimit.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Connection Limits\Limit user sessions - Value: " $Setting.UserSessionLimit.Value
						}
						If($Setting.UserSessionLimitAffectsAdministrators.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Connection Limits\Limits on administrator sessions: " $Setting.UserSessionLimitAffectsAdministrators.State
						}
						If($Setting.UserSessionLimitLogging.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Connection Limits\Logging of logon limit events: " $Setting.UserSessionLimitLogging.State
						}
						If($Setting.HealthMonitoring.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Health Monitoring and Recovery\Health monitoring: " $Setting.HealthMonitoring.State
						}
						If($Setting.HealthMonitoringTests.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Health Monitoring and Recovery\Health monitoring tests: " 
							[xml]$XML = $Setting.HealthMonitoringTests.Value
							ForEach($Test in $xml.hmrtests.tests.test)
							{
								line 4 "Name           : " $test.name
								line 4 "File Location  : " $test.file
								If($test.arguments)
								{
									line 4 "Arguments      : " $test.arguments
								}
								line 4 "Description    : " $test.description
								line 4 "Interval       : " $test.interval
								line 4 "Time-out       : " $test.timeout
								line 4 "Threshold      : " $test.threshold
								line 4 "Recovery action: " $test.recoveryAction
								line 0 ""
							}
						}
						If($Setting.MaximumServersOfflinePercent.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Health Monitoring and Recovery\Max % of servers with logon control - Value: " $Setting.MaximumServersOfflinePercent.Value
						}
						If($Setting.CpuManagementServerLevel.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Memory/CPU\CPU management server level - Value: "
							switch ($Setting.CpuManagementServerLevel.Value)
							{
								"NoManagement" {line 4 "No CPU utilization management"}
								"Fair"         {line 4 "Fair sharing of CPU between sessions"}
								"Preferential" {line 4 "Preferential Load Balancing"}
								Default {line 4 "CPU management server level could not be determined: $($Setting.CpuManagementServerLevel.Value)"}
							}
						}
						If($Setting.MemoryOptimization.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Memory/CPU\Memory optimization: " $Setting.MemoryOptimization.State
						}
						If($Setting.MemoryOptimizationExcludedPrograms.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Memory/CPU\Memory optimization application exclusion list - Values: "
							$array = $Setting.MemoryOptimizationExcludedPrograms.Values
							foreach( $element in $array)
							{
								line 4 $element
							}
						}
						If($Setting.MemoryOptimizationIntervalType.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Memory/CPU\Memory optimization interval - Value: " -nonewline
							switch ($Setting.MemoryOptimizationIntervalType.Value)
							{
								"AtStartup" {line 0 "Only at startup time"}
								"Daily"     {line 0 "Daily"}
								"Weekly"    {line 0 "Weekly"}
								"Monthly"   {line 0 "Monthly"}
								Default {line 0 " could not be determined: $($Setting.MemoryOptimizationIntervalType.Value)"}
							}
						}
						If($Setting.MemoryOptimizationDayOfMonth.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Memory/CPU\Memory optimization schedule: day of month - Value: " $Setting.MemoryOptimizationDayOfMonth.Value
						}
						If($Setting.MemoryOptimizationDayOfWeek.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Memory/CPU\Memory optimization schedule: day of week - Value: " $Setting.MemoryOptimizationDayOfWeek.Value
						}
						If($Setting.MemoryOptimizationTime.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Memory/CPU\Memory optimization schedule: Time (H:MM TT): " -nonewline
							ConvertNumberToTime $Setting.MemoryOptimizationTime.Value
						}
						If($Setting.OfflineClientTrust.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Offline Applications\Offline app client trust: " $Setting.OfflineClientTrust.State
						}
						If($Setting.OfflineEventLogging.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Offline Applications\Offline app event logging: " $Setting.OfflineEventLogging.State
						}
						If($Setting.OfflineLicensePeriod.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Offline Applications\Offline app license period - Days: " $Setting.OfflineLicensePeriod.Value
						}
						If($Setting.OfflineUsers.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Offline Applications\Offline app users: " 
							$array = $Setting.OfflineUsers.Values
							foreach( $element in $array)
							{
								line 4 $element
							}
						}
						If($Setting.RebootCustomMessage.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot custom warning: " $Setting.RebootCustomMessage.State
						}
						If($Setting.RebootCustomMessageText.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot custom warning text - Value: " 
							line 4 $Setting.RebootCustomMessageText.Value
						}
						If($Setting.RebootDisableLogOnTime.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot logon disable time - Value: "
							switch ($Setting.RebootDisableLogOnTime.Value)
							{
								"DoNotDisableLogOnsBeforeReboot" {line 4 "Do not disable logons before reboot"}
								"Disable5MinutesBeforeReboot"    {line 4 "Disable 5 minutes before reboot"}
								"Disable10MinutesBeforeReboot"   {line 4 "Disable 10 minutes before reboot"}
								"Disable15MinutesBeforeReboot"   {line 4 "Disable 15 minutes before reboot"}
								"Disable30MinutesBeforeReboot"   {line 4 "Disable 30 minutes before reboot"}
								"Disable60MinutesBeforeReboot"   {line 4 "Disable 60 minutes before reboot"}
								Default {line 4 "Reboot logon disable time could not be determined: $($Setting.RebootDisableLogOnTime.Value)"}
							}
						}
						If($Setting.RebootScheduleFrequency.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot schedule frequency - Days: " $Setting.RebootScheduleFrequency.Value
						}
						If($Setting.RebootScheduleRandomizationInterval.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot schedule randomization interval - Minutes: " $Setting.RebootScheduleRandomizationInterval.Value
						}
						If($Setting.RebootScheduleStartDate.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot schedule start date - Date (MM/DD/YYYY): " -nonewline
							$Tmp = ConvertIntegerToDate $Setting.RebootScheduleStartDate.Value
							line 0 $Tmp
						}
						If($Setting.RebootScheduleTime.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot schedule time - Time (H:MM TT): " -nonewline
							ConvertNumberToTime $Setting.RebootScheduleTime.Value 						
						}
						If($Setting.RebootWarningInterval.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot warning interval - Value: "
							switch ($Setting.RebootWarningInterval.Value)
							{
								"Every1Minute"   {line 4 "Every 1 Minute"}
								"Every3Minutes"  {line 4 "Every 3 Minutes"}
								"Every5Minutes"  {line 4 "Every 5 Minutes"}
								"Every10Minutes" {line 4 "Every 10 Minutes"}
								"Every15Minutes" {line 4 "Every 15 Minutes"}
								Default {line 4 "Reboot warning interval could not be determined: $($Setting.RebootWarningInterval.Value)"}
							}
						}
						If($Setting.RebootWarningStartTime.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot warning start time - Value: "
							switch ($Setting.RebootWarningStartTime.Value)
							{
								"Start5MinutesBeforeReboot"  {line 4 "Start 5 Minutes Before Reboot"}
								"Start10MinutesBeforeReboot" {line 4 "Start 10 Minutes Before Reboot"}
								"Start15MinutesBeforeReboot" {line 4 "Start 15 Minutes Before Reboot"}
								"Start30MinutesBeforeReboot" {line 4 "Start 30 Minutes Before Reboot"}
								"Start60MinutesBeforeReboot" {line 4 "Start 60 Minutes Before Reboot"}
								Default {line 4 "Reboot warning start time could not be determined: $($Setting.RebootWarningStartTime.Value)"}
							}
						}
						If($Setting.RebootWarningMessage.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Reboot warning to users: " $Setting.RebootWarningMessage.State
						}
						If($Setting.ScheduledReboots.State -ne "NotConfigured")
						{
							line 3 "Server Settings\Reboot Behavior\Scheduled reboots: " $Setting.ScheduledReboots.State
						}
						If($Setting.FilterAdapterAddresses.State -ne "NotConfigured")
						{
							line 3 "Virtual IP\Virtual IP adapter address filtering: " $Setting.FilterAdapterAddresses.State
						}
						If($Setting.EnhancedCompatibilityPrograms.State -ne "NotConfigured")
						{
							line 3 "Virtual IP\Virtual IP compatibility programs list - Values: " 
							$array = $Setting.EnhancedCompatibilityPrograms.Values
							foreach( $element in $array)
							{
								line 4 $element
							}
						}
						If($Setting.EnhancedCompatibility.State -ne "NotConfigured")
						{
							line 3 "Virtual IP\Virtual IP enhanced compatibility: " $Setting.EnhancedCompatibility.State
						}
						If($Setting.FilterAdapterAddressesPrograms.State -ne "NotConfigured")
						{
							line 3 "Virtual IP\Virtual IP filter adapter addresses programs list - Values: " 
							$array = $Setting.FilterAdapterAddressesPrograms.Values
							foreach( $element in $array)
							{
								line 4 $element
							}
						}
						If($Setting.VirtualLoopbackSupport.State -ne "NotConfigured")
						{
							line 3 "Virtual IP\Virtual IP loopback support: " $Setting.VirtualLoopbackSupport.State
						}
						If($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured")
						{
							line 3 "Virtual IP\Virtual IP virtual loopback programs list - Values: " 
							$array = $Setting.VirtualLoopbackPrograms.Values
							foreach( $element in $array)
							{
								line 4 $element
							}
						}
						If($Setting.TrustXmlRequests.State -ne "NotConfigured")
						{
							line 3 "XML Service\Trust XML requests: " $Setting.TrustXmlRequests.State
						}
						If($Setting.XmlServicePort.State -ne "NotConfigured")
						{
							line 3 "XML Service\XML service port - Value: " $Setting.XmlServicePort.Value
						}
					}
					Else
					{
						line 2 "User settings:"
						If($Setting.ClipboardRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\Client clipboard redirection: " $Setting.ClipboardRedirection.State
						}
						If($Setting.DesktopLaunchForNonAdmins.State -ne "NotConfigured")
						{
							line 3 "ICA\Desktop launches: " $Setting.DesktopLaunchForNonAdmins.State
						}
						If($Setting.NonPublishedProgramLaunching.State -ne "NotConfigured")
						{
							line 3 "ICA\Launching of non-published programs during client connection: " $Setting.NonPublishedProgramLaunching.State
						}
						If($Setting.FlashAcceleration.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash acceleration: " $Setting.FlashAcceleration.State
						}
						If($Setting.FlashUrlColorList.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash background color list - Values: "
							$Values = $Setting.FlashUrlColorList.Values
							ForEach($Value in $Values)
							{
								line 4 $Value
							}
							$Values = $null
						}
						If($Setting.FlashBackwardsCompatibility.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash backwards compatibility: " $Setting.FlashBackwardsCompatibility.State
						}
						If($Setting.FlashDefaultBehavior.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash default behavior - Value: "
							switch ($Setting.FlashDefaultBehavior.Value)
							{
								"Block"   {line 4 "Block Flash player"}
								"Disable" {line 4 "Disable Flash acceleration"}
								"Enable"  {line 4 "Enable Flash acceleration"}
								Default {line 4 "Flash default behavior could not be determined: $($Setting.FlashDefaultBehavior.Value)"}
							}
						}
						If($Setting.FlashEventLogging.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash event logging: " $Setting.FlashEventLogging.State
						}
						If($Setting.FlashIntelligentFallback.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash intelligent fallback: " $Setting.FlashIntelligentFallback.State
						}
						If($Setting.FlashLatencyThreshold.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash latency threshold - Value (milliseconds): " $Setting.FlashLatencyThreshold.Value
						}
						If($Setting.FlashServerSideContentFetchingWhitelist.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash server-side content fetching URL list - Values: "
							$Values = $Setting.FlashServerSideContentFetchingWhitelist.Values
							ForEach($Value in $Values)
							{
								line 4 $Value
							}
							$Values = $null
						}
						If($Setting.FlashUrlCompatibilityList.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash URL compatibility list: " 
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
								line 4 "Action: " $Action -NoNewLine
								line 1 "URL: "$Url
							}
							$Values = $null
							$Action = $null
							$Url = $null
						}
						If($Setting.AllowSpeedFlash.State -ne "NotConfigured")
						{
							line 3 "ICA\Adobe Flash Delivery\Legacy Server Side Optimizations\Flash quality adjustment - Value: "
							switch ($Setting.AllowSpeedFlash.Value)
							{
								"NoOptimization"      {line 4 "Do not optimize Adobe Flash animation options"}
								"AllConnections"      {line 4 "Optimize Adobe Flash animation options for all connections"}
								"RestrictedBandwidth" {line 4 "Optimize Adobe Flash animation options for low bandwidth connections only"}
								Default {line 4 "Flash quality adjustment could not be determined: $($Setting.AllowSpeedFlash.Value)"}
							}
						}
						If($Setting.AudioPlugNPlay.State -ne "NotConfigured")
						{
							line 3 "ICA\Audio\Audio Plug N Play: " $Setting.AudioPlugNPlay.State
						}
						If($Setting.AudioQuality.State -ne "NotConfigured")
						{
							line 3 "ICA\Audio\Audio quality - Value: "
							switch ($Setting.AudioQuality.Value)
							{
								"Low"    {line 4 "Low - for low-speed connections"}
								"Medium" {line 4 "Medium - optimized for speech"}
								"High"   {line 4 "High - high definition audio"}
								Default {line 4 "Audio quality could not be determined: $($Setting.AudioQuality.Value)"}
							}
						}
						If($Setting.ClientAudioRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\Audio\Client audio redirection: " $Setting.ClientAudioRedirection.State
						}
						If($Setting.MicrophoneRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\Audio\Client microphone redirection: " $Setting.MicrophoneRedirection.State
						}
						If($Setting.AudioBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\Audio redirection bandwidth limit - Value (Kbps): " $Setting.AudioBandwidthLimit.Value
						}
						If($Setting.AudioBandwidthPercent.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\Audio redirection bandwidth limit percent - Value: " $Setting.AudioBandwidthPercent.Value
						}
						If($Setting.USBBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\Client USB device redirection bandwidth limit - Value: " $Setting.USBBandwidthLimit.Value
						}
						If($Setting.USBBandwidthPercent.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\Client USB device redirection bandwidth limit percent - Value: " $Setting.USBBandwidthPercent.Value
						}
						If($Setting.ClipboardBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\Clipboard redirection bandwidth limit - Value (Kbps): " $Setting.ClipboardBandwidthLimit.Value
						}
						If($Setting.ClipboardBandwidthPercent.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\Clipboard redirection bandwidth limit percent - Value: " $Setting.ClipboardBandwidthPercent.Value
						}
						If($Setting.ComPortBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\COM port redirection bandwidth limit - Value (Kbps): " $Setting.ComPortBandwidthLimit.Value
						}
						If($Setting.ComPortBandwidthPercent.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\COM port redirection bandwidth limit percent - Value: " $Setting.ComPortBandwidthPercent.Value
						}
						If($Setting.FileRedirectionBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\File redirection bandwidth limit - Value (Kbps): " $Setting.FileRedirectionBandwidthLimit.Value
						}
						If($Setting.FileRedirectionBandwidthPercent.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\File redirection bandwidth limit percent - Value: " $Setting.FileRedirectionBandwidthPercent.Value
						}
						If($Setting.HDXMultimediaBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration bandwidth limit - Value: " $Setting.HDXMultimediaBandwidthLimit.Value
						}
						If($Setting.HDXMultimediaBandwidthPercent.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\HDX MediaStream Multimedia Acceleration bandwidth limit percent - Value: " $Setting.HDXMultimediaBandwidthPercent.Value
						}
						If($Setting.LptBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\LPT port redirection bandwidth limit - Value (Kbps): " $Setting.LptBandwidthLimit.Value
						}
						If($Setting.LptBandwidthLimitPercent.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\LPT port redirection bandwidth limit percent - Value: " $Setting.LptBandwidthLimitPercent.Value
						}
						If($Setting.OverallBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\Overall session bandwidth limit - Value (Kbps): " $Setting.OverallBandwidthLimit.Value
						}
						If($Setting.PrinterBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\Printer redirection bandwidth limit - Value (Kbps): " $Setting.PrinterBandwidthLimit.Value
						}
						If($Setting.PrinterBandwidthPercent.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\Printer redirection bandwidth limit percent - Value: " $Setting.PrinterBandwidthPercent.Value
						}
						If($Setting.TwainBandwidthLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\TWAIN device redirection bandwidth limit - Value (Kbps): " $Setting.TwainBandwidthLimit.Value
						}
						If($Setting.TwainBandwidthPercent.State -ne "NotConfigured")
						{
							line 3 "ICA\Bandwidth\TWAIN device redirection bandwidth limit percent - Value: " $Setting.TwainBandwidthPercent.Value
						}
						If($Setting.DesktopWallpaper.State -ne "NotConfigured")
						{
							line 3 "ICA\Desktop UI\Desktop wallpaper: " $Setting.DesktopWallpaper.State
						}
						If($Setting.MenuAnimation.State -ne "NotConfigured")
						{
							line 3 "ICA\Desktop UI\Menu animation: " $Setting.MenuAnimation.State
						}
						If($Setting.WindowContentsVisibleWhileDragging.State -ne "NotConfigured")
						{
							line 3 "ICA\Desktop UI\View window contents while dragging: " $Setting.WindowContentsVisibleWhileDragging.State
						}
						If($Setting.AutoConnectDrives.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Auto connect client drives: " $Setting.AutoConnectDrives.State
						}
						If($Setting.ClientDriveRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Client drive redirection: " $Setting.ClientDriveRedirection.State
						}
						If($Setting.ClientFixedDrives.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Client fixed drives: " $Setting.ClientFixedDrives.State
						}
						If($Setting.ClientFloppyDrives.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Client floppy drives: " $Setting.ClientFloppyDrives.State
						}
						If($Setting.ClientNetworkDrives.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Client network drives: " $Setting.ClientNetworkDrives.State
						}
						If($Setting.ClientOpticalDrives.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Client optical drives: " $Setting.ClientOpticalDrives.State
						}
						If($Setting.ClientRemoveableDrives.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Client removable drives: " $Setting.ClientRemoveableDrives.State
						}
						If($Setting.HostToClientRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Host to client redirection: " $Setting.HostToClientRedirection.State
						}
						If($Setting.ReadOnlyMappedDrive.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Read-only client drive access: " $Setting.ReadOnlyMappedDrive.State
						}
						If($Setting.SpecialFolderRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Special folder redirection: " $Setting.SpecialFolderRedirection.State
						}
						If($Setting.AsynchronousWrites.State -ne "NotConfigured")
						{
							line 3 "ICA\File Redirection\Use asynchronous writes: " $Setting.AsynchronousWrites.State
						}
						If($Setting.MultiStream.State -ne "NotConfigured")
						{
							line 3 "ICA\Multi-Stream Connections\Multi-Stream: " $Setting.MultiStream.State
						}
						If($Setting.ClientComPortsAutoConnection.State -ne "NotConfigured")
						{
							line 3 "ICA\Port Redirection\Auto connect client COM ports: " $Setting.ClientComPortsAutoConnection.State
						}
						If($Setting.ClientLptPortsAutoConnection.State -ne "NotConfigured")
						{
							line 3 "ICA\Port Redirection\Auto connect client LPT ports: " $Setting.ClientLptPortsAutoConnection.State
						}
						If($Setting.ClientComPortRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\Port Redirection\Client COM port redirection: " $Setting.ClientComPortRedirection.State
						}
						If($Setting.ClientLptPortRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\Port Redirection\Client LPT port redirection: " $Setting.ClientLptPortRedirection.State
						}
						If($Setting.ClientPrinterRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Client printer redirection: " $Setting.ClientPrinterRedirection.State
						}
						If($Setting.DefaultClientPrinter.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Default printer - Choose client's default printer: " 
							switch ($Setting.DefaultClientPrinter.Value)
							{
								"ClientDefault" {line 4 "Set default printer to the client's main printer"}
								"DoNotAdjust"   {line 4 "Do not adjust the user's default printer"}
								Default {line 0 "Default printer could not be determined: $($Setting.DefaultClientPrinter.Value)"}
							}
						}
						If($Setting.AutoCreationEventLogPreference.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Printer auto-creation event log preference - Value: " 
							switch ($Setting.AutoCreationEventLogPreference.Value)
							{
								"LogErrorsOnly"        {line 4 "Log errors only"}
								"LogErrorsAndWarnings" {line 4 "Log errors and warnings"}
								"DoNotLog"             {line 4 "Do not log errors or warnings"}
								Default {line 4 "Printer auto-creation event log preference could not be determined: $($Setting.AutoCreationEventLogPreference.Value)"}
							}
						}
						If($Setting.SessionPrinters.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Session printers:" 
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
											line 4 "Server       : $server"
											line 4 "Shared Name  : $share"
										}
									}
									Else
									{
										$tmp = $element.SubString( 0, 4 )
										Switch ($tmp)
										{
											"copi" 
											{
												$txt="Count        :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													line 4 "$txt $tmp2"
												}
											}
											"coll"
											{
												$txt="Collate      :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													line 4 "$txt $tmp2"
												}
											}
											"scal"
											{
												$txt="Scale (%)    :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													line 4 "$txt $tmp2"
												}
											}
											"colo"
											{
												$txt="Color        :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													line 4 "$txt " -nonewline
													Switch ($tmp2)
													{
														1 {line 0 "Monochrome"}
														2 {line 0 "Color"}
														Default {line 4 "Color could not be determined: $($element)"}
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
													line 4 "$txt " -nonewline
													Switch ($tmp2)
													{
														-1 {line 0 "150 dpi"}
														-2 {line 0 "300 dpi"}
														-3 {line 0 "600 dpi"}
														-4 {line 0 "1200 dpi"}
														Default 
														{
															line 0 "Custom..."
															line 4 "X resolution : " $tmp2
														}
													}
												}
											}
											"yres"
											{
												$txt="Y resolution :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													line 4 "$txt $tmp2"
												}
											}
											"orie"
											{
												$txt="Orientation  :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													line 4 "$txt " -nonewline
													switch ($tmp2)
													{
														"portrait"  {line 0 "Portrait"}
														"landscape" {line 0 "Landscape"}
														Default {line 4 "Orientation could not be determined: $($Element)"}
													}
												}
											}
											"dupl"
											{
												$txt="Duplex       :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													line 4 "$txt " -nonewline
													switch ($tmp2)
													{
														1 {line 0 "Simplex"}
														2 {line 0 "Vertical"}
														3 {line 0 "Horizontal"}
														Default {line 4 "Duplex could not be determined: $($Element)"}
													}
												}
											}
											"pape"
											{
												$txt="Paper Size   :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													line 4 "$txt " -nonewline
													switch ($tmp2)
													{
														1   {line 0  "Letter"}
														2   {line 0  "Letter Small"}
														3   {line 0  "Tabloid"}
														4   {line 0  "Ledger"}
														5   {line 0  "Legal"}
														6   {line 0  "Statement"}
														7   {line 0  "Executive"}
														8   {line 0  "A3"}
														9   {line 0  "A4"}
														10  {line 0  "A4 Small"}
														11  {line 0  "A5"}
														12  {line 0  "B4 (JIS)"}
														13  {line 0  "B5 (JIS)"}
														14  {line 0  "Folio"}
														15  {line 0  "Quarto"}
														16  {line 0  "10X14"}
														17  {line 0  "11X17"}
														18  {line 0  "Note"}
														19  {line 0  "Envelope #9"}
														20  {line 0  "Envelope #10"}
														21  {line 0  "Envelope #11"}
														22  {line 0  "Envelope #12"}
														23  {line 0  "Envelope #14"}
														24  {line 0  "C Size Sheet"}
														25  {line 0  "D Size Sheet"}
														26  {line 0  "E Size Sheet"}
														27  {line 0  "Envelope DL"}
														28  {line 0  "Envelope C5"}
														29  {line 0  "Envelope C3"}
														30  {line 0  "Envelope C4"}
														31  {line 0  "Envelope C6"}
														32  {line 0  "Envelope C65"}
														33  {line 0  "Envelope B4"}
														34  {line 0  "Envelope B5"}
														35  {line 0  "Envelope B6"}
														36  {line 0  "Envelope Italy"}
														37  {line 0  "Envelope Monarch"}
														38  {line 0  "Envelope Personal"}
														39  {line 0  "US Std Fanfold"}
														40  {line 0  "German Std Fanfold"}
														41  {line 0  "German Legal Fanfold"}
														42  {line 0  "B4 (ISO)"}
														43  {line 0  "Japanese Postcard"}
														44  {line 0  "9X11"}
														45  {line 0  "10X11"}
														46  {line 0  "15X11"}
														47  {line 0  "Envelope Invite"}
														48  {line 0  "Reserved - DO NOT USE"}
														49  {line 0  "Reserved - DO NOT USE"}
														50  {line 0  "Letter Extra"}
														51  {line 0  "Legal Extra"}
														52  {line 0  "Tabloid Extra"}
														53  {line 0  "A4 Extra"}
														54  {line 0  "Letter Transverse"}
														55  {line 0  "A4 Transverse"}
														56  {line 0  "Letter Extra Transverse"}
														57  {line 0  "A Plus"}
														58  {line 0  "B Plus"}
														59  {line 0  "Letter Plus"}
														60  {line 0  "A4 Plus"}
														61  {line 0  "A5 Transverse"}
														62  {line 0  "B5 (JIS) Transverse"}
														63  {line 0  "A3 Extra"}
														64  {line 0  "A5 Extra"}
														65  {line 0  "B5 (ISO) Extra"}
														66  {line 0  "A2"}
														67  {line 0  "A3 Transverse"}
														68  {line 0  "A3 Extra Transverse"}
														69  {line 0  "Japanese Double Postcard"}
														70  {line 0  "A6"}
														71  {line 0  "Japanese Envelope Kaku #2"}
														72  {line 0  "Japanese Envelope Kaku #3"}
														73  {line 0  "Japanese Envelope Chou #3"}
														74  {line 0  "Japanese Envelope Chou #4"}
														75  {line 0  "Letter Rotated"}
														76  {line 0  "A3 Rotated"}
														77  {line 0  "A4 Rotated"}
														78  {line 0  "A5 Rotated"}
														79  {line 0  "B4 (JIS) Rotated"}
														80  {line 0  "B5 (JIS) Rotated"}
														81  {line 0  "Japanese Postcard Rotated"}
														82  {line 0  "Double Japanese Postcard Rotated"}
														83  {line 0  "A6 Rotated"}
														84  {line 0  "Japanese Envelope Kaku #2 Rotated"}
														85  {line 0  "Japanese Envelope Kaku #3 Rotated"}
														86  {line 0  "Japanese Envelope Chou #3 Rotated"}
														87  {line 0  "Japanese Envelope Chou #4 Rotated"}
														88  {line 0  "B6 (JIS)"}
														89  {line 0  "B6 (JIS) Rotated"}
														90  {line 0  "12X11"}
														91  {line 0  "Japanese Envelope You #4"}
														92  {line 0  "Japanese Envelope You #4 Rotated"}
														93  {line 0  "PRC 16K"}
														94  {line 0  "PRC 32K"}
														95  {line 0  "PRC 32K(Big)"}
														96  {line 0  "PRC Envelope #1"}
														97  {line 0  "PRC Envelope #2"}
														98  {line 0  "PRC Envelope #3"}
														99  {line 0  "PRC Envelope #4"}
														100 {line 0 "PRC Envelope #5"}
														101 {line 0 "PRC Envelope #6"}
														102 {line 0 "PRC Envelope #7"}
														103 {line 0 "PRC Envelope #8"}
														104 {line 0 "PRC Envelope #9"}
														105 {line 0 "PRC Envelope #10"}
														106 {line 0 "PRC 16K Rotated"}
														107 {line 0 "PRC 32K Rotated"}
														108 {line 0 "PRC 32K(Big) Rotated"}
														109 {line 0 "PRC Envelope #1 Rotated"}
														110 {line 0 "PRC Envelope #2 Rotated"}
														111 {line 0 "PRC Envelope #3 Rotated"}
														112 {line 0 "PRC Envelope #4 Rotated"}
														113 {line 0 "PRC Envelope #5 Rotated"}
														114 {line 0 "PRC Envelope #6 Rotated"}
														115 {line 0 "PRC Envelope #7 Rotated"}
														116 {line 0 "PRC Envelope #8 Rotated"}
														117 {line 0 "PRC Envelope #9 Rotated"}
														Default {line 4 "Paper Size could not be determined: $($element)"}
													}
												}
											}
											"form"
											{
												$txt="Form Name    :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													If($tmp2.length -gt 0)
													{
														line 4 "$txt $tmp2"
													}
												}
											}
											"true"
											{
												$txt="TrueType     :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													line 4 "$txt " -nonewline
													switch ($tmp2)
													{
														1 {line 0 "Bitmap"}
														2 {line 0 "Download"}
														3 {line 0 "Substitute"}
														4 {line 0 "Outline"}
														Default {line 4 "TrueType could not be determined: $($Element)"}
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
													line 4 "$txt $tmp2"
												}
											}
											"loca" 
											{
												$txt="Location     :"
												$index = $element.SubString( 0 ).IndexOf( '=' )
												if( $index -ge 0 )
												{
													$tmp2 = $element.SubString( $index + 1 )
													If($tmp2.length -gt 0)
													{
														line 4 "$txt $tmp2"
													}
												}
											}
											Default {line 4 "Session printer setting could not be determined: $($Element)"}
										}
									}
								}
								line 0 ""
							}
						}
						If($Setting.WaitForPrintersToBeCreated.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Wait for printers to be created (desktop): " $Setting.WaitForPrintersToBeCreated.Values
						}
						If($Setting.ClientPrinterAutoCreation.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Client Printers\Auto-create client printers: "
							switch ($Setting.ClientPrinterAutoCreation.Value)
							{
								"DoNotAutoCreate"    {line 4 "Do not auto-create client printers"}
								"DefaultPrinterOnly" {line 4 "Auto-create the client's default printer only"}
								"LocalPrintersOnly"  {line 4 "Auto-create local (non-network) client printers only"}
								"AllPrinters"        {line 4 "Auto-create all client printers"}
								Default {line 4 "Auto-create client printers could not be determined: $($Setting.ClientPrinterAutoCreation.Value)"}
							}
						}
						If($Setting.GenericUniversalPrinterAutoCreation.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Client Printers\Auto-create generic universal printer: " $Setting.GenericUniversalPrinterAutoCreation.Value
						}
						If($Setting.ClientPrinterNames.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Client Printers\Client printer names - Value: " 
							switch ($Setting.ClientPrinterNames.Value)
							{
								"StandardPrinterNames" {line 4 "Standard printer names"}
								"LegacyPrinterNames"   {line 4 "Legacy printer names"}
								Default {line 4 "Client printer names could not be determined: $($Setting.ClientPrinterNames.Value)"}
							}
						}
						If($Setting.DirectConnectionsToPrintServers.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Client Printers\Direct connections to print servers: " $Setting.DirectConnectionsToPrintServers.State
						}
						If($Setting.PrinterDriverMappings.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Client Printers\Printer driver mapping and compatibility - Value: " 
							$array = $Setting.PrinterDriverMappings.Values
							foreach( $element in $array)
							{
								line 4 $element
							}
						}
						If($Setting.PrinterPropertiesRetention.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Client Printers\Printer properties retention - Value: " 
							switch ($Setting.PrinterPropertiesRetention.Value)
							{
								"SavedOnClientDevice"   {line 4 "Saved on the client device only"}
								"RetainedInUserProfile" {line 4 "Retained in user profile only"}
								"FallbackToProfile"     {line 4 "Held in profile only if not saved on client"}
								"DoNotRetain"           {line 4 "Do not retain printer properties"}
								Default {line 4 "Printer properties retention could not be determined: $($Setting.PrinterPropertiesRetention.Value)"}
							}
						}
						If($Setting.RetainedAndRestoredClientPrinters.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Client Printers\Retained and restored client printers: " $Setting.RetainedAndRestoredClientPrinters.State
						}
						If($Setting.InboxDriverAutoInstallation.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Drivers\Automatic installation of in-box printer drivers: " $Setting.InboxDriverAutoInstallation.State
						}
						If($Setting.UniversalDriverPriority.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Drivers\Universal driver preference - Value: " $Setting.UniversalDriverPriority.Value
						}
						If($Setting.UniversalPrintDriverUsage.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Drivers\Universal print driver usage - Value: " 
							switch ($Setting.UniversalPrintDriverUsage.Value)
							{
								"SpecificOnly"       {line 4 "Use only printer model specific drivers"}
								"UpdOnly"            {line 4 "Use universal printing only"}
								"FallbackToUpd"      {line 4 "Use universal printing only if requested driver is unavailable"}
								"FallbackToSpecific" {line 4 "Use printer model specific drivers only if universal printing is unavailable"}
								Default {line 4 "Universal print driver usage could not be determined: $($Setting.UniversalPrintDriverUsage.Value)"}
							}
						}
						If($Setting.EMFProcessingMode.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Universal Printing\Universal printing EMF processing mode - Value: " 
							switch ($Setting.EMFProcessingMode.Value)
							{
								"ReprocessEMFsForPrinter" {line 4 "Reprocess EMFs for printer"}
								"SpoolDirectlyToPrinter"  {line 4 "Spool directly to printer"}
								Default {line 4 "Universal printing EMF processing mode could not be determined: $($Setting.EMFProcessingMode.Value)"}
							}
						}
						If($Setting.ImageCompressionLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Universal Printing\Universal printing image compression limit - Value: " 
							switch ($Setting.ImageCompressionLimit.Value)
							{
								"NoCompression"       {line 4 "No compression"}
								"LosslessCompression" {line 4 "Best quality (lossless compression)"}
								"MinimumCompression"  {line 4 "High quality"}
								"MediumCompression"   {line 4 "Standard quality"}
								"MaximumCompression"  {line 4 "Reduced quality (maximum compression)"}
								Default {line 4 "Universal printing image compression limit could not be determined: $($Setting.ImageCompressionLimit.Value)"}
							}
						}
						If($Setting.UPDCompressionDefaults.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Universal Printing\Universal printing optimization default - Value: "
							$Tmp = $Setting.UPDCompressionDefaults.Value.replace(";","`n`t`t`t`t")
							line 4 $Tmp
							$Tmp = $null
						}
						If($Setting.UniversalPrintingPreviewPreference.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Universal Printing\Universal printing preview preference - Value: " 
							switch ($Setting.UniversalPrintingPreviewPreference.Value)
							{
								"NoPrintPreview"        {line 4 "Do not use print preview for auto-created or generic universal printers"}
								"AutoCreatedOnly"       {line 4 "Use print preview for auto-created printers only"}
								"GenericOnly"           {line 4 "Use print preview for generic universal printers only"}
								"AutoCreatedAndGeneric" {line 4 "Use print preview for both auto-created and generic universal printers"}
								Default {line 4 "Universal printing preview preference could not be determined: $($Setting.UniversalPrintingPreviewPreference.Value)"}
							}
						}
						If($Setting.DPILimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Printing\Universal Printing\Universal printing print quality limit - Value: " 
							switch ($Setting.DPILimit.Value)
							{
								"Draft"            {line 4 "Draft (150 DPI)"}
								"LowResolution"    {line 4 "Low Resolution (300 DPI)"}
								"MediumResolution" {line 4 "Medium Resolution (600 DPI)"}
								"HighResolution"   {line 4 "High Resolution (1200 DPI)"}
								"Unlimited "       {line 4 "No Limit"}
								Default {line 4 "Universal printing print quality limit could not be determined: $($Setting.DPILimit.Value)"}
							}
						}
						If($Setting.MinimumEncryptionLevel.State -ne "NotConfigured")
						{
							line 3 "ICA\Security\SecureICA minimum encryption level - Value: " 
							switch ($Setting.MinimumEncryptionLevel.Value)
							{
								"Unknown" {line 4 "Unknown encryption"}
								"Basic"   {line 4 "Basic"}
								"LogOn"   {line 4 "RC5 (128 bit) logon only"}
								"Bits40"  {line 4 "RC5 (40 bit)"}
								"Bits56"  {line 4 "RC5 (56 bit)"}
								"Bits128" {line 4 "RC5 (128 bit)"}
								Default {line 4 "SecureICA minimum encryption level could not be determined: $($Setting.MinimumEncryptionLevel.Value)"}
							}
						}
						If($Setting.ConcurrentLogOnLimit.State -ne "NotConfigured")
						{
							line 3 "ICA\Session limits\Concurrent logon limit - Value: " $Setting.ConcurrentLogOnLimit.Value
						}
						If($Setting.SessionDisconnectTimer.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Disconnected session timer: " $Setting.SessionDisconnectTimer.State
						}
						If($Setting.SessionDisconnectTimerInterval.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Disconnected session timer interval - Value (minutes): " $Setting.SessionDisconnectTimerInterval.Value
						}
						If($Setting.LingerDisconnectTimerInterval.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Linger Disconnect Timer Interval - Value (minutes): " $Setting.LingerDisconnectTimerInterval.Value
						}
						If($Setting.LingerTerminateTimerInterval.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Linger Terminate Timer Interval - Value (minutes): " $Setting.LingerTerminateTimerInterval.Value
						}
						If($Setting.PrelaunchDisconnectTimerInterval.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Pre-launch Disconnect Timer Interval - Value (minutes): " $Setting.PrelaunchDisconnectTimerInterval.Value
						}
						If($Setting.PrelaunchTerminateTimerInterval.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Pre-launch Terminate Timer Interval - Value (minutes): " $Setting.PrelaunchTerminateTimerInterval.Value
						}
						If($Setting.SessionConnectionTimer.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Session connection timer: " $Setting.SessionConnectionTimer.State
						}
						If($Setting.SessionConnectionTimerInterval.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Session connection timer interval - Value (minutes): " $Setting.SessionConnectionTimerInterval.Value
						}
						If($Setting.SessionIdleTimer.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Session idle timer: " $Setting.SessionIdleTimer.State
						}
						If($Setting.SessionIdleTimerInterval.State -ne "NotConfigured")
						{
							line 3 "ICA\Session Limits\Session idle timer interval - Value (minutes): " $Setting.SessionIdleTimerInterval.Value
						}
						If($Setting.ShadowInput.State -ne "NotConfigured")
						{
							line 3 "ICA\Shadowing\Input from shadow connections: " $Setting.ShadowInput.State
						}
						If($Setting.ShadowLogging.State -ne "NotConfigured")
						{
							line 3 "ICA\Shadowing\Log shadow attempts: " $Setting.ShadowLogging.State
						}
						If($Setting.ShadowUserNotification.State -ne "NotConfigured")
						{
							line 3 "ICA\Shadowing\Notify user of pending shadow connections: " $Setting.ShadowUserNotification.State
						}
						If($Setting.ShadowAllowList.State -ne "NotConfigured")
						{
							Line 3 "ICA\Shadowing\Users who can shadow other users - Value: " 
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
								Line 4 $tmp
							}
						}
						If($Setting.ShadowDenyList.State -ne "NotConfigured")
						{
							line 3 "ICA\Shadowing\Users who cannot shadow other users - Value: " 
							$array = $Setting.ShadowDenyList.Values
							foreach( $element in $array)
							{
								$x = $element.indexof("/",8)
								$tmp = $element.substring(8,$x-8)
								Line 4 $tmp
							}
						}
						If($Setting.LocalTimeEstimation.State -ne "NotConfigured")
						{
							line 3 "ICA\Time Zone Control\Estimate local time for legacy clients: " $Setting.LocalTimeEstimation.State
						}
						If($Setting.SessionTimeZone.State -ne "NotConfigured")
						{
							line 3 "ICA\Time Zone Control\Use local time of client - Value: " 
							switch ($Setting.SessionTimeZone.Value)
							{
								"UseServerTimeZone" {line 4 "Use server time zone"}
								"UseClientTimeZone" {line 4 "Use client time zone"}
								Default {line 4 "Use local time of client could not be determined: $($Setting.SessionTimeZone.Value)"}
							}
						}
						If($Setting.TwainRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\TWAIN devices\Client TWAIN device redirection: " $Setting.TwainRedirection.State
						}
						If($Setting.TwainCompressionLevel.State -ne "NotConfigured")
						{
							line 3 "ICA\TWAIN devices\TWAIN compression level - Value: " 
							switch ($Setting.TwainCompressionLevel.Value)
							{
								"None"   {line 4 "None"}
								"Low"    {line 4 "Low"}
								"Medium" {line 4 "Medium"}
								"High"   {line 4 "High"}
								Default {line 4 "TWAIN compression level could not be determined: $($Setting.TwainCompressionLevel.Value)"}
							}
						}
						If($Setting.UsbDeviceRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\USB devices\Client USB device redirection: " $Setting.UsbDeviceRedirection.State
						}
						If($Setting.UsbDeviceRedirectionRules.State -ne "NotConfigured")
						{
							line 3 "ICA\USB devices\Client USB device redirection rules - Values: " 
							$array = $Setting.UsbDeviceRedirectionRules.Values
							foreach( $element in $array)
							{
								line 4 $element
							}
						}
						If($Setting.UsbPlugAndPlayRedirection.State -ne "NotConfigured")
						{
							line 3 "ICA\USB devices\Client USB Plug and Play device redirection: " $Setting.UsbPlugAndPlayRedirection.State
						}
						If($Setting.FramesPerSecond.State -ne "NotConfigured")
						{
							line 3 "ICA\Visual Display\Max Frames Per Second - Value (fps): " $Setting.FramesPerSecond.Value
						}
						If($Setting.ProgressiveCompressionLevel.State -ne "NotConfigured")
						{
							line 3 "ICA\Visual Display\Moving Images\Progressive compression level - Value: " -nonewline
							switch ($Setting.ProgressiveCompressionLevel.Value)
							{
								"UltraHigh" {line 0 "Ultra high"}
								"VeryHigh"  {line 0 "Very high"}
								"High"      {line 0 "High"}
								"Normal"    {line 0 "Normal"}
								"Low"       {line 0 "Low"}
								Default {line 0 "Progressive compression level could not be determined: $($Setting.ProgressiveCompressionLevel.Value)"}
							}
						}
						If($Setting.ProgressiveCompressionThreshold.State -ne "NotConfigured")
						{
							line 3 "ICA\Visual Display\Moving Images\Progressive compression threshold value - Value (Kbps): " $Setting.ProgressiveCompressionThreshold.Value
						}
						If($Setting.ExtraColorCompression.State -ne "NotConfigured")
						{
							line 3 "ICA\Visual Display\Still Images\Extra Color Compression: " $Setting.ExtraColorCompression.State
						}
						If($Setting.ExtraColorCompressionThreshold.State -ne "NotConfigured")
						{
							line 3 "ICA\Visual Display\Still Images\Extra Color Compression Threshold - Value (Kbps): " $Setting.ExtraColorCompressionThreshold.Value
						}
						If($Setting.ProgressiveHeavyweightCompression.State -ne "NotConfigured")
						{
							line 3 "ICA\Visual Display\Still Images\Heavyweight compression: " $Setting.ProgressiveHeavyweightCompression.State
						}
						If($Setting.LossyCompressionLevel.State -ne "NotConfigured")
						{
							line 3 "ICA\Visual Display\Still Images\Lossy compression level - Value: " 
							switch ($Setting.LossyCompressionLevel.Value)
							{
								"None"   {line 4 "None"}
								"Low"    {line 4 "Low"}
								"Medium" {line 4 "Medium"}
								"High"   {line 4 "High"}
								Default {line 4 "Lossy compression level could not be determined: $($Setting.LossyCompressionLevel.Value)"}
							}
						}
						If($Setting.LossyCompressionThreshold.State -ne "NotConfigured")
						{
							line 3 "ICA\Visual Display\Still Images\Lossy compression threshold value - Value (Kbps): " $Setting.LossyCompressionThreshold.Value
						}
						If($Setting.SessionImportance.State -ne "NotConfigured")
						{
							line 3 "Server Session Settings\Session importance - Value: " 
							switch ($Setting.SessionImportance.Value)
							{
								"Low"    {line 4 "Low"}
								"Normal" {line 4 "Normal"}
								"High"   {line 4 "High"}
								Default {line 4 "Session importance could not be determined: $($Setting.SessionImportance.Value)"}
							}
						}
						If($Setting.SingleSignOn.State -ne "NotConfigured")
						{
							line 3 "Server Session Settings\Single Sign-On: " $Setting.SingleSignOn.State
						}
						If($Setting.SingleSignOnCentralStore.State -ne "NotConfigured")
						{
							line 3 "Server Session Settings\Single Sign-On central store - Value: " $Setting.SingleSignOnCentralStore.Value
						}
					}
				}
			}
			Else
			{
				line 2 "Unable to retrieve settings"
			}
		
			Write-Output $global:output
			$global:output = $null
			$Filter = $null
			$Settings = $null
		}
	}
	Else 
	{
		line 0 "Citrix Policy information could not be retrieved."
	}
		
	$Policies = $null
	$global:output = $null
}
