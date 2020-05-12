#Original Script created 8/17/2010 by Michael Bogobowicz, Citrix Systems.
#To contact, please message @mcbogo on Twitter
#This script is designed to be run on a XenApp 6 server

#Modifications by Carl Webster, CTP and independent consultant
#webster@carlwebster.com
#@carlwebster on Twitter
#http://www.CarlWebster.com
#modified from the original script for XenApp 6.5
#originally released to the Citrix community on October 7, 2011
#update October 9, 2011: fixed the formatting of the Health Monitoring & Recovery policy setting

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

#Script begins
$global:output = ""

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
		line 1 "Administrator type: "$Administrator.AdministratorType -nonewline
		line 0 " Administrator"
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
				line 2 $farmprivilege
			}
	
			line 1 "Folder Privileges:"
			ForEach($folderprivilege in $Administrator.FolderPrivileges) 
			{
				$test = $folderprivilege.ToString()
				$folderlabel = $test.substring(0, $test.IndexOf(":") + 1)
				line 2 $folderlabel
				$test1 = $test.substring($test.IndexOf(":") + 1)
				$folderpermissions = $test1.replace(",","`n`t`t`t")
				line 3 $folderpermissions
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
		line 2 "Application name (Browser name): " $Application.BrowserName
		line 2 "Disable application: " -NoNewLine
		If ($Application.Enabled) 
		{
		  line 0 "False"
		} 
		Else
		{
		  line 0 "True"
		}
		line 2 "Hide disabled application: " $Application.HideWhenDisabled
		line 2 "Application description: " $Application.Description
	
		#type properties
		line 2 "Application Type: " $Application.ApplicationType
		line 2 "Folder path: " $Application.FolderPath
		line 2 "Content Address: " $Application.ContentAddress
	
		#if a streamed app
		If($streamedapp)
		{
			line 2 "Citrix streaming application profile address: " $Application.ProfileLocation
			line 2 "Application to launch from the Citrix streaming application profile: " $Application.ProfileProgramName
			line 2 "Extra command line parameters: " $Application.ProfileProgramArguments
			#if streamed, Offline access properties
			If($Application.OfflineAccessAllowed)
			{
				line 2 "Enable offline access: " $Application.OfflineAccessAllowed
			}
			If($Application.CachingOption)
			{
				line 2 "Cache preference: " $Application.CachingOption
			}
		}
		
		#location properties
		If(!$streamedapp)
		{
			line 2 "Command line: " $Application.CommandLineExecutable
			line 2 "Working directory: " $Application.WorkingDirectory
			
			#servers properties
			If($AppServerInfoResults)
			{
				line 2 "Servers:"
				ForEach($servername in $AppServerInfo.ServerNames)
				{
					line 3 $servername
				}
				line 2 "Workergroups:"
				ForEach($workergroup in $AppServerInfo.WorkerGroupNames)
				{
					line 3 $workergroup
				}
			}
			Else
			{
				line 3 "Unable to retrieve a list of Servers for this application"
				line 3 "Unable to retrieve a list of Worker Groups for this application"
			}
		}
	
		#users properties
		If($Application.AnonymousConnectionsAllowed)
		{
			line 2 "Allow anonymous users: " $Application.AnonymousConnectionsAllowed
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
		line 2 "Client application folder: " $Application.ClientFolder
		If($Application.AddToClientStartMenu)
		{
			line 2 "Add to client's start menu: " $Application.AddToClientStartMenu
		}
		If($Application.StartMenuFolder)
		{
			line 2 "Start menu folder: " $Application.StartMenuFolder
		}
		If($Application.AddToClientDesktop)
		{
			line 2 "Add shortcut to the client's desktop: " $Application.AddToClientDesktop
		}
	
		#access control properties
		If($Application.ConnectionsThroughAccessGatewayAllowed)
		{
			line 2 "Allow connections made through AGAE: " $Application.ConnectionsThroughAccessGatewayAllowed
		}
		If($Application.OtherConnectionsAllowed)
		{
			line 2 "Any connection: " $Application.OtherConnectionsAllowed
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
				line 2 "No File Type Associations exist for this application"
			}
		}
		Else
		{
			line 2 "Unable to retrieve the list of File Type Associations for this application"
		}
	
		#if streamed app, Alternate profiles
		If($streamedapp)
		{
			If($Application.AlternateProfiles)
			{
				line 2 "Primary application profile location: " $Application.AlternateProfiles
			}
		
			#if streamed app, User privileges properties
			If($Application.RunAsLeastPrivilegedUser)
			{
				line 2 "Run application as a least-privileged user account: " $Application.RunAsLeastPrivilegedUser
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
	
		line 2 "Allow only one instance of application for each user: " -NoNewLine
	
		If ($Application.MultipleInstancesPerUserAllowed) 
		{
			line 0 "False"
		} 
		Else
		{
			line 0 "True"
		}
	
		If($Application.CpuPriorityLevel)
		{
			line 2 "Application importance: " $Application.CpuPriorityLevel
		}
		
		#client options properties
		If($Application.AudioRequired)
		{
			line 2 "Enable legacy audio: " $Application.AudioRequired
		}
		If($Application.AudioType)
		{
			line 2 "Minimum requirement: " $Application.AudioType
		}
		If($Application.SslConnectionEnable)
		{
			line 2 "Enable SSL and TLS protocols: " $Application.SslConnectionEnabled
		}
		If($Application.EncryptionLevel)
		{
			line 2 "Encryption: " $Application.EncryptionLevel
		}
		If($Application.EncryptionRequire)
		{
			line 2 "Minimum requirement: " $Application.EncryptionRequired
		}
	
		line 2 "Start this application without waiting for printers to be created: " -NoNewLine
		If ($Application.WaitOnPrinterCreation) 
		{
			line 0 "False"
		} 
		Else
		{
			line 0 "True"
		}
		
		#appearance properties
		If($Application.WindowType)
		{
			line 2 "Session window size: " $Application.WindowType
		}
		If($Application.ColorDepth)
		{
			line 2 "Maximum color quality: " $Application.ColorDepth
		}
		If($Application.TitleBarHidden)
		{
			line 2 "Hide application title bar: " $Application.TitleBarHidden
		}
		If($Application.MaximizedOnStartup)
		{
			line 2 "Maximize application at startup: " $Application.MaximizedOnStartup
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
				Line 1 "Date: " $ConfigLogItem.Date
				Line 1 "Account: " $ConfigLogItem.Account
				Line 1 "Change description: " $ConfigLogItem.Description
				Line 1 "Type of change: " $ConfigLogItem.TaskType
				Line 1 "Type of item: " $ConfigLogItem.ItemType
				Line 1 "Name of item: " $ConfigLogItem.ItemName
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
	
		line 1 "Load balancing policy name: " $LoadBalancingPolicy.PolicyName
		line 2 "Load balancing policy description: " $LoadBalancingPolicy.Description
		line 2 "Load balancing policy enabled: " $LoadBalancingPolicy.Enabled
		line 2 "Load balancing policy priority: " $LoadBalancingPolicy.Priority
	
		line 2 "Filter based on Access Control: " $LoadBalancingPolicyFilter.AccessControlEnabled
		If($LoadBalancingPolicyFilter.AccessControlEnabled)
		{
			line 2 "Apply to connections made through Access Gateway: " $LoadBalancingPolicyFilter.AllowConnectionsThroughAccessGateway
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
				line 3 "Apply to all client IP addresses: " $LoadBalancingPolicyFilter.ApplyToAllClientIPAddresses
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
				line 3 "Apply to all client names: " $LoadBalancingPolicyFilter.ApplyToAllClientNames
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
			line 3 "Apply to anonymous users: " $LoadBalancingPolicyFilter.ApplyToAnonymousAccounts
			If($LoadBalancingPolicyFilter.ApplyToAllExplicitAccounts)
			{
				line 3 "Apply to all explicit (non-anonymous) users: " $LoadBalancingPolicyFilter.ApplyToAllExplicitAccounts
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
		If($LoadBalancingPolicyConfiguration.StreamingDeliveryProtocolState)
		{
			line 2 "Set the delivery protocols for applications streamed to client"
			line 3 $LoadBalancingPolicyConfiguration.StreamingDeliveryOption
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
			line 3 "Report full load when the number of context switches per second is greater than this value: " $LoadEvaluator.ContextSwitches[1]
			line 3 "Report no load when the number of context switches per second is less than or equal to this value: " $LoadEvaluator.ContextSwitches[0]
		}
	
		If($LoadEvaluator.CpuUtilizationEnabled)
		{
			line 2 "CPU Utilization Settings"
			line 3 "Report full load when the processor utilization percentage is greater than this value: " $LoadEvaluator.CpuUtilization[1]
			line 3 "Report no load when the processor utilization percentage is less than or equal to this value: " $LoadEvaluator.CpuUtilization[0]
		}
	
		If($LoadEvaluator.DiskDataIOEnabled)
		{
			line 2 "Disk Data I/O Settings"
			line 3 "Report full load when the total disk I/O in kilobytes per second is greater than this value: " $LoadEvaluator.DiskDataIO[1]
			line 3 "Report no load when the total disk I/O in kilobytes per second is less than or equal to this value: " $LoadEvaluator.DiskDataIO[0]
		}
	
		If($LoadEvaluator.DiskOperationsEnabled)
		{
			line 2 "Disk Operations Settings"
			line 3 "Report full load when the total number of read and write operations per second is greater than this value: " $LoadEvaluator.DiskOperations[1]
			line 3 "Report no load when the total number of read and write operations per second is less than or equal to this value: " $LoadEvaluator.DiskOperations[0]
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
			line 3 "Impact of logons on load: " $LoadEvaluator.LoadThrottling
			
		}
	
		If($LoadEvaluator.MemoryUsageEnabled)
		{
			line 2 "Memory Usage Settings"
			line 3 "Report full load when the memory usage is greater than this value: " $LoadEvaluator.MemoryUsage[1]
			line 3 "Report no load when the memory usage is less than or equal to this value: " $LoadEvaluator.MemoryUsage[0]
		}
	
		If($LoadEvaluator.PageFaultsEnabled)
		{
			line 2 "Page Faults Settings"
			line 3 "Report full load when the number of page faults per second is greater than this value: " $LoadEvaluator.PageFaults[1]
			line 3 "Report no load when the number of page faults per second is less than or equal to this value: " $LoadEvaluator.PageFaults[0]
		}
	
		If($LoadEvaluator.PageSwapsEnabled)
		{
			line 2 "Page Swaps Settings"
			line 3 "Report full load when the number of page swaps per second is greater than this value: " $LoadEvaluator.PageSwaps[1]
			line 3 "Report no load when the number of page swaps per second is less than or equal to this value: " $LoadEvaluator.PageSwaps[0]
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
		line 2 "Server FQDN: " $server.ServerFqdn
		line 2 "Product: " $server.CitrixProductName -NoNewLine
		line 0 ", " $server.CitrixEdition -NoNewLine
		line 0 " Edition"
		line 2 "Version: " $server.CitrixVersion
		line 2 "Service Pack: " $server.CitrixServicePack
		line 2 "Operating System Type: " -NoNewLine
		If($server.Is64Bit)
		{
			line 0 "64 bit"
		} 
		Else 
		{
			line 0 "32 bit"
		}
		line 2 "IP Address: " $server.IPAddresses
		line 2 "Logon: " -NoNewLine
		If($server.LogOnsEnabled)
		{
			line 0 "Enabled"
		} 
		Else 
		{
			line 0 "Disabled"
		}
		line 2 "Logon Control Mode: " $Server.LogOnMode
		line 2 "Product Installation Date: " $server.CitrixInstallDate
		line 2 "Operating System Version: " $server.OSVersion -NoNewLine
		line 0 " " $server.OSServicePack
		line 2 "Zone: " $server.ZoneName
		line 2 "Election Preference: " $server.ElectionPreference
		line 2 "Folder: " $server.FolderPath
		line 2 "Product Installation Path: " $server.CitrixInstallPath
		If($server.LicenseServerName)
		{
			line 2 "License Server Name: " $server.LicenseServerName
			line 2 "License Server Port: " $server.LicenseServerPortNumber
		}
		If($server.ICAPortNumber -gt 0)
		{
			line 2 "ICA Port Number: " $server.ICAPortNumber
		}
		If($server.RDPPortNumber -gt 0)
		{
			line 2 "RDP Port Number: " $server.RDPPortNumber
		}
		line 2 "Is the Print Spooler on this server healthy: " $Server.IsSpoolerHealthy
		line 2 "Power Management Control Mode: " $server.PcmMode
		
		#applications published to server
		$Applications = Get-XAApplication -ServerName $server.ServerName -EA 0 | sort-object FolderPath, DisplayName
		If( $? -and $Applications )
		{
			line 2 "Published applications:"
			ForEach($app in $Applications)
			{
				line 0 ""
				line 3 "Display name: " $app.DisplayName
				line 3 "Folder path: " $app.FolderPath
			}
		}
		#Citrix hotfixes installed
		$hotfixes = Get-XAServerHotfix -ServerName $server.ServerName -EA 0 | sort-object HotfixName
		If( $? -and $hotfixes )
		{
			line 0 ""
			line 2 "Citrix Hotfixes:"
			ForEach($hotfix in $hotfixes)
			{
				line 0 ""
				line 3 "Hotfix: " $hotfix.HotfixName
				line 3 "Installed by: " $hotfix.InstalledBy
				line 3 "Installed date: " $hotfix.InstalledOn
				line 3 "Hotfix type: " $hotfix.HotfixType
				line 3 "Valid: " $hotfix.Valid
				line 3 "Hotfixes replaced: "
				ForEach($Replaced in $hotfix.HotfixesReplaced)
				{
					line 4 $Replaced
				}
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
			line 2 "Organization Units:"
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
				line 3 "Folder path: " $app.FolderPath
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
				line 0  " " $server.ElectionPreference
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

Echo "Please wait while Citrix Policies are retrieved..."
$Policies = Get-CtxGroupPolicy -EA 0 | sort-object PolicyName
If( $? )
{
	line 0 ""
	line 0 "Policies:"
	ForEach($Policy in $Policies)
	{
		line 1 "Policy Name: " $Policy.PolicyName
		line 2 "Type: " $Policy.Type
		line 2 "Description: " $Policy.Description
		line 2 "Enabled: " $Policy.Enabled
		line 2 "Priority: " $Policy.Priority

		$filter = Get-CtxGroupPolicyFilter -PolicyName $Policy.PolicyName -EA 0

		If( $? )
		{
			If($Filter -and $filter.FilterName -and ($filter.FilterName.Trim() -ne ""))
			{
				Line 2 "Filter name: " $filter.FilterName
				Line 2 "Filter type: " $filter.FilterType
				Line 2 "Filter enabled: " $filter.Enabled
				Line 2 "Filter mode: " $filter.Mode
				Line 2 "Filter value: " $filter.FilterValue

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
						line 3 "ICA\ICA listener connection timeout - Value: " $Setting.IcaListenerTimeout.Value
					}
					If($Setting.IcaListenerPortNumber.State -ne "NotConfigured")
					{
						line 3 "ICA\ICA listener port number - Value: " $Setting.IcaListenerPortNumber.Value
					}
					If($Setting.AutoClientReconnect.State -ne "NotConfigured")
					{
						line 3 "ICA\Auto Client Reconnect\Auto client reconnect - Value: " $Setting.AutoClientReconnect.State
					}
					If($Setting.AutoClientReconnectLogging.State -ne "NotConfigured")
					{
						line 3 "ICA\Auto Client Reconnect\Auto client reconnect logging - Value: " $Setting.AutoClientReconnectLogging.Value
					}
					If($Setting.IcaRoundTripCalculation.State -ne "NotConfigured")
					{
						line 3 "ICA\End User Monitoring\ICA round trip calculation - Value: " $Setting.IcaRoundTripCalculation.State
					}
					If($Setting.IcaRoundTripCalculationInterval.State -ne "NotConfigured")
					{
						line 3 "ICA\End User Monitoring\ICA round trip calculation interval (Seconds) - Value: " $Setting.IcaRoundTripCalculationInterval.Value
					}
					If($Setting.IcaRoundTripCalculationWhenIdle.State -ne "NotConfigured")
					{
						line 3 "ICA\End User Monitoring\ICA round trip calculations for idle connections - Value: " $Setting.IcaRoundTripCalculationWhenIdle.State
					}
					If($Setting.DisplayMemoryLimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Graphics\Display memory limit: " $Setting.DisplayMemoryLimit.Value
					}
					If($Setting.DisplayDegradePreference.State -ne "NotConfigured")
					{
						line 3 "ICA\Graphics\Display mode degrade preference: " $Setting.DisplayDegradePreference.Value
					}
					If($Setting.DynamicPreview.State -ne "NotConfigured")
					{
						line 3 "ICA\Graphics\Dynamic Windows Preview: " $Setting.DynamicPreview.State
					}
					If($Setting.ImageCaching.State -ne "NotConfigured")
					{
						line 3 "ICA\Graphics\Image caching - Value: " $Setting.ImageCaching.State
					}
					If($Setting.MaximumColorDepth.State -ne "NotConfigured")
					{
						line 3 "ICA\Graphics\Maximum allowed color depth: " $Setting.MaximumColorDepth.Value
					}
					If($Setting.DisplayDegradeUserNotification.State -ne "NotConfigured")
					{
						line 3 "ICA\Graphics\Notify user when display mode is degraded - Value: " $Setting.DisplayDegradeUserNotification.State
					}
					If($Setting.QueueingAndTossing.State -ne "NotConfigured")
					{
						line 3 "ICA\Graphics\Queueing and tossing - Value: " $Setting.QueueingAndTossing.State
					}
					If($Setting.PersistentCache.State -ne "NotConfigured")
					{
						line 3 "ICA\Graphics\Caching\Persistent Cache Threshold - Value: " $Setting.PersistentCache.Value
					}
					If($Setting.IcaKeepAliveTimeout.State -ne "NotConfigured")
					{
						line 3 "ICA\Keep ALive\ICA keep alive timeout - Value: " $Setting.IcaKeepAliveTimeout.Value
					}
					If($Setting.IcaKeepAlives.State -ne "NotConfigured")
					{
						line 3 "ICA\Keep ALive\ICA keep alives - Value: " $Setting.IcaKeepAlives.Value
					}
					If($Setting.MultimediaConferencing.State -ne "NotConfigured")
					{
						line 3 "ICA\Multimedia\Multimedia conferencing - Value: " $Setting.MultimediaConferencing.State
					}
					If($Setting.MultimediaAcceleration.State -ne "NotConfigured")
					{
						line 3 "ICA\Multimedia\Windows Media Redirection - Value: " $Setting.MultimediaAcceleration.State
					}
					If($Setting.MultimediaAccelerationDefaultBufferSize.State -ne "NotConfigured")
					{
						line 3 "ICA\Multimedia\Windows Media Redirection Buffer Size - Value: " $Setting.MultimediaAccelerationDefaultBufferSize.Value
					}
					If($Setting.MultimediaAccelerationUseDefaultBufferSize.State -ne "NotConfigured")
					{
						line 3 "ICA\Multimedia\Windows Media Redirection Buffer Size Use - Value: " $Setting.MultimediaAccelerationUseDefaultBufferSize.State
					}
					If($Setting.MultiPortPolicy.State -ne "NotConfigured")
					{
						line 3 "ICA\MultiStream Connections\Multi-Port Policy - Value: " 
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
						line 3 "ICA\MultiStream Connections\Multi-Stream - Value: " $Setting.MultiStreamPolicy.State
					}
					If($Setting.PromptForPassword.State -ne "NotConfigured")
					{
						line 3 "ICA\Security\Prompt for password - Value: " $Setting.PromptForPassword.State
					}
					If($Setting.IdleTimerInterval.State -ne "NotConfigured")
					{
						line 3 "ICA\Server Limits\Server idle timer interval - Value: " $Setting.IdleTimerInterval.Value
					}
					If($Setting.SessionReliabilityConnections.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Reliability\Session reliability connections - Value: " $Setting.SessionReliabilityConnections.State
					}
					If($Setting.SessionReliabilityPort.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Reliability\Session reliability port number - Value: " $Setting.SessionReliabilityPort.Value
					}
					If($Setting.SessionReliabilityTimeout.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Reliability\Session reliability timeout - Value: " $Setting.SessionReliabilityTimeout.Value
					}
					If($Setting.Shadowing.State -ne "NotConfigured")
					{
						line 3 "ICA\Shadowing\Shadowing - Value: " $Setting.Shadowing.State
					}
					If($Setting.LicenseServerHostName.State -ne "NotConfigured")
					{
						line 3 "Licensing\License server host name: " $Setting.LicenseServerHostName.Value
					}
					If($Setting.LicenseServerPort.State -ne "NotConfigured")
					{
						line 3 "Licensing\License server port: " $Setting.LicenseServerPort.Value
					}
					If($Setting.FarmName.State -ne "NotConfigured")
					{
						line 3 "Power and Capacity Management\Farm name: " $Setting.FarmName.Value
					}
					If($Setting.WorkloadName.State -ne "NotConfigured")
					{
						line 3 "Power and Capacity Management\Workload name: " $Setting.WorkloadName.Value
					}
					If($Setting.ConnectionAccessControl.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Connection access control - Value: " $Setting.ConnectionAccessControl.Value
					}
					If($Setting.DnsAddressResolution.State -ne "NotConfigured")
					{
						line 3 "Server Settings\DNS address resolution - Value: " $Setting.DnsAddressResolution.State
					}
					If($Setting.FullIconCaching.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Full icon caching - Value: " $Setting.FullIconCaching.State
					}
					If($Setting.LoadEvaluator.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Load Evaluator Name - Value: " $Setting.LoadEvaluator.Value
					}
					If($Setting.ProductEdition.State -ne "NotConfigured")
					{
						line 3 "Server Settings\XenApp product edition - Value: " $Setting.ProductEdition.Value
					}
					If($Setting.ProductModel.State -ne "NotConfigured")
					{
						line 3 "Server Settings\XenApp product model - Value: " $Setting.ProductModel.Value
					}
					If($Setting.UserSessionLimit.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Connection Limits\Limit user sessions - Value: " $Setting.UserSessionLimit.Value
					}
					If($Setting.UserSessionLimitAffectsAdministrators.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Connection Limits\Limits on administrator sessions - Value: " $Setting.UserSessionLimitAffectsAdministrators.State
					}
					If($Setting.UserSessionLimitLogging.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Connection Limits\Logging of logon limit events - Value: " $Setting.UserSessionLimitLogging.State
					}
					If($Setting.HealthMonitoring.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Health Monitoring and Recovery\Health monitoring - Value: " $Setting.HealthMonitoring.State
					}
					If($Setting.HealthMonitoringTests.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Health Monitoring and Recovery\Health monitoring tests - Value: " 
						[xml]$XML = $Setting.HealthMonitoringTests.Value
						ForEach($Test in $xml.hmrtests.tests.test)
						{
							line 4 "Name: " $test.name
							line 4 "File Location: " $test.file
							If($test.arguments)
							{
								line 4 "Arguments: " $test.arguments
							}
							line 4 "Description: " $test.description
							line 4 "Interval (seconds): " $test.interval
							line 4 "Time-out (seconds): " $test.timeout
							line 4 "Threshold: " $test.threshold
							line 4 "Recovery action: " $test.recoveryAction
							line 0 ""
						}
					}
					If($Setting.MaximumServersOfflinePercent.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Health Monitoring and Recovery\Maximum percent of servers with logon control - Value: " $Setting.MaximumServersOfflinePercent.Value
					}
					If($Setting.CpuManagementServerLevel.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Memory/CPU\CPU management server level - Value: " $Setting.CpuManagementServerLevel.Value
					}
					If($Setting.MemoryOptimization.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Memory/CPU\Memory optimization - Value: " $Setting.MemoryOptimization.State
					}
					If($Setting.MemoryOptimizationExcludedPrograms.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Memory/CPU\Memory optimization application exclusion list - Value: " $Setting.MemoryOptimizationExcludedPrograms.Value
					}
					If($Setting.MemoryOptimizationIntervalType.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Memory/CPU\Memory optimization interval - Value: " $Setting.MemoryOptimizationIntervalType.Value
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
						line 3 "Server Settings\Memory/CPU\Memory optimization schedule: time - Value: " $Setting.MemoryOptimizationTime.Value
					}
					If($Setting.OfflineClientTrust.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Offline Applications\Offline app client trust - Value: " $Setting.OfflineClientTrust.State
					}
					If($Setting.OfflineEventLogging.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Offline Applications\Offline app event logging - Value: " $Setting.OfflineEventLogging.State
					}
					If($Setting.OfflineLicensePeriod.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Offline Applications\Offline app license period - Value: " $Setting.OfflineLicensePeriod.Value
					}
					If($Setting.OfflineUsers.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Offline Applications\Offline app users - Value: " $Setting.OfflineUsers.Value
					}
					If($Setting.RebootCustomMessage.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot custom warning - Value: " $Setting.RebootCustomMessage.State
					}
					If($Setting.RebootCustomMessageText.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot custom warning text - Value: " $Setting.RebootCustomMessageText.Value
					}
					If($Setting.RebootDisableLogOnTime.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot logon disable time - Value: " $Setting.RebootDisableLogOnTime.Value
					}
					If($Setting.RebootScheduleFrequency.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot schedule frequency - Value: " $Setting.RebootScheduleFrequency.Value
					}
					If($Setting.RebootScheduleRandomizationInterval.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot schedule randomization interval - Value: " $Setting.RebootScheduleRandomizationInterval.Value
					}
					If($Setting.RebootScheduleStartDate.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot schedule start date - Value: " $Setting.RebootScheduleStartDate.Value
					}
					If($Setting.RebootScheduleTime.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot schedule time - Value: " $Setting.RebootScheduleTime.Value
					}
					If($Setting.RebootWarningInterval.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot warning interval - Value: " $Setting.RebootWarningInterval.Value
					}
					If($Setting.RebootWarningStartTime.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot warning start time - Value: " $Setting.RebootWarningStartTime.Value
					}
					If($Setting.RebootWarningMessage.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Reboot warning to users - Value: " $Setting.RebootWarningMessage.State
					}
					If($Setting.ScheduledReboots.State -ne "NotConfigured")
					{
						line 3 "Server Settings\Reboot Behavior\Scheduled reboots  - Value: " $Setting.ScheduledReboots.State
					}
					If($Setting.FilterAdapterAddresses.State -ne "NotConfigured")
					{
						line 3 "Virtual IP\Virtual IP adapter address filtering - Value: " $Setting.FilterAdapterAddresses.State
					}
					If($Setting.EnhancedCompatibilityPrograms.State -ne "NotConfigured")
					{
						line 3 "Virtual IP\Virtual IP compatibility programs list - Value: " $Setting.EnhancedCompatibilityPrograms.Value
					}
					If($Setting.EnhancedCompatibility.State -ne "NotConfigured")
					{
						line 3 "Virtual IP\Virtual IP enhanced compatibility - Value: " $Setting.EnhancedCompatibility.State
					}
					If($Setting.FilterAdapterAddressesPrograms.State -ne "NotConfigured")
					{
						line 3 "Virtual IP\Virtual IP filter adapter addresses programs list - Value: " $Setting.FilterAdapterAddressesPrograms.Value
					}
					If($Setting.VirtualLoopbackSupport.State -ne "NotConfigured")
					{
						line 3 "Virtual IP\Virtual IP loopback support - Value: " $Setting.VirtualLoopbackSupport.State
					}
					If($Setting.VirtualLoopbackPrograms.State -ne "NotConfigured")
					{
						line 3 "Virtual IP\Virtual IP virtual loopback programs list - Value: " $Setting.VirtualLoopbackPrograms.Value
					}
					If($Setting.TrustXmlRequests.State -ne "NotConfigured")
					{
						line 3 "XML Service\Trust XML requests - Value: " $Setting.TrustXmlRequests.State
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
						line 3 "ICA\Client clipboard redirection - Value: " $Setting.ClipboardRedirection.State
					}
					If($Setting.DesktopLaunchForNonAdmins.State -ne "NotConfigured")
					{
						line 3 "ICA\Desktop launches - Value: " $Setting.DesktopLaunchForNonAdmins.State
					}
					If($Setting.NonPublishedProgramLaunching.State -ne "NotConfigured")
					{
						line 3 "ICA\Launching of non-published programs during client connection - Value: " $Setting.NonPublishedProgramLaunching.State
					}
					If($Setting.FlashAcceleration.State -ne "NotConfigured")
					{
						line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash acceleration - Value: " $Setting.FlashAcceleration.State
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
						line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash backwards compatibility - Value: " $Setting.FlashBackwardsCompatibility.State
					}
					If($Setting.FlashDefaultBehavior.State -ne "NotConfigured")
					{
						line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash default behavior - Value: " $Setting.FlashDefaultBehavior.Value
					}
					If($Setting.FlashEventLogging.State -ne "NotConfigured")
					{
						line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash event logging - Value: " $Setting.FlashEventLogging.State
					}
					If($Setting.FlashIntelligentFallback.State -ne "NotConfigured")
					{
						line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash intelligent fallback - Value: " $Setting.FlashIntelligentFallback.State
					}
					If($Setting.FlashLatencyThreshold.State -ne "NotConfigured")
					{
						line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash latency threshold - Value: " $Setting.FlashLatencyThreshold.Value
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
						line 3 "ICA\Adobe Flash Delivery\Flash Redirection\Flash URL compatibility list - Values: " 
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
						line 3 "ICA\Adobe Flash Delivery\Legacy Server Side Optimizations\Flash quality adjustment - Value: " $Setting.AllowSpeedFlash.Value
					}
					If($Setting.AudioPlugNPlay.State -ne "NotConfigured")
					{
						line 3 "ICA\Audio\Audio Plug N Play " $Setting.AudioPlugNPlay.State
					}
					If($Setting.AudioQuality.State -ne "NotConfigured")
					{
						line 3 "ICA\Audio\Audio quality - Value: " $Setting.AudioQuality.Value
					}
					If($Setting.ClientAudioRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\Audio\Client audio redirection - Value: " $Setting.ClientAudioRedirection.State
					}
					If($Setting.MicrophoneRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\Audio\Client microphone redirection - Value: " $Setting.MicrophoneRedirection.State
					}
					If($Setting.AudioBandwidthLimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\Audio redirection bandwidth limit - Value: " $Setting.AudioBandwidthLimit.Value
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
						line 3 "ICA\Bandwidth\Clipboard redirection bandwidth limit - Value: " $Setting.ClipboardBandwidthLimit.Value
					}
					If($Setting.ClipboardBandwidthPercent.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\Clipboard redirection bandwidth limit percent - Value: " $Setting.ClipboardBandwidthPercent.Value
					}
					If($Setting.ComPortBandwidthLimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\COM port redirection bandwidth limit - Value: " $Setting.ComPortBandwidthLimit.Value
					}
					If($Setting.ComPortBandwidthPercent.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\COM port redirection bandwidth limit percent - Value: " $Setting.ComPortBandwidthPercent.Value
					}
					If($Setting.FileRedirectionBandwidthLimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\File redirection bandwidth limit - Value: " $Setting.FileRedirectionBandwidthLimit.Value
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
						line 3 "ICA\Bandwidth\LPT port redirection bandwidth limit - Value: " $Setting.LptBandwidthLimit.Value
					}
					If($Setting.LptBandwidthLimitPercent.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\LPT port redirection bandwidth limit percent - Value: " $Setting.LptBandwidthLimitPercent.Value
					}
					If($Setting.OverallBandwidthLimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\Overall session bandwidth limit - Value: " $Setting.OverallBandwidthLimit.Value
					}
					If($Setting.PrinterBandwidthLimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\Printer redirection bandwidth limit - Value: " $Setting.PrinterBandwidthLimit.Value
					}
					If($Setting.PrinterBandwidthPercent.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\Printer redirection bandwidth limit percent - Value: " $Setting.PrinterBandwidthPercent.Value
					}
					If($Setting.TwainBandwidthLimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\TWAIN device redirection bandwidth limit - Value: " $Setting.TwainBandwidthLimit.Value
					}
					If($Setting.TwainBandwidthPercent.State -ne "NotConfigured")
					{
						line 3 "ICA\Bandwidth\TWAIN device redirection bandwidth limit percent - Value: " $Setting.TwainBandwidthPercent.Value
					}
					If($Setting.DesktopWallpaper.State -ne "NotConfigured")
					{
						line 3 "ICA\Desktop UI\Desktop wallpaper - Value: " $Setting.DesktopWallpaper.State
					}
					If($Setting.MenuAnimation.State -ne "NotConfigured")
					{
						line 3 "ICA\Desktop UI\Menu animation - Value: " $Setting.MenuAnimation.State
					}
					If($Setting.WindowContentsVisibleWhileDragging.State -ne "NotConfigured")
					{
						line 3 "ICA\Desktop UI\View window contents while dragging - Value: " $Setting.WindowContentsVisibleWhileDragging.State
					}
					If($Setting.AutoConnectDrives.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Auto connect client drives - Value: " $Setting.AutoConnectDrives.State
					}
					If($Setting.ClientDriveRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Client drive redirection - Value: " $Setting.ClientDriveRedirection.State
					}
					If($Setting.ClientFixedDrives.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Client fixed drives - Value: " $Setting.ClientFixedDrives.State
					}
					If($Setting.ClientFloppyDrives.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Client floppy drives - Value: " $Setting.ClientFloppyDrives.State
					}
					If($Setting.ClientNetworkDrives.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Client network drives - Value: " $Setting.ClientNetworkDrives.State
					}
					If($Setting.ClientOpticalDrives.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Client optical drives - Value: " $Setting.ClientOpticalDrives.State
					}
					If($Setting.ClientRemoveableDrives.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Client removable drives - Value: " $Setting.ClientRemoveableDrives.State
					}
					If($Setting.HostToClientRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Host to client redirection - Value: " $Setting.HostToClientRedirection.State
					}
					If($Setting.ReadOnlyMappedDrive.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Read-only client drive access - Value: " $Setting.ReadOnlyMappedDrive.State
					}
					If($Setting.SpecialFolderRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Special folder redirection - Value: " $Setting.SpecialFolderRedirection.State
					}
					If($Setting.AsynchronousWrites.State -ne "NotConfigured")
					{
						line 3 "ICA\File Redirection\Use asynchronous writes - Value: " $Setting.AsynchronousWrites.State
					}
					If($Setting.MultiStream.State -ne "NotConfigured")
					{
						line 3 "ICA\Multi-Stream Connections\Multi-Stream - Value: " $Setting.MultiStream.State
					}
					If($Setting.ClientComPortsAutoConnection.State -ne "NotConfigured")
					{
						line 3 "ICA\Port Redirection\Auto connect client COM ports - Value: " $Setting.ClientComPortsAutoConnection.State
					}
					If($Setting.ClientLptPortsAutoConnection.State -ne "NotConfigured")
					{
						line 3 "ICA\Port Redirection\Auto connect client LPT ports - Value: " $Setting.ClientLptPortsAutoConnection.State
					}
					If($Setting.ClientComPortRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\Port Redirection\Client COM port redirection - Value: " $Setting.ClientComPortRedirection.State
					}
					If($Setting.ClientLptPortRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\Port Redirection\Client LPT port redirection - Value: " $Setting.ClientLptPortRedirection.State
					}
					If($Setting.ClientPrinterRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Client printer redirection - Value: " $Setting.ClientPrinterRedirection.State
					}
					If($Setting.DefaultClientPrinter.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Default printer - Value: " $Setting.DefaultClientPrinter.Value
					}
					If($Setting.AutoCreationEventLogPreference.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Printer auto-creation event log preference - Value: " $Setting.AutoCreationEventLogPreference.Value
					}
					If($Setting.SessionPrinters.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Session printers - Value: " $Setting.SessionPrinters.State
					}
					If($Setting.WaitForPrintersToBeCreated.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Wait for printers to be created (desktop) - Value: " $Setting.WaitForPrintersToBeCreated.State
					}
					If($Setting.ClientPrinterAutoCreation.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Client Printers\Auto-create client printers - Value: " $Setting.ClientPrinterAutoCreation.Value
					}
					If($Setting.GenericUniversalPrinterAutoCreation.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Client Printers\Auto-create generic universal printer - Value: " $Setting.GenericUniversalPrinterAutoCreation.Value
					}
					If($Setting.ClientPrinterNames.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Client Printers\Client printer names - Value: " $Setting.ClientPrinterNames.Value
					}
					If($Setting.DirectConnectionsToPrintServers.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Client Printers\Direct connections to print servers - Value: " $Setting.DirectConnectionsToPrintServers.State
					}
					If($Setting.PrinterDriverMappings.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Client Printers\Printer driver mapping and compatibility - Value: " $Setting.PrinterDriverMappings.State
					}
					If($Setting.PrinterPropertiesRetention.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Client Printers\Printer properties retention - Value: " $Setting.PrinterPropertiesRetention.Value
					}
					If($Setting.RetainedAndRestoredClientPrinters.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Client Printers\Retained and restored client printers - Value: " $Setting.RetainedAndRestoredClientPrinters.State
					}
					If($Setting.InboxDriverAutoInstallation.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Drivers\Automatic installation of in-box printer drivers - Value: " $Setting.InboxDriverAutoInstallation.State
					}
					If($Setting.UniversalDriverPriority.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Drivers\Universal driver preference - Value: " $Setting.UniversalDriverPriority.Value
					}
					If($Setting.UniversalPrintDriverUsage.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Drivers\Universal print driver usage - Value: " $Setting.UniversalPrintDriverUsage.Value
					}
					If($Setting.EMFProcessingMode.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Universal Printing\Universal printing EMF processing mode - Value: " $Setting.EMFProcessingMode.Value
					}
					If($Setting.ImageCompressionLimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Universal Printing\Universal printing image compression limit - Value: " $Setting.ImageCompressionLimit.Value
					}
					If($Setting.UPDCompressionDefaults.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Universal Printing\Universal printing optimization default - Value: " -NoNewLine
						$Tmp = "`t`t" + $Setting.UPDCompressionDefaults.Value.replace(";","`n`t`t`t`t")
						line 4 $Tmp
						$Tmp = $null
					}
					If($Setting.UniversalPrintingPreviewPreference.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Universal Printing\Universal printing preview preference - Value: " $Setting.UniversalPrintingPreviewPreference.Value
					}
					If($Setting.DPILimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Printing\Universal Printing\Universal printing print quality limit - Value: " $Setting.DPILimit.Value
					}
					If($Setting.MinimumEncryptionLevel.State -ne "NotConfigured")
					{
						line 3 "ICA\Security\SecureICA minimum encryption level - Value: " $Setting.MinimumEncryptionLevel.Value
					}
					If($Setting.ConcurrentLogOnLimit.State -ne "NotConfigured")
					{
						line 3 "ICA\Session limits\Concurrent logon limit - Value: " $Setting.ConcurrentLogOnLimit.Value
					}
					If($Setting.SessionDisconnectTimer.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Disconnected session timer - Value: " $Setting.SessionDisconnectTimer.State
					}
					If($Setting.SessionDisconnectTimerInterval.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Disconnected session timer interval - Value: " $Setting.SessionDisconnectTimerInterval.Value
					}
					If($Setting.LingerDisconnectTimerInterval.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Linger Disconnect Timer Interval - Value: " $Setting.LingerDisconnectTimerInterval.Value
					}
					If($Setting.LingerTerminateTimerInterval.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Linger Terminate Timer Interval - Value: " $Setting.LingerTerminateTimerInterval.Value
					}
					If($Setting.PrelaunchDisconnectTimerInterval.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Pre-launch Disconnect Timer Interval - Value: " $Setting.PrelaunchDisconnectTimerInterval.Value
					}
					If($Setting.PrelaunchTerminateTimerInterval.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Pre-launch Terminate Timer Interval - Value: " $Setting.PrelaunchTerminateTimerInterval.Value
					}
					If($Setting.SessionConnectionTimer.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Session connection timer - Value: " $Setting.SessionConnectionTimer.State
					}
					If($Setting.SessionConnectionTimerInterval.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Session connection timer interval - Value: " $Setting.SessionConnectionTimerInterval.Value
					}
					If($Setting.SessionIdleTimer.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Session idle timer - Value: " $Setting.SessionIdleTimer.State
					}
					If($Setting.SessionIdleTimerInterval.State -ne "NotConfigured")
					{
						line 3 "ICA\Session Limits\Session idle timer interval - Value: " $Setting.SessionIdleTimerInterval.Value
					}
					If($Setting.ShadowInput.State -ne "NotConfigured")
					{
						line 3 "ICA\Shadowing\Input from shadow connections - Value: " $Setting.ShadowInput.State
					}
					If($Setting.ShadowLogging.State -ne "NotConfigured")
					{
						line 3 "ICA\Shadowing\Log shadow attempts - Value: " $Setting.ShadowLogging.State
					}
					If($Setting.ShadowUserNotification.State -ne "NotConfigured")
					{
						line 3 "ICA\Shadowing\Notify user of pending shadow connections - Value: " $Setting.ShadowUserNotification.State
					}
					If($Setting.ShadowAllowList.State -ne "NotConfigured")
					{
						line 3 "ICA\Shadowing\Users who can shadow other users - Value: " $Setting.ShadowAllowList.Value
					}
					If($Setting.ShadowDenyList.State -ne "NotConfigured")
					{
						line 3 "ICA\Shadowing\Users who cannot shadow other users - Value: " $Setting.ShadowDenyList.Value
					}
					If($Setting.LocalTimeEstimation.State -ne "NotConfigured")
					{
						line 3 "ICA\Time Zone Control\Estimate local time for legacy clients - Value: " $Setting.LocalTimeEstimation.State
					}
					If($Setting.SessionTimeZone.State -ne "NotConfigured")
					{
						line 3 "ICA\Time Zone Control\Use local time of client - Value: " $Setting.SessionTimeZone.Value
					}
					If($Setting.TwainRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\TWAIN devices\Client TWAIN device redirection - Value: " $Setting.TwainRedirection.State
					}
					If($Setting.TwainCompressionLevel.State -ne "NotConfigured")
					{
						line 3 "ICA\TWAIN devices\TWAIN compression level - Value: " $Setting.TwainCompressionLevel.Value
					}
					If($Setting.UsbDeviceRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\USB devices\Client USB device redirection - Value: " $Setting.UsbDeviceRedirection.State
					}
					If($Setting.UsbDeviceRedirectionRules.State -ne "NotConfigured")
					{
						line 3 "ICA\USB devices\Client USB device redirection rules - Value: " $Setting.UsbDeviceRedirectionRules.Value
					}
					If($Setting.UsbPlugAndPlayRedirection.State -ne "NotConfigured")
					{
						line 3 "ICA\USB devices\Client USB Plug and Play device redirection - Value: " $Setting.UsbPlugAndPlayRedirection.State
					}
					If($Setting.FramesPerSecond.State -ne "NotConfigured")
					{
						line 3 "ICA\Visual Display\Max Frames Per Second - Value: " $Setting.FramesPerSecond.Value
					}
					If($Setting.ProgressiveCompressionLevel.State -ne "NotConfigured")
					{
						line 3 "ICA\Visual Display\Moving Images\Progressive compression level - Value: " $Setting.ProgressiveCompressionLevel.Value
					}
					If($Setting.ProgressiveCompressionThreshold.State -ne "NotConfigured")
					{
						line 3 "ICA\Visual Display\Moving Images\Progressive compression threshold value - Value: " $Setting.ProgressiveCompressionThreshold.Value
					}
					If($Setting.ExtraColorCompression.State -ne "NotConfigured")
					{
						line 3 "ICA\Visual Display\Still Images\Extra Color Compression - Value: " $Setting.ExtraColorCompression.State
					}
					If($Setting.ExtraColorCompressionThreshold.State -ne "NotConfigured")
					{
						line 3 "ICA\Visual Display\Still Images\Extra Color Compression Threshold - Value: " $Setting.ExtraColorCompressionThreshold.Value
					}
					If($Setting.ProgressiveHeavyweightCompression.State -ne "NotConfigured")
					{
						line 3 "ICA\Visual Display\Still Images\Heavyweight compression - Value: " $Setting.ProgressiveHeavyweightCompression.State
					}
					If($Setting.LossyCompressionLevel.State -ne "NotConfigured")
					{
						line 3 "ICA\Visual Display\Still Images\Lossy compression level - Value: " $Setting.LossyCompressionLevel.Value
					}
					If($Setting.LossyCompressionThreshold.State -ne "NotConfigured")
					{
						line 3 "ICA\Visual Display\Still Images\Lossy compression threshold value - Value: " $Setting.LossyCompressionThreshold.Value
					}
					If($Setting.SessionImportance.State -ne "NotConfigured")
					{
						line 3 "Server Session Settings\Session importance - Value: " $Setting.SessionImportance.Value
					}
					If($Setting.SingleSignOn.State -ne "NotConfigured")
					{
						line 3 "Server Session Settings\Single Sign-On - Value: " $Setting.SingleSignOn.State
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
	line 0 "Citrix Policy information could not be retrieved.  Was the Citrix.GroupPolicy.Command module imported?"
}
$Policies = $null
$global:output = $null

