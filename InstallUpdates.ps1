<#
.SYNOPSIS
   When executed the script checks to see if the date and time are within the computers mainenance window, if so, it will install the updates from WSUS
.DESCRIPTION
   Checks the PatchInfo for the patchgroup values and determines if the time and date are within the defined maintenance window. If it is inside the window, it will Windows Update service will be forced to check/install/reboot any updates.
.PARAMETER SkipPatchGroupCheck
   Skips the service window check of Patchgroup and forces installation of any outstanding updates. Rebooting the computer is a possibility."
.PARAMETER LogFile
   Location of where the log file/list of patches is located. Default is C:\PSWindowsUpdate.log
.PARAMETER ShowPatchesOnly
   Does not install any patches. This displays all patches that are available to the system.
.PARAMETER ShowValidPatchGroups
   Displays the valid patch groups, then exits
.EXAMPLE
   InstallUpdates.ps1 -SkippatchGroupCheck
   Force install any pending updates regardless of the date and time
.EXAMPLE
   InstallUpdates.ps1 -LogFile C:\Temp\Updatelog.txt
.EXAMPLE
   InstallUpdates.ps1 -ShowPatchesOnly
   Display any missing patches and exit without installing anything
#>
<#

Version: 1.0.6
Date: 2015-02-25

Revisions
============

2015-02-25  v1.0.6 Added ConfirmEachPatch parameter so that user can choose individual patches to download and install
2015-02-18  v1.0.5 Fixed bug so progress bar works with only one patch. Fixed bug so that SkipPatchCheck/ShowPatchesOnly will force patching regardless of AD search/values.
2014-12-18  v1.0.4 Fixed Obfuscation for how PS1 file are included. Also fixed common function to not use PSScriptRoot as it does not exist on PowerShell v2
2014-12-09  v1.0.1 Updated obfuscating code to fix spacing issue
2014-11-24  v1.0.0 Initial Release

#>


[cmdletbinding(SupportsShouldProcess=$True)]
param (
	[string]$LogFile = "$($ENV:TEMP)\PSWindowsUpdate.log",
    [string]$DisplayPatchGroupDescription = "",
	[switch]$SkipPatchGroupCheck = $false,
    [switch]$ConfirmEachPatch = $false,
    [switch]$ShowValidPatchGroups = $false,
	[switch]$ShowPatchesOnly = $false
)

#################################
### START OF COMMON FUNCTIONS ###
#################################
#Common Code for Patch Solution
. "$PSScriptRoot\PatchFunctions.ps1"

<#
.SYNOPSIS
  Waits for a user to press a key. By default the program waits for a Y or N key. Perfect for the Y/N prompts.
.Description
	Waits for a Y/N combination. Custom keys can be passed in for custom prompts. IE: YNC (Yes/No/Cancel)
.Outputs
	Key-Pressed (Alpha). 		
.Example
	Write-Host "Should we continue? [Y/N]" -NoNewline
	$answer = get-YesNoPrompt "YN"
.Example
	Write-Host "Should we continue? [Yes/No/Cancel]" -NoNewline
	$answer = get-YesNoPrompt "YNC"
.Link
#>
function get-YesNoPrompt {
	param (
		$validKeys = "YN"
	)

	do {
		$x = $Host.UI.RawUI.ReadKey("NoEcho, IncludeKeyDown")
		[string]$keyPressed = $x.Character
	} until ($($validKeys.ToUpper()).Contains($keyPressed.ToUpper()))
	Write-Host $keyPressed.ToUpper()
	return $keyPressed.ToUpper()
}

Function Get-DaysBitmask {
    return @{1="Mon"; 2="Tue"; 4="Wed"; 8="Thu"; 16="Fri"; 32="Sat"; 64="Sun"; 128="Manual"; }
}

function Set-ServerInBlackoutBMC {
	param (
		[string]$startWindow,
		[string]$endWindow
	)
	
    $Logfile = Get-ScriptParameter "LogFile"
	# Convert conventional date time to ALS required Datetime format
	$BlackoutStartWindow = Get-Date $startWindow -format yyyy/M/d-H:mm	
	$BlackoutEndWindow = Get-Date $endWindow -format yyyy/M/d-H:mm	
	$icalStartTime = Get-Date $startWindow -Format yyyyMMddTHHmm00
	$icalrdate = $icalStartTime.SubString(0,8)
	
	# Determine the number of days, hours and minutes the blackout period is
	#P=Use Period Format, DT is # of days, H=Hours, M=Minutes
	$BlackoutDuration = ([DateTime]$endWindow - [DateTime]$startWindow)	
	$BlackoutDuractionText = "P" + $BlackoutDuration.Days + "DT" + $BlackoutDuration.Hours + "H" + $BlackoutDuration.Minutes + "M"

	$username = "p950psn"
	$password = "p950psn" | ConvertTo-SecureString -AsPlainText -Force
	$cred = New-Object System.Management.Automation.PSCredential -argumentlist $username, $password
	
	# Some values are set from http://als.sbcore.net/BEM_Blackout/seb.js Validation Script
	# Values are: icalStartTime, icalDuration, icalrdate, popstart
	# If the VBScript values are not set here, then you cannot remove the server from ALS!
	$postParams = @{
						User=$cred.UserName;
						Comment='Automated Patching';
						serverlist=$Env:COMPUTERNAME;
						StartTime=$BlackoutStartWindow;
						StopTime=$BlackoutEndWindow;
						Action="Add";
						icalStartTime=$icalStartTime;
						icalDuration=$BlackoutDuractionText;
						icalrdate=$icalrdate;
						popstart="1";
					};
	#$postParams.GetEnumerator() | foreach { write-host "`t$($_.Name):`t$($_.Value)"}
	try {
		$webresult = Invoke-WebRequest -UseBasicParsing -Uri http://als.sbcore.net/BEM_Blackout/blackout_list.php -Method POST -Body $postParams -Credential $cred -ErrorAction SilentlyContinue
	} catch {
		$FixPSVersion = @"
dism /online /enable-feature:NetFx2-ServerCore
dism /online /enable-feature:MicrosoftWindowsPowerShell
dism /online /enable-feature:NetFx2-ServerCore-WOW64
start /wait cmd /c "\\fileserver\inst$\Microsoft\Microsoft Select\Server pool\Windows 2008\PowerShellFor2008R2Core\dotNetFx40_Full_x86_x64_SC.exe" /q /norestart
start /wait cmd /c "\\fileserver.fspa.myntet.se\inst$\Microsoft\Microsoft Select\Server pool\Windows 2008\PowerShellFor2008R2Core\Windows6.1-KB2506143-x64.msu" /quiet /forcerestart

-For computers in the dmz use:
start /wait cmd /c "\\wsus-dmz01\WSUS\PowerShellFor2008R2Core\dotNetFx40_Full_x86_x64_SC.exe" /q /norestart
start /wait cmd /c "\\wsus-dmz01\WSUS\PowerShellFor2008R2Core\Windows6.1-KB2506143-x64.msu" /quiet /forcerestart
"@
		Write-Host "Setting computer in ALS Black failed. Only works on Windows 8/2012+"
		Write-Host -ForegroundColor Yellow  "Run the following commands to install PowerShell 3 on Windows 2008 and R2 machines"
		Write-Host -ForegroundColor Yellow $FixPSVersion
		Write-Output "Setting computer in ALS Black failed. Only works on Windows 8/2012+" | Out-File -Append -FilePath $LogFile
		Write-Output "Run the following commands to install PowerShell 3 on Windows 2008 and R2 machines" | Out-File -Append -FilePath $LogFile
		Write-Output $FixPSVersion | Out-File -Append -FilePath $LogFile
	}
}

function Set-ServerInBlackout {
	param (
		[string]$startWindow,
		[string]$endWindow
	)

    #Set-ServerInBlackoutBMC $startWindow $endWindow
}

function Test-ComputerRecentlyInstalled {
	param (
		$NonProductionNetworks = @()
	)

    $NonProductionNetworks = $NonProductionNetworks | Where {$_ -ne ""}
	$RecentlyInstalled = $false
	$ipAddresses = (Get-NetIPAddress).IPAddress
	foreach ($ipAddress in $ipAddresses) {
		foreach ($Network in $NonProductionNetworks) {
			if ($ipAddress -match $Network) {
				Write-Host -ForegroundColor Green "Found an IP Address on an Non Production Network: $ipAddress"
				$RecentlyInstalled = $true
				break
			}
		}
		if ($RecentlyInstalled -eq $true) {
			break
		}
	}
	
	if ($RecentlyInstalled -eq $false) {
		$InstallDate = ([WMI]'').ConvertToDateTime((Get-WmiObject Win32_OperatingSystem).InstallDate)
		$TimeSpanInfo = New-TimeSpan ($InstallDate) -End (Get-Date)		
        If ([int]$Script:Config.NonProduction.ForcePatchInstallComputerAge -gt 0) {
            $ForcePatchInstallComputerAge = [int]$Script:Config.NonProduction.ForcePatchInstallComputerAge
        } else {
            $ForcePatchInstallComputerAge = 7
        }

		if ($TimeSpanInfo.Days -le $ForcePatchInstallComputerAge) {
			Write-Host -ForegroundColor Green "Computer was installed in the last $ForcePatchInstallComputerAge days"
			$RecentlyInstalled = $true
		}
	}
		
	return $RecentlyInstalled
}

Function Get-WURebootStatus {
	[CmdletBinding (
    	SupportsShouldProcess=$True,
        ConfirmImpact="Low"
    )]
    Param (
		[Boolean]$Silent = $False,
		[Boolean]$AutoReboot = $False
	)
	
	Begin {}
	
	Process	{
        $ComputerName = "localhost"
		$objSystemInfo= New-Object -ComObject "Microsoft.Update.SystemInfo"
		$RebootRequired = $objSystemInfo.RebootRequired
		Switch($RebootRequired)	{
			$true	{
				If($Silent)	{
					Return $true
				} Else {
					if($AutoReboot -ne $true) {
						$Reboot = Read-Host "$($ComputerName): Reboot is required. Do it now ? [Y/N]"
					} Else {
						$Reboot = "Y"
					} 
							
					If($Reboot -eq "Y") {
						Write-Verbose "Rebooting $($ComputerName)"
						Restart-Computer -ComputerName $ComputerName -Force
					}
				}
			}
						
			$false	{ 
				If($Silent) {
					Return $false
				} Else {
					Write-Output "$($ComputerName): Reboot is not Required."
				}
			}
		}
	}
	
	End {}				
}

Function Get-WUList {
	Begin {}

	Process	{
		$objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" #Support local instance only
		$objSession = New-Object -ComObject "Microsoft.Update.Session" #Support local instance only			
		$objSearcher = $objSession.CreateUpdateSearcher()
				
		Try	{
			$search = "IsInstalled = 0"	
			$objResults = $objSearcher.Search($search)
		} Catch {
			If($_ -match "HRESULT: 0x80072EE2") {
				Write-Warning "Probably you don't have connection to Windows Update server"
			} #End If $_ -match "HRESULT: 0x80072EE2"
			Return
		}

		$UpdateCollection = @()
		$UpdateCounter = 0
		if (($objResults.Updates).Count -eq $null) {
			$UpdateCount = 1			
		} else {
			$UpdateCount = ($objResults.Updates).Count
		}
		Foreach($Update in $objResults.Updates) {	
			$UpdateCounter++
			Switch($Update.MaxDownloadSize)	{
				{[System.Math]::Round($_/1KB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1KB,0))+" KB"; break }
				{[System.Math]::Round($_/1MB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1MB,0))+" MB"; break }  
				{[System.Math]::Round($_/1GB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1GB,0))+" GB"; break }    
				{[System.Math]::Round($_/1TB,0) -lt 1024} { $size = [String]([System.Math]::Round($_/1TB,0))+" TB"; break }
				default { $size = $_+"B" }
			}
			
			If($Update.KBArticleIDs -ne "") {
				$KB = "KB"+$Update.KBArticleIDs
			} Else {
				$KB = ""
			}
				
			$Status = ""
			If($Update.IsDownloaded)    {$Status += "D"} else {$status += "-"}
			If($Update.IsInstalled)     {$Status += "I"} else {$status += "-"}
			If($Update.IsMandatory)     {$Status += "M"} else {$status += "-"}
			If($Update.IsHidden)        {$Status += "H"} else {$status += "-"}
			If($Update.IsUninstallable) {$Status += "U"} else {$status += "-"}
			If($Update.IsBeta)          {$Status += "B"} else {$status += "-"} 
	
			Add-Member -InputObject $Update -MemberType NoteProperty -Name ComputerName -Value $env:COMPUTERNAME
			Add-Member -InputObject $Update -MemberType NoteProperty -Name KB -Value $KB
			Add-Member -InputObject $Update -MemberType NoteProperty -Name Size -Value $size
			Add-Member -InputObject $Update -MemberType NoteProperty -Name Status -Value $Status
			Add-Member -InputObject $Update -MemberType NoteProperty -Name X -Value 1
			$Update.PSTypeNames.Clear()
			$Update.PSTypeNames.Add('PSWindowsUpdate.WUList')
			if ($ConfirmEachPatch -eq $false) {
				$UpdateCollection += $Update
			} else {
				Write-Host "[$($UpdateCounter)/$($UpdateCount)] Download and install $($Update.Title) - $($Update.Size)? [Y/N]" -NoNewline
				$answer = get-YesNoPrompt "YN"
				if ($answer -eq "Y") {
					$UpdateCollection += $Update
				}
			}
        }

        Return $UpdateCollection
	}
	
	End{}		
}

Function Get-WUInstall {
	Param (
#		[parameter(ValueFromPipelineByPropertyName=$true)]
		$updates,
		[Boolean]$AcceptAll = $false,
		[Boolean]$AutoReboot = $false
	)

	Begin {}

	Process	{
		$objSystemInfo = New-Object -ComObject "Microsoft.Update.SystemInfo"	
		If($objSystemInfo.RebootRequired) {
			Write-Warning "Reboot is required to continue"
			If($AutoReboot) {
				Restart-Computer -Force
			} 
		} 

		If($Updates.Count -eq 0) {
			Return
		} 
				
		$objServiceManager = New-Object -ComObject "Microsoft.Update.ServiceManager" 
		$objSession = New-Object -ComObject "Microsoft.Update.Session" 		
        
        #############################
		# Accept any EULAs
        #############################        
        $UpdateCounter = 1
		$logCollection = @()
        $updatecount = $updates.count
		if ($updatecount -eq $null) {
			$updatecount = 1
		}
		
		Foreach ($Update in $updates) {	
    		Write-Progress -Activity "[1/3] Preparing updates" -Status "[$UpdateCounter/$updatecount] $($Update.Title) $size" -PercentComplete ([int]($UpdateCounter/$updatecount * 100))
			If($Update.EulaAccepted -eq 0) { 
				$Update.AcceptEula()
			}
							
			$log = New-Object PSObject -Property @{
				Title = $Update.Title
				KB = $Update.KB
				Size = $Update.Size
				Status = $Update.Status
				Stage = "Preparing"
			}
				
			$logCollection += $log
            $UpdateCounter++
		}
		Write-Progress -Activity "[1/3] Preparing updates" -Status "Completed" -Completed
        $logCollection

        #############################        
        # Download updates
        #############################

		$objCollectionDownload = New-Object -ComObject "Microsoft.Update.UpdateColl" 
		$UpdateCounter = 1
        Foreach($Update in $Updates) {
			Write-Progress -Activity "[2/3] Downloading updates" -Status "[$UpdateCounter/$updatecount] $($Update.Title) $size" -PercentComplete ([int]($UpdateCounter/$updatecount * 100))
			Write-Host "Downloading: $($Update.Title) - $($Update.Size)"

			$objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
			$objCollectionTmp.Add($Update) | Out-Null
					
			$Downloader = $objSession.CreateUpdateDownloader() 
			$Downloader.Updates = $objCollectionTmp
			Try {
				$DownloadResult = $Downloader.Download()
			} Catch {
				If($_ -match "HRESULT: 0x80240044") {
					Write-Warning "Security policy does not allow a non-administator to perform this task"
				}					
				Return
			}				
			Switch -Exact ($DownloadResult.ResultCode) {
				0   { $Status = "NotStarted" }
				1   { $Status = "InProgress" }
				2   { $Status = "Downloaded" }
				3   { $Status = "DownloadedWithErrors" }
				4   { $Status = "Failed" }
				5   { $Status = "Aborted" }
			}
			$log = New-Object PSObject -Property @{
				Title = $Update.Title
				KB = $Update.KB
				Size = $Update.Size
				Status = $Status
				Stage = "Downloading"
			}								
			$log

			If($DownloadResult.ResultCode -eq 2) {
				$objCollectionDownload.Add($Update) | Out-Null
			}
				
            $UpdateCounter++
		}		
		Write-Progress -Activity "[3/$NumberOfStage] Downloading updates" -Status "Completed" -Completed

        #############################
        # Install the Update
        #############################
		$NeedsReboot = $false
        $UpdateCounter = 1
		Foreach($Update in $objCollectionDownload) {   
				Write-Progress -Activity "[3/3] Installing updates" -Status "[$UpdateCounter/$updatecount] $($Update.Title)" -PercentComplete ([int]($UpdateCounter/$updatecount * 100))
					
				$objCollectionTmp = New-Object -ComObject "Microsoft.Update.UpdateColl"
				$objCollectionTmp.Add($Update) | Out-Null
					
				$objInstaller = $objSession.CreateUpdateInstaller()
				$objInstaller.Updates = $objCollectionTmp
						
				Try {
                    $InstallResult = $objInstaller.Install()
				} Catch {
					If($_ -match "HRESULT: 0x80240044") {
						Write-Warning "Your security policy don't allow a non-administator identity to perform this task"
					}						
					Return
				} 
					
				If (!$NeedsReboot) { 
					$NeedsReboot = $installResult.RebootRequired 
				}
					
				Switch -exact ($InstallResult.ResultCode) {
					0   { $Status = "NotStarted"}
					1   { $Status = "InProgress"}
					2   { $Status = "Installed"}
					3   { $Status = "InstalledWithErrors"}
					4   { $Status = "Failed"}
					5   { $Status = "Aborted"}
				}
				   
				$log = New-Object PSObject -Property @{
					Title = $Update.Title
					KB = $Update.KB
					Size = $Update.Size
					Status = $Status
					Stage = "Installation"
				} #End PSObject Property					
				$log
			
				$UpdateCounter++
        }
		Write-Progress -Activity "[3/3] Installing updates" -Status "Completed" -Completed
				
		If($NeedsReboot) {
			If($AutoReboot) {
				Restart-Computer -Force
			} Else {
				Return "Reboot is required, Please do this as soon as possible."
			}
		}
	}
	
	End{}		
}


function Get-PatchInfo {
    param (
        [string]$Source = "AD"
    )

    If ($Source -eq "Registry") {
        $SoftwareCo = $Script:Config.PatchGroupRegistrySoftwareKey

	    If (Test-Path "HKLM:\Software\Wow6432Node\$SoftwareCo\PatchInfo") {
		    $PatchInfoKey = Get-Item "HKLM:\Software\Wow6432Node\$SoftwareCo\PatchInfo"
	    } elseIf (Test-Path "HKLM:\Software\$SoftwareCo\PatchInfo") {
		    $PatchInfoKey = Get-Item "HKLM:\Software\$SoftwareCo\PatchInfo"
	    } elseif ($NonProductionDomain -eq $true) {
		    Write-Host -ForegroundColor Yellow "Could not find $SoftwareCo\PatchInfo registry key. This is a Test machine, using PatchGroup 173035"
		    $PatchInfoKey = ""
	    } else {
		    Write-Host -ForegroundColor Red "Could not find $SoftwareCo\PatchInfo registry key"
		    exit 2
	    }

	    if ($PatchInfoKey -ne "") {
		    $PatchDay = (Get-ItemProperty -Path $PatchInfoKey.PSPath -Name PatchDay -ErrorAction SilentlyContinue).PatchDay
		    $PatchHour = (Get-ItemProperty -Path $PatchInfoKey.PSPath -Name PatchTime -ErrorAction SilentlyContinue).PatchHour
		    $PatchMinute = (Get-ItemProperty -Path $PatchInfoKey.PSPath -Name PatchTime -ErrorAction SilentlyContinue).PatchMinute
		    $PatchInterval = (Get-ItemProperty -Path $PatchInfoKey.PSPath -Name PatchInterval -ErrorAction SilentlyContinue).PatchInterval
            $PatchInfo = "$($PatchDay),$($PatchHour),$($PatchMinute),$($PatchInterval)"
	    }	
    } ElseIf ($Source -eq "AD") {

        $strFilter = "(&(objectCategory=Computer)(Name=$($Env:Computername)))"

        $objDomain = New-Object System.DirectoryServices.DirectoryEntry
        $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
        $objSearcher.SearchRoot = $objDomain
        $objSearcher.PageSize = 5
        $objSearcher.Filter = $strFilter
        $objSearcher.SearchScope = "Subtree"

        $colProplist = @("Name", $Script:Config.PatchGroupADAttribute)
        foreach ($i in $colPropList){
            $objSearcher.PropertiesToLoad.Add($i) | Out-Null
        }

        $colResults = $objSearcher.FindAll()

        foreach ($objResult in $colResults) {
            $objItem = $objResult.Properties
            if ($objItem[$Script:Config.PatchGroupADAttribute] -ne $null) {
                $PatchInfo = [string]$objItem[$Script:Config.PatchGroupADAttribute]
            } else {
                $PatchInfo = ""
            }
        }        
        #$PatchInfo = (Get-ADComputer -Filter "Name -eq '$Env:Computername'" -Properties $Script:Config.PatchGroupADAttribute).($Script:Config.PatchGroupADAttribute)
        #if ($PatchInfo -eq $null) {
        #    $PatchInfo = ""
        #}
    }
    return $PatchInfo
}

function main {
    if ((Get-Config) -eq $false) {
        Write-Error "Could not read the configuration file. Exiting."
        return 1
    }

    $Logfile = Get-ScriptParameter "LogFile"
    $NonProductionDomains = $Script:config.NonProduction.Domain
    $NonProductionNetworks = $Script:config.NonProduction.Network
    $SkipPatchGroupCheck = Get-ScriptParameter "SkipPatchGroupCheck"
    $ShowPatchesOnly = Get-ScriptParameter "ShowPatchesOnly"
    $ShowValidPatchGroups = Get-ScriptParameter "ShowValidPatchGroups"

    Write-Host "Logfile: $Logfile"

    if ($ShowValidPatchGroups -eq $true) {
        $ValidPatchGroups = Get-ValidPatchGroups
        Write-Host (Get-PatchGroupDescription $ValidPatchGroups $true)
        return 0
    }

    # Check to make sure script is run with Elevated Privleges
    $User = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Role = (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if ((!$Role) -and ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name -ne "NT AUTHORITY\NETWORK SERVICE")) {
	    Write-Warning "To perform some operations you must run an elevated Windows PowerShell console/NETWORK Service Account."	
        return 3
    }

    #
    # Checkpoint - Are we in a service window?
    #
    $NonProductionComputer = Test-ComputerIsNonProduction
    if (($SkipPatchGroupCheck -eq $false) -and ($ShowPatchesOnly -eq $false) -and ((Test-ComputerRecentlyInstalled $NonProductionNetworks) -eq $false)) {		
		$PatchInfo = Get-PatchInfo -Source $Script:Config.PatchGroupLocation
		$PatchDay, $PatchHour, $PatchMinute, $PatchInterval = $PatchInfo.split(",")

	    if ($NonProductionComputer -eq $true) {
            # Check to see if the patchgroup was set
		    if (($PatchDay -eq $null) -or ($PatchHour -eq $null) -or ($PatchMinute -eq $null) -or ($PatchInterval -eq $null)) {
			    $PatchDay = "127"
			    $PatchHour = "00"
			    $PatchMinute = "30"
			    $PatchInterval = "300"
			    Write-Host -ForegroundColor Yellow "PATCHGROUP is not set. This is an Non Production machine, using PatchGroup $($PatchDay),$($PatchHour),$($PatchMinute),$($PatchInterval)"
		    }
	    } elseif (($PatchDay -eq $null) -or ($PatchHour -eq $null) -or ($PatchMinute -eq $null) -or ($PatchInterval -eq $null)) {
		    Write-Host -ForegroundColor Red "Could not find values for one of the PatchGroup numbers. Exiting"
		    return 5
	    }

        # Check to see if the patchgroup found is valid, if not use a default one
        $ValidPatchGroups = Get-ValidPatchGroups
	    if ((Test-IsValidPatchGroup $PatchInfo) -eq $false) {
		    if ($NonProductionComputer -eq $true) {
			    $PatchDay = "127"
			    $PatchHour = "00"
			    $PatchMinute = "30"
			    $PatchInterval = "300"
			    Write-Host -ForegroundColor Yellow "PATCHGROUP is incorrect. This is an Non Production machine, using PatchGroup $($PatchDay),$($PatchHour),$($PatchMinute),$($PatchInterval)"
		    } else {
			    Write-Host -ForegroundColor Red "Patch group in the AD/Registry/DB is invalid"
			    return 6
		    }
	    }
	
	    # QA/Non Production Computers can patch during change stop, all other computers cannot
	    if (($NonProductionComputer -eq $false) -and ($(Test-ComputerIsNonProduction) -eq $false)) {
            if (Test-IsChangeStop -eq $true) {
                $ChangeStopEnds = Get-EndOfChangeStop
			    Write-Host -ForegroundColor Red "Date is now in change stop. Change stop ends on $ChangeStopEnds. Exiting"
			    return 7
		    }
	    } else {
		    Write-Host -ForegroundColor Yellow "CHANGE STOP - This is a Non Production machine, Patching will continue"			
	    }
	
	    # Some systems use SE or PB languages and not Mon, Tue. Let's convert into english.
	    $culture = New-Object system.globalization.cultureinfo("en-US")
	    $DayOfWeekIndex = get-date -UFormat %u
	    $dayOfWeek = $culture.DateTimeFormat.AbbreviatedDayNames[$DayOfWeekIndex]
	
        $PatchgroupDayOfWeek = Convert-PatchGroupToWeekday "$($PatchDay),$($PatchHour),$($PatchMinute),$($PatchInterval)"
	    if (($PatchgroupDayOfWeek -notcontains "$dayOfWeek") -and ($PatchgroupDayOfWeek -ne "Everyday")) {
            $PatchGroupDescription = Get-PatchGroupDescription "$($PatchDay),$($PatchHour),$($PatchMinute),$($PatchInterval)"
		    Write-Host -ForegroundColor Red "Today is not a patch day. Today is: [$dayOfWeek]. Require [$PatchGroupDescription]"
		    return 8
	    } else {
		    Write-Host "Patch day is today ($DayOfWeek)"
	    }

	    # Determine the start and stop time for the patchgroup
	    if ($PatchHour -ge 30) {
		    $startWindow = Get-Date (($PatchHour - 30).ToString().PadLeft(2, "0") + ":30:00")
	    } else {
		    $startWindow = Get-Date ($PatchHour.ToString().PadLeft(2, "0") + ":00:00")
	    }

	    $startWindow = Get-Date ($PatchHour.ToString().PadLeft(2, "0") + ":" + $PatchMinute.ToString().PadLeft(2, "0") + " :00")
	    $endWindow = (Get-Date $startWindow).AddMinutes([int]$PatchInterval)
	
	    if (((Get-Date) -gt (Get-Date $startWindow)) -and ((Get-Date) -le (Get-Date $endWindow))) {
		    Write-Host "From a time perspective, we are inside the Maintanance Window"
	    } else {
            $PatchGroupDescription = Get-PatchGroupDescription "$($PatchDay),$($PatchHour),$($PatchMinute),$($PatchInterval)"
		    Write-Host -ForegroundColor Red "Outside Maintenance Window. Maintenance Window is $($PatchGroupDescription)"
		    return 9
	    }
    } else {
	    Write-Host -ForegroundColor Green "Skipping PatchGroup/Maintenance Window check"
    }


    #
    # Install Patches
    #
    Write-Host "Fetching Updates known to this computer"
    $updates = Get-WUList -WarningAction SilentlyContinue
    if ($updates -ne $null) {
        $updatecount = $updates.count
		if ($updatecount -eq $null) {
			$updatecount = 1
		}
	    Write-Host "...Found $updatecount Updates"
	    $updates | % {Write-Host "`t$($_.KB)`t$($_.title)"}

	    if ($ShowPatchesOnly -eq $false) {
		    # This start/endWindow code will be used if the caller used the $SkipPatchGroupCheck parameter
		    if ((!$startWindow) -or (!$endWindow)) {
			    $BlackoutStartWindow = Get-Date 
			    $BlackoutEndWindow = Get-Date (Get-Date $BlackoutStartWindow).AddMinutes(120)
			    if ($SkipPatchGroupCheck) {
				    Write-Host -ForegroundColor Yellow "SkipPatchGroupCheck was used, using current time for Blackout start time"
			    } else {
				    Write-Host -ForegroundColor Red "Could not determine start/stop time for the patchgroup, using current time for Blackout start time"
			    }
		    } else {
			    $BlackoutStartWindow = $startWindow
			    $BlackoutEndWindow = $endWindow
		    }
		    Set-ServerInBlackout $BlackoutStartWindow $BlackoutEndWindow
	
		    Write-Host "Starting Installation of Updates"	
		    Write-Output "=======================================" | Out-File -Append -FilePath $LogFile
		    Get-Date | Out-File -Append -FilePath $LogFile
		    Write-Output "Script Caller: $($env:UserDomain)\$($env:Username)" | Out-File -Append -FilePath $LogFile
		
		    Write-Host "Starting Updates"
            $autoreboot = $true 
            $acceptall = $true
		    Get-WUInstall $updates $autoreboot $acceptall -WarningAction SilentlyContinue | Out-File -Append -FilePath $LogFile
	    }
    } else {
	    Write-Host "...Found 0 Updates"
    }
    $silent = $true
    $autoreboot = $false
    if ((Get-WURebootStatus $silent $autoreboot) -eq $true) {
        $autoreboot = $true
	    Get-WURebootStatus $silent $autoreboot
    }
    return 0
}

$rc = main
exit $rc
