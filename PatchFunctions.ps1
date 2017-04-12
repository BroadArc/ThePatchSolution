#################################
### START OF COMMON FUNCTIONS ###
#################################
#Common Code for Patch Solution
#These code blocks need to be copied into each script before they can be obfuscated. Otherwise each script will have different function names/calls
#

#$helperScript = "$PSScriptRoot\PatchFunctions.ps1"
#if ((Test-Path $helperScript) -eq $false) {
#    Write-Error "Could not locate $helperScript helper script."
#    exit 1
#} else {
#    . "$PSScriptRoot\PatchFunctions.ps1"
#}

Function Get-ScriptParameter {
    param (
        [string]$ParamName
    )
    return (Get-Variable | where {$_.Name -eq $ParamName}).Value
}

Function Get-DaysBitmask {
    param (
        $UseShortNames = $true
    )
    If ($UseShortNames) {
        return @{1="Mon"; 2="Tue"; 4="Wed"; 8="Thu"; 16="Fri"; 32="Sat"; 64="Sun"; 128="Manual"; }
    } else {
        return @{1="Monday"; 2="Tuesday"; 4="Wednesday"; 8="Thursday"; 16="Friday"; 32="Saturday"; 64="Sunday"; 128="Manual"; }    
    }
}

Function Get-Config {    
    If (!$PSScriptRoot) {
        $ScriptRoot = Split-Path -Path $Script:MyInvocation.MyCommand.Path -Parent
    } else {
        $ScriptRoot = $PSScriptRoot
    }

    if ($ScriptRoot) {
        $configFile = "$($ScriptRoot)\PatchConfig.xml"
    } else {
        $configFile = "PatchConfig.xml"
    }

    If ((Test-Path $configFile) -eq $false) {
        return $false
    } else {
        # load it into an XML object:
        $xml = New-Object -TypeName XML
        $xml.Load($configFile)

        $Script:config = $xml.Config

        # Let's check a few things
        if ($Script:config.PatchGroupLocation -notmatch ("AD|REGISTRY")) {
            Write-Host "Invalid Config Value - PatchGroupLocation"
            return $false
        } elseif ($Script:config.ValidPatchGroups.Patchgroup.Count -le 0) {
            Write-Host "Invalid Config Value - ValidPatchGroups. No Patchgroup values found."
            return $false
        } elseif (($Script:config.PatchGroupLocation -eq ("AD")) -and ($Script:Config.SelectSingleNode("//PatchGroupADAttribute") -eq $null)) {
            Write-Host "Could not determine which AD Attribute to use to locate the PatchGroup information."
            return $false
        }

        if ($Script:Config.SelectSingleNode("//NonProduction/ForcePatchInstallComputerAge") -eq $null) {
            $Network = $xml.CreateElement("ForcePatchInstallComputerAge")
            $NonProduction = Select-XML -Xml $xml -XPath '//Config/NonProduction'
            $xml.Config.NonProduction.AppendChild($Network)
        }
        if ($Script:Config.SelectSingleNode("//NonProduction/Network") -eq $null) {
            $Network = $xml.CreateElement("Network")
            $NonProduction = Select-XML -Xml $xml -XPath '//Config/NonProduction'
            $xml.Config.NonProduction.AppendChild($Network)
        }
        if ($Script:Config.SelectSingleNode("//NonProduction/Domain") -eq $null) {
            $Domain = $xml.CreateElement("Domain")
            $NonProduction = Select-XML -Xml $xml -XPath '//Config/NonProduction'
            $xml.Config.NonProduction.AppendChild($Domain)
        }
        if ($Script:config.NonProduction.ComputerNamePattern.Count -le 0) {
            $ComputerNamePattern = $xml.CreateElement("ComputerNamePattern")
            $NonProduction = Select-XML -Xml $xml -XPath '//Config/NonProduction'
            $xml.Config.NonProduction.AppendChild($ComputerNamePattern)

        }
        $Script:Config = $xml.Config
    }
    return $true
}

Function Get-ValidPatchGroups {
    # Run the script with -ShowValidPatchGroups to view the valid patchgroups
    return $Script:config.ValidPatchgroups.Patchgroup
}

Function Test-IsValidPatchGroup {
    param (
		[Parameter(Mandatory=$True,
		ValueFromPipeline=$True)]
        [AllowEmptyString()]
        [string]$Patchgroup
    )

    if ($Patchgroup) {
        $PatchDay, $PatchHour, $PatchMinute, $PatchInterval = $Patchgroup.split(",")
        if ($Script:Config.ValidPatchgroups.Patchgroup.contains("$($PatchDay),$($PatchHour),$($PatchMinute),$($PatchInterval)") -eq $false) {
            return $false
        } else {
            return $true
        }
    } else {
        return $false
    }      
}

Function Get-PatchGroupDescription {
    param (
		[Parameter(Mandatory=$True,
		ValueFromPipeline=$True)]
        [AllowEmptyString()]
        [string[]]$Patchgroups,
		[Parameter(Mandatory=$False,
		ValueFromPipeline=$True)]
        [boolean]$DisplayPatchGroup
    )
    if ($DisplayPatchGroup) {
        Write-Host "Patchgroup is made up of 4 octets separated by a ,"
        Write-Host "`tOctet 1 = Week day"
        Write-Host "`tOctet 2 = Starting Hour"
        Write-Host "`tOctet 3 = Starting Minute"
        Write-Host "`tOctet 4 = Duration of Service Window (minutes)"
        Write-Host "Weekdays are made up of bits"
        $DaysBitMask = [collections.sortedlist](Get-DaysBitmask)
        foreach ($Bit in $DaysBitMask.GetEnumerator().Name) {
            write-host "`tBit $Bit = $($DaysBitMask[$Bit])"
        }
        write-host
    }

    $output = @()
    foreach ($Patchgroup in $Patchgroups) {
        if ((Test-IsValidPatchGroup $Patchgroup) -eq $true) {
            $DaysofWeek = Convert-PatchGroupToWeekday $Patchgroup
            $PatchDay, $PatchHour, $PatchMinute, $PatchInterval = $Patchgroup.split(",")
            $PatchInterval = $PatchInterval -replace "^0*",""

            $timespan = New-Timespan -Minutes $PatchInterval
            $Hours = ($timespan.Days * 24) + ($timespan.Hours)

            if ($timespan.Minutes -eq 0) {
                $minutes = ""
            } else {
                $minutes = " and $($timespan.Minutes) minutes"
            }

            if ($DisplayPatchGroup) {
                $PatchGroupValue = " Patchgroup = $Patchgroup"
            }

	        $startWindow = Get-Date ($PatchHour.ToString().PadLeft(2, "0") + ":" + $PatchMinute.ToString().PadLeft(2, "0") + " :00")
	        $endWindow = (Get-Date $startWindow).AddMinutes([int]$PatchInterval) 
            $output += "$DaysOfWeek at $($startWindow.Hour.ToString().PadLeft(2, "0")):$($startWindow.Minute.ToString().PadLeft(2, "0")) - $($endWindow.Hour.ToString().PadLeft(2, "0")):$($endWindow.Minute.ToString().PadLeft(2, "0")) ($($Hours) Hours)$($minutes)$($PatchGroupValue)"
        } else {
            $output += "Invalid/Unknown"
        }
        If ($DisplayPatchGroup) {
            Write-Host $output
            $output = ""
        }
    }
    return $output
}

function Test-ComputerIsNonProduction {
    $IsNonProduction = $false
    $NonProductionDomains = $Script:Config.NonProduction.Domain | where {$_ -ne ""}
    $NonProductionNetworks = $Script:Config.NonProduction.Network | where {$_ -ne ""}
    $NonProductionComputerNamePattern = $Script:Config.NonProduction.ComputerNamePattern | where {$_ -ne ""}

    if ($NonProductionDomains.Contains((gwmi WIN32_ComputerSystem).Domain)) {
        $IsNonProduction = $True
    }

    if ($IsNonProduction -eq $false) {
        $ComputerName = (gwmi WIN32_ComputerSystem).Name
        foreach ($pattern in $NonProductionComputerNamePattern) {
            if ($ComputerName -match $pattern) {
                $IsNonProduction = $True
                break
            }
        }
    }
    
    if ($IsNonProduction -eq $false) {
	    $ipAddresses = (Get-NetIPAddress).IPAddress
	    foreach ($ipAddress in $ipAddresses) {
            foreach ($pattern in $NonProductionNetworks) {
                if ($ipAddress -match $pattern) {
                    $IsNonProduction = $True
                    break
                }
            }
        }        
    }
    return $IsNonProduction
}

function Convert-PatchGroupToWeekday {
    param (
        [string]$Patchgroup,
        $UseShortNames = $true
    )

    $PatchDay = ($Patchgroup.split(","))[0] -replace "^0*","";
    
    $DaysOfWeek = @();
    $DaysBitMask = Get-DaysBitmask $UseShortNames
    if ($PatchDay -eq 127) {
        $DaysOfWeek = @("Everyday")
    } elseif ($PatchDay -ne "") {
        $DaysBitmask.Keys | where { $_ -band $PatchDay } | foreach { $DaysOfWeek += $DaysBitmask.Get_Item($_) }
    }
    return $DaysOfWeek
}

Function Get-PatchTuesday {
    [int]$WeekNumber = 2
    [int]$WeekDay = 2

    $FirstDayOfMonth = Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0
    [int]$FirstDayofMonthDay = $FirstDayOfMonth.DayOfWeek
    $Difference = $WeekDay - $FirstDayofMonthDay
    If ($Difference -lt 0) {
        $DaysToAdd = 7 - ($FirstDayofMonthDay - $WeekDay)
    } elseif ($difference -eq 0 ) {
        $DaysToAdd = 0
    }else {
        $DaysToAdd = $Difference
    }
    
    $FirstWeekDayofMonth = $FirstDayOfMonth.AddDays($DaysToAdd)
    $DaysToAdd = ($WeekNumber -1)*7
    $TheDay = $FirstWeekDayofMonth.AddDays($DaysToAdd)
    If (!($TheDay.Month -eq $FirstDayOfMonth.Month -and $TheDay.Year -eq $FirstDayOfMonth.Year)) {
        $TheDay = $null
    }
    return [DateTime]$TheDay
}

Function Get-PatchTuesdayChangeStopDays {
    if ([int]$Script:Config.ChangeStop.PatchTuesdayChangeStopLength -gt 0) {
        return $Script:Config.ChangeStop.PatchTuesdayChangeStopLength
    } 
    return 10
}

Function Get-DaysLeftInPatchTuesdayChangeStop {
    [int]$dayOfMonth = get-date -UFormat %d #Day of the month - 2 digits (05)
    $PatchTuesday = Get-PatchTuesday

    write-host "Microsoft Patch Tuesday is $PatchTuesday"
    $endOfChangeStop = (Get-PatchTuesdayChangeStopDays + [int]($PatchTuesday.Day)) - $dayOfMonth

	if (($dayofMonth -ge $PatchTuesday.Day) -and ($dayOfMonth -lt ($PatchTuesday.Day) + (Get-PatchTuesdayChangeStopDays))) {
        return $endOfChangeStop
	} else {
        return 0
    }
}

Function Get-EndOfChangeStop {
    $DaysLeftInPatchTuesdayChangeStop = Get-DaysLeftInPatchTuesdayChangeStop
    If ($DaysLeftInPatchTuesdayChangeStop -gt 0) {
        return (Get-Date).AddDays($DaysLeftInPatchTuesdayChangeStop)
    }

    $Today = Get-Date
#    $Today = Get-Date -Month 12 -Day 23 -Year 2014
    
    Foreach ($Changestop in $Script:Config.ChangeStop.Date) {
        [string]$StartDay, [string]$EndDay = $ChangeStop.Split("-")

        if ($StartDay -match "(.*)/(.*)") {
            [int]$startMonth = $matches[1]
            [int]$startDay = $matches[2]
            $FullStartDate = Get-Date -Day $startDay -Month $startMonth -Hour 00 -Minute 00 -Second 00
        } else {
            [int]$startMonth = 0
            [int]$startDay = $startDay
            $FullStartDate = Get-Date -Day $startDay -Hour 00 -Minute 00 -Second 00
        }     

        if ($EndDay -match "(.*)/(.*)") {
            [int]$endMonth = $matches[1]
            [int]$EndDay = $matches[2]
            [int]$EndYear = $Today.Year
            if ($startMonth -gt $endMonth) {
                [int]$EndYear = $Today.Year + 1               
            }                
            $FullEndDate = Get-Date -Day $EndDay -Month $endMonth -Year $EndYear -Hour 00 -Minute 00 -Second 00
        } else {
            [int]$endMonth = 0
            [int]$EndDay = $EndDay

            $EndMonth = $Today.Month
            $EndYear = $Today.Year
            if ($EndDay -lt $StartDay) {
                [int]$EndMonth = $Today.Month + 1
                If ($EndMonth -gt 12) {
                    $EndMonth = $EndMonth - 12
                    [int]$EndYear = $Today.Year + 1
                }
            }  

            $FullEndDate = Get-Date -Day $EndDay -Month $endMonth -Year $EndYear -Hour 00 -Minute 00 -Second 00
        }     

        #write-host "$Changestop - [$startMonth] $StartDay, [$endMonth] $EndDay"
        #write-host "FULL: $FullStartDate, $FullEndDate"
        #write-host

        if (($Today -ge $FullStartDate) -and ($Today -le $FullEndDate)) {
            return $FullEndDate
        } 
    }
    return $false
}

Function Test-IsChangeStop {

    $EndOfChangeStop = Get-EndOfChangeStop
    If ($EndOfChangeStop) {
        Write-Host "Change Stop Ends: $EndOfChangeStop"
        return $true
    } else {
        return $false
    }
}

###############################
### END OF COMMON FUNCTIONS ###
###############################