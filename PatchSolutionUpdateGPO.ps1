<#
.SYNOPSIS
   Updates and manages the Scheduled Tasks in the given GPO for the Patch Solution
.DESCRIPTION
   Checks the PatchInfo for the patchgroup values and determines if the time and date are within the defined maintenance window. If it is inside the window, it will Windows Update service will be forced to check/install/reboot any updates.
.PARAMETER ComputerName
   The computer name of the WSUS server that hosts the patch solution/scripts
.PARAMETER GroupPolicyName
   Name of the Group Policy to update
.SWITCH OverwritePatchgroups
   Completely overwrite the existing Scheduled Task information. There is no confirmation.
.EXAMPLE
   PatchSolutionUpdateGPOI.ps1 -GroupPolicyName "My PatchSolution GPO"
   Uses the above GPO and updates the scheduled task values in it.
.EXAMPLE
   PatchSolutionUpdateGPOI.ps1 -OverwritePatchgroups
   Removes any existing scheduled tasks and rebuilds them from the config file
#>
<#

Version: 1.0.4
Date: 2014-12-18

Revisions
============

2014-12-18  v1.0.0 Initial Release

#>

param (
	[string]$ComputerName = "wsus",
	[string]$GroupPolicyName = "PatchSolution",
	[switch]$OverwritePatchgroups
)

#################################
### START OF COMMON FUNCTIONS ###
#################################
#Common Code for Patch Solution
. "$PSScriptRoot\PatchFunctions.ps1"

Function New-XmlElement {
    param (
        [xml]$XmlContent,
        [System.Xml.XmlElement]$XMLParent,
        [string]$Element,
        [string]$ElementValue
    )

    $NewElement = $XmlContent.CreateElement($Element)
    $NewElement.InnerText = $ElementValue
    $XMLParent.AppendChild($NewElement) | Out-Null

    Return $XmlContent
}


Function Create-ScheduledTasksXMLHeader {
    param (
        [string]$TaskVersion = "TaskV2"
    )

    $XmlContent = New-Object XML

    $XmlDecl = $XmlContent.CreateXmlDeclaration("1.0", "UTF-8", $null)

    $ScheduledTasksElement = $XmlContent.CreateElement("ScheduledTasks")
    $ScheduledTasksElement.SetAttribute("clsid", "{CC63F200-7309-4ba0-B154-A71CD118DBCC}") | Out-Null
    $XmlContent.InsertBefore($XmlDecl, $XmlContent.DocumentElement) | Out-Null
    $ScheduledTasks = $XmlContent.AppendChild($ScheduledTasksElement)

    Return $XmlContent
}


Function Create-ScheduledTaskXMLNode {
    param (
        [string]$TaskVersion = "TaskV2",
        [string]$Taskname,
        [string]$PatchGroup,
        [string]$ComputerName
    )

    If ($Patchgroup -eq "-1,-1,-1,-1") {
        $ManualTask = $true
    } else {
        $ManualTask = $false
    }


    If ($TaskVersion -eq "TaskV2") {
        $clsid = "{D8896631-B747-47a7-84A6-C155337F3BC8}" 
    } else {
        $clsid = "{D8896631-B747-47a7-84A6-C155337F3BC8}" 
    }

    $uid = "{$([guid]::NewGuid())}"
    $changed = get-date -Format "YYYY-M-d H:M:s"
    $author = "Author"

    $XmlContent = New-Object XML

    $TaskNode = $XmlContent.CreateElement($TaskVersion)
    $TaskNode.SetAttribute("clsid", $clsid) | Out-Null
    $TaskNode.SetAttribute("removePolicy", 1) | Out-Null
    $TaskNode.SetAttribute("userContext", 0) | Out-Null
    $TaskNode.SetAttribute("uid", $uid) | Out-Null
    $TaskNode.SetAttribute("image", 1) | Out-Null
    $TaskNode.SetAttribute("name", $TaskName) | Out-Null
    $TaskNode.SetAttribute("desc", (Get-PatchGroupDescription $PatchGroup)) | Out-Null
    $XmlContent.AppendChild($TaskNode) | Out-Null

    $TaskNode = $XmlContent.CreateElement("Properties")
    $TaskNode.SetAttribute("name", $TaskName) | Out-Null
    $TaskNode.SetAttribute("logonType", "S4U") | Out-Null
    $TaskNode.SetAttribute("runAs", "NT AUTHORITY\SYSTEM") | Out-Null
    $TaskNode.SetAttribute("action", "R") | Out-Null
    $XmlContent.$TaskVersion.AppendChild($TaskNode) | Out-Null

    # Task Info
    $TaskNode = $XmlContent.CreateElement("Task")
    $TaskNode.SetAttribute("version", "1.2") | Out-Null
    $Task = $XmlContent.$TaskVersion.Properties.AppendChild($TaskNode)
    
    # Registraion Info
    $TaskNode = $XmlContent.CreateElement("RegistrationInfo")
    $RegistrationInfo = $Task.AppendChild($TaskNode)

    $XmlContent = New-XmlElement $XmlContent $RegistrationInfo "Author" $author
    $XmlContent = New-XmlElement $XmlContent $RegistrationInfo "Description" "Built by $(Split-Path $MyInvocation.ScriptName -Leaf)"


    # Principals Info
    $TaskNode = $XmlContent.CreateElement("Principals")
    $Principals = $XmlContent.$TaskVersion.Properties.Task.AppendChild($TaskNode) 

        # Principal Info
        $TaskNode = $XmlContent.CreateElement("Principal")
        $TaskNode.SetAttribute("id", $author) | Out-Null
        $Principal = $Principals.AppendChild($TaskNode)

        $XmlContent = New-XmlElement $XmlContent $Principal "RunLevel" "HighestAvailable"
        $XmlContent = New-XmlElement $XmlContent $Principal "UserId" "NT AUTHORITY\SYSTEM"
        $XmlContent = New-XmlElement $XmlContent $Principal "LogonType" "S4U"

    # Settings Info
    $TaskNode = $XmlContent.CreateElement("Settings")
    $Settings = $XmlContent.$TaskVersion.Properties.Task.AppendChild($TaskNode) 

        # Idle Settings
        $TaskNode = $XmlContent.CreateElement("IdleSettings")
        $IdleSettings = $Settings.AppendChild($TaskNode)

        $XmlContent = New-XmlElement $XmlContent $IdleSettings "Duration" "PT10M"
        $XmlContent = New-XmlElement $XmlContent $IdleSettings "WaitTimeout" "PT1H"
        $XmlContent = New-XmlElement $XmlContent $IdleSettings "StopOnIdleEnd" "true"
        $XmlContent = New-XmlElement $XmlContent $IdleSettings "RestartOnIdle" "false"

    $XmlContent = New-XmlElement $XmlContent $Settings "MultipleInstancesPolicy" "IgnoreNew"
    $XmlContent = New-XmlElement $XmlContent $Settings "DisallowStartIfOnBatteries" "false"
    $XmlContent = New-XmlElement $XmlContent $Settings "StopIfGoingOnBatteries" "false"
    $XmlContent = New-XmlElement $XmlContent $Settings "AllowHardTerminate" "true"
    $XmlContent = New-XmlElement $XmlContent $Settings "StartWhenAvailable" "true"
    $XmlContent = New-XmlElement $XmlContent $Settings "AllowStartOnDemand" "true"
    $XmlContent = New-XmlElement $XmlContent $Settings "Enabled" "true"
    $XmlContent = New-XmlElement $XmlContent $Settings "Hidden" "false"
    $XmlContent = New-XmlElement $XmlContent $Settings "ExecutionTimeLimit" "PT8H"
    $XmlContent = New-XmlElement $XmlContent $Settings "Priority" "8"

    # Actions Info
    $TaskNode = $XmlContent.CreateElement("Actions")
    $TaskNode.SetAttribute("Context", $author) | Out-Null
    $Actions = $Task.AppendChild($TaskNode)

    $ExecNode = $XmlContent.CreateElement("Exec")
    $Exec = $Actions.AppendChild($ExecNode) 

    $XmlContent = New-XmlElement $XmlContent $Exec "Command" "PowerShell.exe"
    If ($ManualTask -eq $false) {
        $XmlContent = New-XmlElement $XmlContent $Exec "Arguments" "-ExecutionPolicy Bypass -File \\$ComputerName\WSUS\InstallUpdates-Cryptic.ps1"
    } else {
        $XmlContent = New-XmlElement $XmlContent $Exec "Arguments" "-ExecutionPolicy Bypass -File \\$ComputerName\WSUS\InstallUpdates-Cryptic.ps1 -SkipPatchGroupCheck"
    }
    $XmlContent = New-XmlElement $XmlContent $Exec "WorkingDirectory" "C:\Windows\System32"

    If ($Patchgroup -ne "-1,-1,-1,-1") {

        # Triggers
        $TriggersNode = $XmlContent.CreateElement("Triggers")
        $Triggers = $Task.AppendChild($TriggersNode)

            # CalendarTrigger
            $CalendarNode = $XmlContent.CreateElement("CalendarTrigger")
            $CalenderTrigger1 = $Triggers.AppendChild($CalendarNode)
            $PatchDay, $PatchHour, $PatchMinute, $PatchInterval = $Patchgroup.split(",")
            $XmlContent = New-XmlElement $XmlContent $CalenderTrigger1 "StartBoundary"  "$(Get-Date -Format yyyy-M-d)T$($PatchHour):$($PatchMinute):00"
            $XmlContent = New-XmlElement $XmlContent $CalenderTrigger1 "Enabled" "true"
    
            $Weekdays = (Convert-PatchGroupToWeekday $PatchGroup $false).Split(" ")

                # Schedule By ---- Daily or Weekly
                If ($PatchDay -eq 127) {
                    $ScheduleEveryDayNode = $XmlContent.CreateElement("ScheduleByDay")
                    $ScheduleEveryDay = $CalenderTrigger1.AppendChild($ScheduleEveryDayNode)
                    $XmlContent = New-XmlElement $XmlContent $ScheduleEveryDay "DaysInterval" "1"
                } ElseIf ($Weekday -notlike "*Manual*") {
                    $ScheduleByWeekNode = $XmlContent.CreateElement("ScheduleByWeek")
                    $ScheduleByWeek = $CalenderTrigger1.AppendChild($ScheduleByWeekNode)
                    $XmlContent = New-XmlElement $XmlContent $ScheduleByWeek "WeeksInterval" "1"

                        $DaysOfWeekNode = $XmlContent.CreateElement("DaysOfWeek")
                        $DaysOfWeek = $ScheduleByWeek.AppendChild($DaysOfWeekNode)
                        ForEach ($Weekday in $Weekdays) {
                            $DayOfWeekNode = $XmlContent.CreateElement($Weekday)
                            $DayOfWeek = $DaysOfWeek.AppendChild($DayOfWeekNode)
                        }
                }

            $XmlContent = New-XmlElement $XmlContent $CalenderTrigger1 "ExecutionTimeLimit" "PT4H"

                # Repetition
                $RepetitionNode = $XmlContent.CreateElement("Repetition")
                $Repetition = $CalenderTrigger1.AppendChild($RepetitionNode)
                $XmlContent = New-XmlElement $XmlContent $RepetitionNode "Interval" "PT1H"
                $XmlContent = New-XmlElement $XmlContent $RepetitionNode "Duration" "P1D"
                $XmlContent = New-XmlElement $XmlContent $RepetitionNode "StopAtDurationEnd" "false"

            #Reboot Trigger
            $BootTriggerNode = $XmlContent.CreateElement("BootTrigger")
            $BootTrigger = $Triggers.AppendChild($BootTriggerNode)
            $XmlContent = New-XmlElement $XmlContent $BootTrigger "Enabled"  "true"
            $XmlContent = New-XmlElement $XmlContent $BootTrigger "Delay" "PT1M"
            $XmlContent = New-XmlElement $XmlContent $BootTrigger "ExecutionTimeLimit" "PT2H"


        # Filters
        $FiltersNode = $XmlContent.CreateElement("Filters")
        $Filters = $XmlContent.$TaskVersion.AppendChild($FiltersNode)

        $FilterCollection = $XmlContent.CreateElement("FilterCollection")
        $FilterCollection.SetAttribute("not", "0") | Out-Null
        $FilterCollection.SetAttribute("bool", "AND") | Out-Null
        $Filters = $Filters.AppendChild($FilterCollection)

        $FilterLdap  = $XmlContent.CreateElement("FilterLdap")
        $FilterLdap.SetAttribute("not", "0") | Out-Null
        $FilterLdap.SetAttribute("bool", "AND") | Out-Null
        $FilterLdap.SetAttribute("attribute", "ExtensionAttribute1") | Out-Null
        $FilterLdap.SetAttribute("variableName", "Patchgroup") | Out-Null
        $FilterLdap.SetAttribute("binding", "LDAP:") | Out-Null
        $FilterLdap.SetAttribute("searchFilter", "(&(objectCategory=computer)(objectClass=computer)(cn=%ComputerName%)(ExtensionAttribute1=$Patchgroup))") | Out-Null
        $Filters = $Filters.AppendChild($FilterLdap)
    }
    Return $XmlContent
}


Function Main {
    # Main Program
    if ((Get-Config) -eq $false) {
        Write-Error "Could not read the configuration file. Exiting."
        return 1
    }

    $ComputerName = Get-ScriptParameter("ComputerName")
    $GroupPolicyName = Get-ScriptParameter("GroupPolicyName")

#    $GroupPolicyName = "Corporate - WSUS Patching"

    Write-Host -ForegroundColor Green "Using GPO: $GroupPolicyName"
    $GroupPolicyScheduledTaskVersion = "TaskV2"

    $PatchingGpo = Get-GPO -Name $GroupPolicyName -ErrorAction SilentlyContinue
    If ($PatchingGpo -eq $null) {
        Exit 1
    }

    $ScheduledTaskGPPDir = "\\$($PatchingGpo.DomainName)\SYSVOL\$($PatchingGpo.DomainName)\Policies\{$($PatchingGpo.Id)}\Machine\Preferences\ScheduledTasks\"
    $ScheduledTaskGPPFile = "$ScheduledTaskGPPDir\ScheduledTasks.xml"
    if ((Test-Path $ScheduledTaskGPPDir) -eq $false) {
        New-Item -ItemType Directory -Path $ScheduledTaskGPPDir | Out-Null
    }


    # Read in ScheduledTask.xml File
    if ((Test-Path $ScheduledTaskGPPFile) -eq $false) {
        Write-Warning "Could not find GPO ScheduledTasks.xml file. It will be created."
        $GPPScheduledTaskContent = Create-ScheduledTasksXMLHeader $GroupPolicyScheduledTaskVersion
    } elseif ($OverwritePatchgroups.IsPresent) {
        Write-Warning "Overwriting Existing Group Policy Setings"
        $GPPScheduledTaskContent = Create-ScheduledTasksXMLHeader $GroupPolicyScheduledTaskVersion
    } else {
        $GPPScheduledTaskContent = [xml](Get-Content $ScheduledTaskGPPFile)
    }

    # Remove Scheduled Tasks from the GPP if they are not configured in the config file
    Foreach ($TaskName in $GPPScheduledTaskContent.ScheduledTasks.TaskV2.Name) {
        $PatchGroup = $TaskName.Split("-")[$TaskName.Split("-").Count - 1]
        If ((Test-IsValidPatchGroup $PatchGroup) -eq $false) {
            Write-Host "Removing Scheduled Task From GPP - $TaskName No Longer Exists in the Config file"
            $GPPScheduledTaskContent.ScheduledTasks.$GroupPolicyScheduledTaskVersion | Where {$_.Name -eq $TaskName} | % { $GPPScheduledTaskContent.ScheduledTasks.RemoveChild($_) } | Out-Null
        }
    }

    # Add Any new Scheduled Task Nodes for new Patchgroups
    $Patchgroups = $Script:Config.ValidPatchGroups.PatchGroup
    
    # Add one more to signify creating a manual task
    $Patchgroups += "-1,-1,-1,-1"
    Foreach ($PatchGroup in $Patchgroups) {
        $PatchDay, $PatchHour, $PatchMinute, $PatchInterval = $Patchgroup.split(",")
        # Skip over any manual patchgroups
        If ($PatchDay -eq 128) {
            continue
        }

        If ($PatchGroup -eq "-1,-1,-1,-1") {
            $TaskName = "WSUS-Manual"
        } else {
            $TaskName = "WSUS-Patchgroup-$($PatchGroup)"
        }
        if (($GPPScheduledTaskContent.ScheduledTasks.TaskV2)) {
            If ($GPPScheduledTaskContent.ScheduledTasks.TaskV2.Name.Contains($TaskName) -eq $true) {
                write-host -Foreground Yellow "Skipping Scheduled Task: $TaskName already exists"
                continue
            } else {
                write-host "Creating Scheduled Task: $taskname"
            }
        }

        $PatchScript = "\\$ComputerName\WSUS\InstallPatches.ps1"
        $TaskXML = Create-ScheduledTaskXMLNode $GroupPolicyScheduledTaskVersion $TaskName $PatchGroup $ComputerName

        $xml = $GPPScheduledTaskContent.ImportNode(($TaskXML).DocumentElement, $true)
        $GPPScheduledTaskContent.DocumentElement.AppendChild($xml)
    }

    $GPPScheduledTaskContent.Save($ScheduledTaskGPPFile)

}

$rc = Main
