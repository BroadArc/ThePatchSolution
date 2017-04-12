<#
.SYNOPSIS
   Script connects to the WSUS Server and exports the missing patches and emails a summary report
.DESCRIPTION
   Use a schedule task to do a full report to see what is outstanding and a daily report to see which computers will be patched in the next few days
.PARAMETER ComputerName
   The WSUS Server Computer Name that the script will report from
.PARAMETER UseSSL
   Should the connection to the WSUS Server use SSL?
.PARAMETER Port
   The WSUS Port Number to connect
.PARAMETER MailTo
   An array of email address that will receive the summary report. This overrides the configuration file value.
.PARAMETER FullReport
   The script will report on all computers regardless of their patch groups and upcoming patch days
.PARAMETER SkipExportResults
   Tells the script not to export the results to the XLSX file
.PARAMETER SkipEmailResults
   Tells the script to not email the summary report
.EXAMPLE
   PatchReporting.ps1 -SkipEmailResults
   Will report on only computers that will be patched in the upcoming days and it will only export the results. No email will be sent
.EXAMPLE
   PatchReporting.ps1 -SkipEmailResults
   Will report on only computers that will be patched in the upcoming days and it will only email the results. No exported file will be created
.EXAMPLE
   InstallUpdates.ps1 -Full
   Creates a report for all computers. Similiar to the WSUS built in reports, but this will email and export only outstanding updates
.EXAMPLE
   InstallUpdates.ps1 -Full -SkipExportResults
   Creates a report for all computers. Similiar to the WSUS built in reports, but this will email and export only outstanding updates. No export file will be created.
#>

<#

Version: 1.0.4
Date: 2014-12-18

Revisions
============
2014-12-18  v1.0.4 Fixed Obfuscation for how PS1 file are included. Also fixed common function to not use PSScriptRoot as it does not exist on PowerShell v2
2014-12-09  v1.0.3 Full report now adds -FULL to the XLSX file name. Skip over pulling Update GUID Info if SkipExportResults is used.
2014-12-09  v1.0.2 Added domain name to email subject, updated obfuscating code to fix spacing issue
2014-12-02  v1.0.1 Fixed regex bug for showing computers for the next day
2014-11-24  v1.0.0 Initial Release

#>

param (
	[string]$ComputerName = "wsus",
	[boolean]$UseSSL = $false,
	[int]$Port = 8530,
	[string[]]$MailTo = @(),
	[switch]$FullReport,
	[switch]$SkipExportResults,
	[switch]$SkipEmailResults
)

#################################
### START OF COMMON FUNCTIONS ###
#################################
#Common Code for Patch Solution
. "$PSScriptRoot\PatchFunctions.ps1"

function Get-ReportFilePath {
    $NonProductionComputer = Test-ComputerIsNonProduction
    $Folder = "$PSScriptRoot\Reports\WindowsUpdates"
    $FullReport = Get-ScriptParameter "FullReport"
    If ($FullReport) {
        $AppendFullReport = "-Full"
    } else {
        $AppendFullReport = ""
    }
    if ($NonProductionComputer) {
	    $Folder = "$PSScriptRoot\Reports\WindowsUpdates-NonProd"
    } else {
        $Folder = "$PSScriptRoot\Reports\WindowsUpdates"
    }
    $ReportFile = "$Folder\WindowsPatchReport$($AppendFullReport)-$(get-date -format yyyyMMdd-HHmm).xlsx"
    if ((Test-Path $Folder) -eq $false) {
        New-Item -ItemType Directory -Path $Folder -confirm:$false
    }
    return $ReportFile
}

function connect-sqlServer {
	param (
		[string]$SQLServer = "localhost", #use Server\Instance for named SQL instances
		[string]$SQLDBName = "AdventureWorks",
		[string]$Username = "",
		[string]$Password = ""

	)
	$SqlConnection = New-Object System.Data.SqlClient.SqlConnection
	if ($MyInvocation.ScriptName) {
		$scriptname = $(split-path $MyInvocation.ScriptName -Leaf)
	} else {
		$scriptname = "PowerShell Interactive Shell"
	}
	if ($Username.Trim() -eq "") {
		$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True; Application Name = $scriptname"
	} else {

		$SqlConnection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = False; User ID = $Username; Password = $Password; Application Name = $scriptname"
	}
	$SqlConnection.Open()

	return $SqlConnection
}

function Convert-EnvironmentNumberToText {
	param (
		[int]$EnvironmentNumber
	)
	
	Switch ($EnvironmentNumber) {
		0 {return "Test"} 
		1 {return "Development"} 
		2 {return "Acceptance"} 
		3 {return "Production"} 
		default { return "Unknown" }
	}
}



function Export-XLSX {
#
#
# This code was downloaded from Microsoft. All comments and empty lines have been cleaned up to make this script smaller
# http://gallery.technet.microsoft.com/office/Export-XLSX-PowerShell-f2f0c035
#

# Removed Test-FileLock Function
# Changed Parameter Order in Add-XLSXWorksheet to Name, Path. Original is Path, Name
#
	[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
	param(		
        [System.Management.Automation.PSObject]$InputObjects,
		[System.String]$Path,
		[String]$WorkSheetName,
		[Boolean]$Append = 0
		)
	Begin {
		[Boolean]$NoClobber = $false
		[Boolean]$Force = $false
		[Boolean]$NoHeader = $false

		$Null = [Reflection.Assembly]::LoadWithPartialName("WindowsBase")
		$AssemblyLoaded = $False
		ForEach ($asm in [AppDomain]::CurrentDomain.GetAssemblies()) {
			If ($asm.GetName().Name -eq 'WindowsBase') {
				$AssemblyLoaded = $True
			}
		}
		If(-not $AssemblyLoaded) {
			$message = "Could not load 'WindowsBase.dll' assembly from .NET Framework 3.0!"
			$exception = New-Object System.IO.FileNotFoundException $message
			$errorID = 'AssemblyFileNotFound'
			$errorCategory = [Management.Automation.ErrorCategory]::NotInstalled
			$target = 'C:\Program Files\Reference Assemblies\Microsoft\Framework\v3.0\WindowsBase.dll'
			$errorRecord = New-Object Management.Automation.ErrorRecord $exception,$errorID,$errorCategory,$target
			$PSCmdlet.ThrowTerminatingError($errorRecord)
			return
		}
	
		Function Add-XLSXWorkSheet {
			[CmdletBinding()]
			param(
#				[Parameter(Mandatory=$True,
#					Position=0,
#					ValueFromPipeline=$True,
#					ValueFromPipelinebyPropertyName=$True
#				)]
				[String]$Name,
				[String]$Path
			)
			Begin {
				$New_Worksheet_xml = New-Object System.Xml.XmlDocument
				$XmlDeclaration = $New_Worksheet_xml.CreateXmlDeclaration("1.0", "UTF-8", "yes")
				$Null = $New_Worksheet_xml.InsertBefore($XmlDeclaration, $New_Worksheet_xml.DocumentElement)
				$workSheetElement = $New_Worksheet_xml.CreateElement("worksheet")
				$Null = $workSheetElement.SetAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
				$Null = $workSheetElement.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
				$Null = $New_Worksheet_xml.AppendChild($workSheetElement)
				$Null = $New_Worksheet_xml.DocumentElement.AppendChild($New_Worksheet_xml.CreateElement("sheetData"))
			} 
		Process {
				Try {
					$Null = Get-Item -Path $Path -ErrorAction stop
				} Catch {
					$Error.RemoveAt(0)
					$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
					$PSCmdlet.WriteError($NewError)
					Return
				}
				Try {
					$exPkg = [System.IO.Packaging.Package]::Open($Path, [System.IO.FileMode]::Open)
				} catch {
					$_
					Return
				}
				ForEach ($Part in $exPkg.GetParts()) {
					IF($Part.ContentType -eq "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" -or $Part.Uri.OriginalString -eq "/xl/workbook.xml") {
						$WorkBookPart = $Part
						break
					}
				}
				If(-not $WorkBookPart) {
					Write-Error "Excel Workbook not found in : $Path"
					$exPkg.Close()
					return
				}
				$WorkBookRels = $WorkBookPart.GetRelationships()
				$WorkBookRelIds = [System.Collections.ArrayList]@()
				$WorkSheetPartNames = [System.Collections.ArrayList]@()
				ForEach($Rel in $WorkBookRels) {
					$Null = $WorkBookRelIds.Add($Rel.ID)
					If($Rel.RelationshipType -like '*worksheet*' ) {
						$WorkSheetName = Split-Path $Rel.TargetUri.ToString() -Leaf
						$Null = $WorkSheetPartNames.Add($WorkSheetName)
					}
				}
				$IdCounter = 0 
				$NewWorkBookRelId = '' 
				Do{
					$IdCounter++
					If(-not ($WorkBookRelIds -contains "rId$IdCounter")){
						$NewWorkBookRelId = "rId$IdCounter"
					}
				} while($NewWorkBookRelId -eq '')
				$WorksheetCounter = 0 
				$NewWorkSheetPartName = '' 
				Do{
					$WorksheetCounter++
					If(-not ($WorkSheetPartNames -contains "sheet$WorksheetCounter.xml")){
						$NewWorkSheetPartName = "sheet$WorksheetCounter.xml"
					}
				} while($NewWorkSheetPartName -eq '')
				$WorkbookWorksheetNames = [System.Collections.ArrayList]@()
				$WorkBookXmlDoc = New-Object System.Xml.XmlDocument
				$WorkBookXmlDoc.Load($WorkBookPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read))
				ForEach ($Element in $WorkBookXmlDoc.documentElement.Item("sheets").get_ChildNodes()) {
					$Null = $WorkbookWorksheetNames.Add($Element.Name)
				}
				$DuplicateName = ''
				If(-not [String]::IsNullOrEmpty($Name)){
					If($WorkbookWorksheetNames -Contains $Name) {
						$DuplicateName = $Name
						$Name = ''
					}
				} 
				If([String]::IsNullOrEmpty($Name)){
					$WorkSheetNameCounter = 0
					$Name = "Table$WorkSheetNameCounter"
					While($WorkbookWorksheetNames -Contains $Name) {
						$WorkSheetNameCounter++
						$Name = "Table$WorkSheetNameCounter"
					}
					If(-not [String]::IsNullOrEmpty($DuplicateName)){
						Write-Warning "Worksheetname '$DuplicateName' allready exist!`nUsing automatically generated name: $Name"
					}
				}
		
				$Uri_xl_worksheets_sheet_xml = New-Object System.Uri -ArgumentList ("/xl/worksheets/$NewWorkSheetPartName", [System.UriKind]::Relative)
				$Part_xl_worksheets_sheet_xml = $exPkg.CreatePart($Uri_xl_worksheets_sheet_xml, "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
				$dest = $part_xl_worksheets_sheet_xml.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write)
				$New_Worksheet_xml.Save($dest)
				$Null = $WorkBookPart.CreateRelationship($Uri_xl_worksheets_sheet_xml, [System.IO.Packaging.TargetMode]::Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", $NewWorkBookRelId)
		
		
				$WorkBookXmlDoc = New-Object System.Xml.XmlDocument
				$WorkBookXmlDoc.Load($WorkBookPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read))
				$WorkBookXmlSheetNode = $WorkBookXmlDoc.CreateElement('sheet', $WorkBookXmlDoc.DocumentElement.NamespaceURI)
				$Null = $WorkBookXmlSheetNode.SetAttribute('name',$Name)
				$Null = $WorkBookXmlSheetNode.SetAttribute('sheetId',$IdCounter)
				$NamespaceR = $WorkBookXmlDoc.DocumentElement.GetNamespaceOfPrefix("r")
				If($NamespaceR) {
					$Null = $WorkBookXmlSheetNode.SetAttribute('id',$NamespaceR,$NewWorkBookRelId)
				} Else {
					$Null = $WorkBookXmlSheetNode.SetAttribute('id',$NewWorkBookRelId)
				}
				$Null = $WorkBookXmlDoc.DocumentElement.Item("sheets").AppendChild($WorkBookXmlSheetNode)
				$WorkBookXmlDoc.Save($WorkBookPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Write))
		
				$exPkg.Close()
				New-Object -TypeName PsObject -Property @{Uri = $Uri_xl_worksheets_sheet_xml;
														WorkbookRelationID = $NewWorkBookRelId;
														Name = $Name;
														WorkbookPath = $Path
														}
			} 
			End { 
				} 
		}
		Function New-XLSXWorkBook {	
			param(
#				[Parameter(Mandatory=$True,
#					Position=0,
#					ValueFromPipeline=$True,
#					ValueFromPipelinebyPropertyName=$True
#				)]
				[String]$Path,
				[ValidateNotNull()]
				[Switch]$NoClobber,
				[Switch]$Force
			)
			Begin {
				$xl_Workbook_xml = New-Object System.Xml.XmlDocument
				$XmlDeclaration = $xl_Workbook_xml.CreateXmlDeclaration("1.0", "UTF-8", "yes")
				$Null = $xl_Workbook_xml.InsertBefore($XmlDeclaration, $xl_Workbook_xml.DocumentElement)
				$workBookElement = $xl_Workbook_xml.CreateElement("workbook")
				$Null = $workBookElement.SetAttribute("xmlns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
				$Null = $workBookElement.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
				$Null = $xl_Workbook_xml.AppendChild($workBookElement)
				$Null = $xl_Workbook_xml.DocumentElement.AppendChild($xl_Workbook_xml.CreateElement("sheets"))
			} 
			Process {	
				$Path = [System.IO.Path]::ChangeExtension($Path,'xlsx')
				Try {
					Out-File -InputObject "" -FilePath $Path -NoClobber:$NoClobber.IsPresent -Force:$Force.IsPresent -ErrorAction stop
					Remove-Item $Path -Force
				} Catch {
					$Error.RemoveAt(0)
					$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
					$PSCmdlet.WriteError($NewError)
					Return
				}
				Try {
					$exPkg = [System.IO.Packaging.Package]::Open($Path, [System.IO.FileMode]::Create)
				} Catch {
					$_
					return
				}
				$Uri_xl_workbook_xml = New-Object System.Uri -ArgumentList ("/xl/workbook.xml", [System.UriKind]::Relative)
				$Part_xl_workbook_xml = $exPkg.CreatePart($Uri_xl_workbook_xml, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
				$dest = $part_xl_workbook_xml.GetStream([System.IO.FileMode]::Create,[System.IO.FileAccess]::Write)
				$xl_workbook_xml.Save($dest)
				$Null = $exPkg.CreateRelationship($Uri_xl_workbook_xml, [System.IO.Packaging.TargetMode]::Internal, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", "rId1")
				$exPkg.Close()
				Return Get-Item $Path
			} 
			End {
			} 
		}
		Function Export-WorkSheet {
#			[CmdletBinding()]
			param(
#				[Parameter(Mandatory=$true,
#					Position=1,
#					ValueFromPipeline=$true,
#				)]
				[System.Management.Automation.PSObject]$InputObject,

#				[Parameter(Mandatory=$True,
#					Position=0,
#					ValueFromPipeline=$True,
#					ValueFromPipelinebyPropertyName=$True
#				)]
				[System.String]$Path,
				
#				[Parameter(Mandatory=$True,
#					Position=1,
#					ValueFromPipeline=$True,
#					ValueFromPipelinebyPropertyName=$True
#				)]
				[System.Uri]$WorksheetUri,				
				[Boolean]$NoHeader = 0
				
			)
			Begin {
				$exPkg = [System.IO.Packaging.Package]::Open($Path, [System.IO.FileMode]::Open)
				$WorkSheetPart = $exPkg.GetPart($WorksheetUri)
				$WorkSheetXmlDoc = New-Object System.Xml.XmlDocument
				$WorkSheetXmlDoc.Load($WorkSheetPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Read))
				$HeaderWritten = $False
			}
			Process {
				If($InputObject.GetType().Name -match 'byte|short|int32|long|sbyte|ushort|uint32|ulong|float|double|decimal|string') {
					Add-Member -InputObject $InputObject -MemberType NoteProperty -Name ($InputObject.GetType().Name) -Value $InputObject
				}
				
				If( ($NoHeader -eq $true) ){
					$RowNode = $WorkSheetXmlDoc.CreateElement('row', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
					ForEach($Prop in $InputObject.psobject.Properties) {
						$CellNode = $WorkSheetXmlDoc.CreateElement('c', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
						$Null = $CellNode.SetAttribute('t',"inlineStr")
						$Null = $RowNode.AppendChild($CellNode)
						$CellNodeIs = $WorkSheetXmlDoc.CreateElement('is', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
						$Null = $CellNode.AppendChild($CellNodeIs)
						$CellNodeIsT = $WorkSheetXmlDoc.CreateElement('t', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
						$CellNodeIsT.InnerText = [String]$Prop.Name
						$Null = $CellNodeIs.AppendChild($CellNodeIsT)
						$Null = $WorkSheetXmlDoc.DocumentElement.Item("sheetData").AppendChild($RowNode)	
					}
					$HeaderWritten = $True
				}
				$RowNode = $WorkSheetXmlDoc.CreateElement('row', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
				ForEach($Prop in $InputObject.psobject.Properties) {
					$CellNode = $WorkSheetXmlDoc.CreateElement('c', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
					$Null = $CellNode.SetAttribute('t',"inlineStr")
					$Null = $RowNode.AppendChild($CellNode)
					$CellNodeIs = $WorkSheetXmlDoc.CreateElement('is', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
					$Null = $CellNode.AppendChild($CellNodeIs)
					$CellNodeIsT = $WorkSheetXmlDoc.CreateElement('t', $WorkSheetXmlDoc.DocumentElement.Item("sheetData").NamespaceURI)
					$CellNodeIsT.InnerText = [String]$Prop.Value
					$Null = $CellNodeIs.AppendChild($CellNodeIsT)
					$Null = $WorkSheetXmlDoc.DocumentElement.Item("sheetData").AppendChild($RowNode)
				} 
			} 
			End {
				$WorkSheetXmlDoc.Save($WorkSheetPart.GetStream([System.IO.FileMode]::Open,[System.IO.FileAccess]::Write))
				$exPkg.Close()
			} 
		}
	
			$Path = [System.IO.Path]::GetFullPath($Path)
			$Path = [System.IO.Path]::ChangeExtension($Path,'xlsx')
			If((Test-Path $Path) -and $Append -eq $true ) {
				$WorkSheet = Add-XLSXWorkSheet $WorkSheetName $Path
			} Else {
				Try {
					Out-File -InputObject "" -FilePath $Path -NoClobber:$NoClobber.IsPresent -Force:$Force.IsPresent -ErrorAction stop
					Remove-Item $Path -Force
				} Catch {
					$Error.RemoveAt(0)
					$NewError = New-Object System.Management.Automation.ErrorRecord -ArgumentList $_.Exception,$_.FullyQualifiedErrorId,$_.CategoryInfo.Category,$_.TargetObject
					$PSCmdlet.WriteError($NewError)
					Return
				}
				$Null = New-XLSXWorkBook $Path $NoClobber.IsPresent $Force.IsPresent 
				$WorkSheet = Add-XLSXWorkSheet $WorkSheetName $Path
			}
			$HeaderWritten = $False
	} 
	Process {
        If ($HeaderWritten -eq $false) {
            $WriteHeader = 1
        } else {
            $WriteHeader = 0
        }
        ForEach ($InputObject in $InputObjects) {
		    $rc = Export-WorkSheet $InputObject $Path $WorkSheet.Uri $WriteHeader
    		$HeaderWritten = $True
            $WriteHeader = 0
        }
	} 
	End {
	} 
}

function Connect-WsusServer { 
	param (
		[string]$ComputerName = "localhost",
		[boolean]$UseSSL = $false,
		[int]$Port = 8530
	)
	Write-Host -ForegroundColor Green "[WSUS]Connecting to $Computername on Port $Port. SSL = $UseSSL"
	[reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration") | out-null
	$wsus = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($ComputerName, $UseSSL, $Port)
	return $wsus
}

function Get-ComputersFromWsusServer {
	param (
		[Microsoft.UpdateServices.Internal.BaseApi.UpdateServer]$wsus,
		[string]$TargetGroup
	)

	$ComputersInGroup = ($wsus.GetComputerTargetGroups() | Where { $_.Name -eq $TargetGroup }).GetComputerTargets() 

	Write-Host "Found $($ComputersInGroup.Count) Computers"
	return $ComputersInGroup
}


function main {
    # Main Program
    if ((Get-Config) -eq $false) {
        Write-Error "Could not read the configuration file. Exiting."
        return 1
    }

    $ComputerName = Get-ScriptParameter "ComputerName"
    $UseSSL = Get-ScriptParameter "UseSSL"
    $Port = Get-ScriptParameter "Port"
    $mailto = Get-ScriptParameter "mailto"
    $FullReport = Get-ScriptParameter "FullReport"
    $SkipExportResults = Get-ScriptParameter "SkipExportResults"
    $SkipEmailResults = Get-ScriptParameter "SkipEmailResults"

    $ThisComputerInfo = gwmi WIN32_ComputerSystem
    $ReportFile = Get-ReportFilePath

    $wsus = Connect-WsusServer $ComputerName $UseSSL $Port
    if ($wsus -eq $null) {
	    exit 1
    }

    $replicas = $wsus.GetDownstreamServers() | where {$_.IsReplica -eq $true}


    # Set the WSUS Search Criteria to search in the All Computers Group where Updates have been downloaded and Not Installed
    $TargetGroup = "Active Directory"
    $TargetGroup = "All Computers"

    $ComputersInGroup = Get-ComputersFromWsusServer $wsus $TargetGroup
    if ($ComputersInGroup.Count -eq 0) {
        Write-Warning "No Computers are found using the WSUS Server $ComputerName"
        exit 1
    }
    foreach ($replica in $replicas) {
	    $replicaName = $replica.FullDomainName
	    if ($replicaName -notcontains ".") {
		    if ($replicaName -match "^SomeServer[0-3]") {
			    $replicaName += ".testdomain1.com"
		    } elseif ($replicaName -match "^SomeServer[4-9]") {
			    $replicaName += ".testdomain2.com"
		    }
	    }
	    $wsusReplica = Connect-WsusServer -ComputerName $replicaName -UseSSL $replica.UpdateServer.IsConnectionSecureForApiRemoting -Port $replica.UpdateServer.PortNumber
	    $ComputersInGroup += Get-ComputersFromWsusServer $wsusReplica $TargetGroup
    }

    # Remove production machines from report if it is change stop
    If (!$FullReport) {
        if (Test-IsChangeStop -eq $true) {
		    Write-Host "Change Stop, Removing Prod machines from report or people will scream/question"
    #		$ComputersInGroup = $ComputersInGroup | Where {$Script:Config.NonProduction.Domain.Contains($_.FullDomainName) -like "*.qa.myntet.se" -or $_.FullDomainname -like "*.stmmyntet.se" -or $_.FullDomainname -like "TS*"}
		    $ComputersInGroup = $ComputersInGroup | Where {$Script:Config.NonProduction.Domain.Contains($_.FullDomainName) }
	    }
        if ($ComputersInGroup.Count -eq 0) {
            Write-Warning "No Production Computers are found using the WSUS Server $ComputerName"
            exit 1
        }
    }

    $updateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
    $updateScope.IncludedInstallationStates = 'Failed','Downloaded','NotInstalled','InstalledPendingReboot'
    # Bug, this does not get used in the method GetUpdateInstallationInfoPerUpdate. It is setable, but does not get used.
    $updateScope.ApprovedStates = 'LatestRevisionApproved'

    $ComputerPatchInfo = @{}
    $ComputerAssetInfo = @{}

    Write-Host "Gathering Asset information"
    $ADcomputerNames = "-or Name -eq 'NonExistant'"
    $sqlComputerNames = ""
    $ComputersInGroup | Foreach {
	    $ComputerName = $_.FullDomainName -replace "\..*", ""
	    $sqlComputerNames += ", '$ComputerName'"
        $ADComputerNames += "-or Name -eq '$ComputerName'"
	    Write-Host "`tGathering for $($_.FullDomainName)"	
    }
    $ADComputerNames = $ADComputerNames -replace "^-or ", ""
    $sqlComputerNames = $sqlComputerNames -replace "^, ", ""

    # Patch groups are stored centrally in AD
    If ($Script:Config.PatchGroupLocation -eq "AD") {
        $ADPatchInfo = Get-ADComputer -Filter "$ADComputerNames" -Properties Description,$Script:Config.PatchGroupADAttribute
	    for ($counter = 0; $counter -lt $ADPatchInfo.Count; $counter++) {
            $AssetObj = "" | Select-Object Patchgroup, AssetNumber, Description, Environment
		    $ComputerName = $ADPatchInfo[$counter].Name
		    $AssetObj.Patchgroup = $ADPatchInfo[$counter].($Script:Config.PatchGroupADAttribute)
	        if ($AssetObj.Patchgroup -eq $null) {
                $AssetObj.Patchgroup = ""
            }
		    $AssetObj.Environment = 65535
		    $AssetObj.AssetNumber = $ADPatchInfo[$counter].AssetNumber
		    $AssetObj.Description = $ADPatchInfo[$counter].Description
		    $ComputerAssetInfo[$ComputerName] = $AssetObj
	    } 
    } else {
        # Get the computer asset information from a database table

        if ($StmEnvironment -eq $true) {
        ##	$SqlConnection = connect-sqlServer "testSQLServer.TestDomain.com" "ComputerInventory"
        } else {
        ##	$SqlConnection = connect-sqlServer "prodSQLServer.ProdDomain.com" "ComputerInventory"
        }
        $SqlQuery = "SELECT [NAMN/MÄRKNING] as ComputerName, CIID as AssetNumber, ProductNumber, ProductDescription as Description, Environment, Patchgroup FROM [Computerinventory].[dbo].[AssetInfo] WHERE [ComputerName] IN ($sqlComputerNames)"

        $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
        $SqlCmd.Connection = $SqlConnection
        $SqlCmd.CommandText = $SqlQuery

        $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
        $SqlAdapter.SelectCommand = $SqlCmd 
        $DataSet = New-Object System.Data.DataSet
        ##$SqlAdapter.Fill($DataSet) | Out-Null

        Write-Host "Preparing Asset information"

        # Loop through each asset and save this information for reporting purposes
        $ComputerAssetInfo = @{} 
        if ($DataSet.Tables[0].Rows.Count -gt 0) { 
	        for ($counter = 0; $counter -lt $DataSet.Tables[0].Rows.Count; $counter++) {
		        if ($DataSet.Tables[0].Rows[$counter].ComputerName -ne $Null) {
			        $ComputerName = $DataSet.Tables[0].Rows[$counter].ComputerName
			        $AssetNumber = $DataSet.Tables[0].Rows[$counter].AssetNumber
			        $ComputerAssetInfo[$ComputerName] = $DataSet.Tables[0].Rows[$counter]
		        }
	        } 
        }
        ##$SqlConnection.Close() 
        $DataSet.Dispose()
    }


    $ComputersInGroup = $ComputersInGroup | Sort-Object -Property FullDomainName
    # Filter out anything that is not going to be patched in a day or two
    If (!$FullReport) {      

        Switch ([int](Get-Date -UFormat %u)) {
            1 { $ReportDaysBitmask = 2 } # Today is Monday, Lets create a report for Tuesday
            2 { $ReportDaysBitmask = 4 } # Today is Tuesday, Lets create a report for Wednesday
            3 { $ReportDaysBitmask = 8 } # Today is Wednesday, Lets create a report for Thursday 
            4 { $ReportDaysBitmask = 16 + 32 + 64 + 1 } # Today is Thursday, Lets create a report for Friday, Saturday, Sunday, Monday
            5 { $ReportDaysBitmask = 32 + 64 + 1 } # Today is Friday, Lets create a report for Saturday, Sunday, Monday
            6 { $ReportDaysBitmask = 64 + 1 } # Today is Saturday, Lets create a report for Sunday, Monday
            7 { $ReportDaysBitmask = 1 } # Today is Sunday, Lets create a report for Monday
        }
	    Switch (Get-Date -UFormat %u) {
            1 { $DayOfWeekBitmask = 1 } # Monday
            2 { $DayOfWeekBitmask = 2 } # Tuesday
            3 { $DayOfWeekBitmask = 4 } # Wednesday
            4 { $DayOfWeekBitmask = 8 } # Thursday
            5 { $DayOfWeekBitmask = 16 } # Friday
            6 { $DayOfWeekBitmask = 32 } # Saturday
            7 { $DayOfWeekBitmask = 64 } # Sunday
        }

        $UpcomingPatchGroups = ""
        Foreach ($Patchgroup in $Script:Config.ValidPatchgroups.Patchgroup) {
            $PatchGroupWeekdays = ($Patchgroup.Split(","))[0]
            if ($ReportDaysBitmask -band $PatchGroupWeekdays) {
                $UpcomingPatchGroups += "|(^$Patchgroup$)"
            }
        }
        $UpcomingPatchGroups = $UpcomingPatchGroups -Replace "^\|", ""

	    # Convert ComputersInGroup to a Dynamic Array so that we can remove elements from it
	    $ComputersInGroup = [System.Collections.ArrayList]$ComputersInGroup

	    Write-Host -ForegroundColor Green "Filtering Report. Reporting on Computers for the next day or two"
	    Write-Host "`tUpcoming Patchgroups: $UpcomingPatchGroups"
	    $IndexElementsToRemove = @()
	    For ($counter=0; $counter -lt $ComputersInGroup.count; $counter++) {
 		    $ComputerName = $ComputersInGroup[$counter].FullDomainName -replace "\..*", ""
 		    if ($ComputerAssetInfo[$ComputerName].PatchGroup -notmatch $UpcomingPatchGroups) {
 			    Write-Host -ForegroundColor DarkGray "`t[REMOVE] $ComputerName, Patchgroup is $($ComputerAssetInfo[$ComputerName].PatchGroup)"
			    $IndexElementsToRemove += $counter
 		    } else {
 			    Write-Host -ForegroundColor Gray "`t[KEEP]   $ComputerName, Patchgroup is $($ComputerAssetInfo[$ComputerName].PatchGroup)"		
 		    }
	    }
    }
    # Remove the elements from the array in reverse order as the array will become one shorter each time
    $IndexElementsToRemove | Sort-Object -Descending | % { $ComputersInGroup.RemoveRange($_, 1) }

    # Get the computers and their Downloaded/Not Installed UpdateIDs (KB Numbers don't exist in the update info as strings)
    # Store Each Computer and their associated required updates for reporting purposes
    Write-Host "Gathering Updates for each Computer"
    $updateInfo = @{}
    $counter = 0
    $ComputersInGroup | ForEach {
	    $counter++
	    Write-Progress -Activity "Gathering Updates for for $($ComputersInGroup.count) Computers" -CurrentOperation "Found Computer $($_.fulldomainname)" -PercentComplete (($counter / $ComputersInGroup.Count) * 100)
	    $Computername = $_.fulldomainname
	    Write-Host "`tAdding Computer: $Computername"

	    $updates = $_.GetUpdateInstallationInfoPerUpdate($updateScope) | Where {$_.UpdateApprovalAction -ne "NotApproved"}
	    # If no patches are needed, it will not be added to the report and therefore we don't have to lookup the assetinfo/patch info
	    $updates | ForEach {
		    #Write-Host "Computer: $Computername, patch: $($update.Title)"
		    $updateInfo[$_.UpdateId.ToString()] = $null
		    $ComputerPatchInfo[$Computername] += @($_.UpdateId)
	    }
    }


    # We don't send these in the email, so if we're not exporting to the CSV, let's just save some time and skip this!
    if (!$SkipExportResults) {    
        Write-Progress -Activity "Gathering Updates for each Computer for $($ComputersInGroup.count) Computers"
        # Get the Update Info for each patch and store it in memory for reporting purposes
        Write-Host "Gathering Update Details"
        $counter = 0
        Foreach ($UpdateGuid in $($updateInfo.Keys)) {
	        $counter++
	        Write-Progress -Activity "Retrieving $($updateInfo.count) Update Details" -CurrentOperation "Getting Update info for GUID: $($UpdateGuid) " -PercentComplete (($counter / $updateInfo.Count) * 100)
	        Write-Host "`tUpdate info for GUID: $($UpdateGuid)"	
	        $updateInfo[$UpdateGuid] = $wsus.GetUpdate([GUID]$UpdateGuid)
        }
        Write-Progress -Activity "Retrieving Update details" -Completed

    }

    Write-Host "Preparing Reportable Values"
    # Create the report
    $reportArray = @()
    Foreach ($computerFQDN in $ComputerPatchInfo.Keys) {
	    Write-Host "`tBuilding Report info for $computerFQDN"
	    $ComputerName = $computerFQDN -replace "\..*", ""
	    $computerInfo = $ComputersInGroup | Where {$_.FullDomainName -eq $computerFQDN}	
	    $ComputerEnvironment = convert-EnvironmentNumberToText $ComputerAssetInfo[$ComputerName].Environment
	    $patchgroupText = Get-PatchGroupDescription @($ComputerAssetInfo[$ComputerName].Patchgroup)
	    $CountOutstandingPatches = $ComputerPatchInfo[$computerFQDN].count
	    Foreach ($updateGUID in $ComputerPatchInfo[$computerFQDN]) {
		    $reportObject = "" | Select-Object ComputerName,IPAddress,LastReportedStatusTime,AssetNumber,ProductNumber,Description,Environment,PatchGroup,PatchGroupSchedule,CountOutstandingPatches,WsusServer,KBTitle,KBGuid,KBLink
		    $reportObject.ComputerName = $computerFQDN
		    $reportObject.IPAddress = $computerInfo.IPAddress
		    $reportObject.LastReportedStatusTime = $computerInfo.LastReportedStatusTime
		    $reportObject.AssetNumber = $ComputerAssetInfo[$ComputerName].AssetNumber
		    $reportObject.Description = $ComputerAssetInfo[$ComputerName].Description
		    $reportObject.ProductNumber = $ComputerAssetInfo[$ComputerName].ProductNumber
		    $reportObject.Environment = $ComputerEnvironment
		    $reportObject.PatchGroup = $ComputerAssetInfo[$ComputerName].Patchgroup
		    $reportObject.PatchGroupSchedule = $patchgroupText
		    $reportObject.WsusServer = $computerInfo.UpdateServer.Name
		    $reportObject.KBTitle = $updateInfo[$UpdateGuid.Guid].Title
		    $reportObject.CountOutstandingPatches = $CountOutstandingPatches
		    $reportObject.KBGuid = $UpdateGuid.Guid
		    $reportObject.KBLink = $updateInfo[$UpdateGuid.Guid].AdditionalInformationUrls
		    $reportArray += $reportObject
	    }
	    $counter++
    }
    $reportArray = $reportArray | Sort-Object -Property ComputerName


    if (!$SkipExportResults) {
	    # Create the Report
	    $tabnames = @("Updates-Info")
	    $tabnames += $reportArray | Select-Object -Unique -Property ProductNumber | Where {$_.ProductNumber -ne $null} | Select-Object -ExpandProperty ProductNumber | % { "Product-$(if ($_) {$_} else {"Unknown"})"}
	    $tabnames += $reportArray | Select-Object -Unique -Property Patchgroup | Select-Object -ExpandProperty Patchgroup | % { "Patchgroup-$(if ($_) {$_} else {"Unknown"})"}
	
	    $tabname = "Everything"
	    Write-Host "Generating Report"
	    Write-Host "`tExporting $tabname Data"
        $appendToXLSX = $false
	    $rc = Export-XLSX $ReportArray $ReportFile $tabname $appendToXLSX

        $appendToXLSX = $true
	    Foreach ($tabname in $tabnames) {
		    Write-Host "`tExporting $tabname Data"
		    $tabvalue = $tabname -replace "^.*-", ""
		    if ($tabvalue -match "Unknown") {
			    $tabvalue = ""
		    }
            $ReportData = ""
		    if ($tabname -imatch "^Product") {
			    $ReportData = $reportArray | Where {$_.ProductNumber -eq $tabvalue} | Select-Object * -ExcludeProperty ProductNumber,CountOutstandingPatches
		    } elseif ($tabname -imatch "^Patchgroup") {
			    $ReportData = $reportArray | Where {$_.Patchgroup -eq $tabvalue} | Select-Object * -ExcludeProperty PatchGroup,CountOutstandingPatches
		    } elseif ($tabname -imatch "^Updates") {
			    $ReportData = $updateInfo.values | Select-Object Title,LegacyName,MsrcSeverity,KnowledgebaseArticles,SecurityBulletins,AdditionalInformationUrls,UpdateClassificationTitle,ProductTitles,CreationDate,Description
		    }
            Export-XLSX $ReportData $ReportFile $tabname $appendToXLSX
	    }
    }

    if ($reportArray.Count -eq 0) {
	    $SkipEmailResults = $true
	    Write-Host "No Computers required updating. Not emailing the report."
    }

    if (!$SkipEmailResults) {
	    $serverlistHtml = $reportArray | Select-Object -Property @{Name="Computer Name";Expression={$_."ComputerName"}}, @{Name="Description";Expression={$_."Description"}}, @{Name="Product";Expression={$_."ProductDescription"}}, @{Name="Patch Schedule";Expression={$_.PatchGroupSchedule}}, @{Name="# Outstanding Patches";Expression={$_.CountOutstandingPatches}} -Unique | Sort-Object -Property "Product Number", "Computer Name" | ConvertTo-Html -Fragment
	    $HtmlStyle = "<style>"
	    $HtmlStyle += "BODY{background-color:white;}"
	    $HtmlStyle += "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;}"
	    $HtmlStyle += "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:thistle;text-align:left}"
	    $HtmlStyle += "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:palegoldenrod}"
	    $HtmlStyle +=  "</style>"
	
        $PatchTuesday = Get-PatchTuesday
        $PatchTuesdayChangeStopDays = Get-PatchTuesdayChangeStopDays
        $DaysLeftInPatchTuesdayChangeStop = Get-DaysLeftInPatchTuesdayChangeStop

	    $htmlemailbody = @"
	<h1>Upcoming Patches will be Installed</h1>
	<p>
		The full report and details can be viewed at the following link:<br>
		<a href='$ReportFile'>$ReportFile</a><br>
	</p>
	<h2>The following Servers will be patched:</h2>
	
	$serverlistHtml
	
	<p>
        <li>Patch Tuesday this month is: $($PatchTuesday.ToLongDateString())</li>"
		<li>Production Machines:</li>
		<ul>Will be patched between the 1st to the $(($PatchTuesday.AddDays(-1)).ToLongDateString()) and $(($PatchTuesday.AddDays($DaysLeftInPatchTuesdayChangeStop + ($PatchTuesdayChangeStopDays))).ToLongDateString()) to the end of the month according to their Patchgroup</ul>
		<li>Non Production Machines:</li>
		<ul>Are patched everyday of the month according to their Patchgroup. If an invalid/unset patch group is found, they will be patched using the default patch group</ul>
	</p>
	<p>
		<span style='color:#DCDCDC; font-size:8pt'>This report was generated on $($ThisComputerInfo.Name).$($ThisComputerInfo.Domain)</span>
	</p>
	
"@
	    $htmlemailbody = ConvertTo-Html -Head $HtmlStyle -Body $htmlemailbody  | Out-String
	
    #	$mailTo = @("arafuse@broadarc.com")
        if ($mailto.Count -le 0) {            
            $mailTo = ($Script:Config.EmailSendReportTo).Split(",")
        }
        $domain = (Get-WmiObject Win32_NTDomain | Where {$_.DomainName -ne $null}).DomainName
	    if ($FullReport) {
		    $subject = "[$domain] WSUS Report - FULL REPORT"
	    } else {
		    $subject = "[$domain] WSUS Report"
	    }
	    if ($StmEnvironment -eq $true) {
		    $subject = "STM ENVIRONMENT - $subject"
	    }
        Write-Host "Sending mail to $mailTo from $($Script:Config.EmailSendReportFrom) Mail Server: $($Script:Config.SMTPServer)"
	    $rc = Send-MailMessage -To $mailTo -Subject $subject -From "Windows Patch Report <$($Script:Config.EmailSendReportFrom)>" -SmtpServer $Script:Config.SMTPServer -Body $htmlemailbody -BodyAsHtml -ErrorAction SilentlyContinue
    }	

    return 0
}

$rc = main
exit $rc