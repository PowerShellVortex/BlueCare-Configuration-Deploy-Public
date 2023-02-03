# BlueCare Configuration Deploy script
param( [parameter(Position = 0)]$XLS_FileName, [switch]$ApplyConfiguration, [switch]$NoLog, [switch]$LocalTesting )

#region Cleanup Variables

$g_DataSettings = $null

$g_DateTime = Get-Date -Format yyyy.MM.dd-HH.mm.ss

$script:SynchronizedDataLogs = $null
$script:SynchronizedDataFiles = $null

$CSV_FileName = $null
$XLS_FileName = $null
$XLS_SheetName = $null

$g_CfgClientName = $null
$g_CfgMasterPath = $null
$g_CfgSourceFolder = $null
$g_CfgDestinationFolder = $null
$g_CfgLogFolderName = $null
$g_CfgLogFileName = $null
$g_CfgLogFilePath = $null

$dataCSVHash = $null

$g_DataLogErrors = @()
$g_DataLogSuccess = @()
#endregion
#region Global configuration
# Get actual script path
if ( $script:MyInvocation.MyCommand.Path ) { $g_ScriptPath = Split-Path $script:MyInvocation.MyCommand.Path } else { $g_ScriptPath = Split-Path -Parent $psISE.CurrentFile.Fullpath }

# Read XML configuration file
[XML]$g_DataSettings = Get-Content "$g_ScriptPath\BlueCare-Configuration-Deploy-Settings.xml"

# client name, used for prefixes for config and logs folders
$g_CfgClientName = $g_DataSettings.Configuration | % { $_.g_CfgClientName.value }

# XLS sheet name for xls2csv.exe conversion tool
$XLS_SheetName = $g_DataSettings.Configuration | % { $_.XLS_SheetName.value }

# default XLS file name for deploy configuration
$XLS_FileName = $g_DataSettings.Configuration | % { $_.XLS_FileName.value }

# full path of the master folder for configuration and log folders
$g_CfgMasterPath = $g_DataSettings.Configuration | % { $_.g_CfgMasterPath.value }

# full path to the remote destination folder where BlueCare stores its configuration files
if ( $ApplyConfiguration ) {
    $g_CfgDestinationFolder = $g_DataSettings.Configuration | % { $_.g_CfgDestinationFolder.value }
} else { $g_CfgDestinationFolder = "c$\Windows\Temp" }

# full path of the folder where we stored configuration files
$g_CfgSourceFolder = "$g_CfgMasterPath\$g_CfgClientName-Configs"

# log path
$g_CfgLogFolderName = "$g_CfgMasterPath\$g_CfgClientName-Logs"
$g_CfgLogFileName = "$g_CfgClientName-$g_DateTime.log"
$g_CfgLogFilePath = "$g_CfgLogFolderName\$g_CfgLogFileName"

# If you want to store logs locally at the script path (separated from the configuration source) then uncomment lines below
# $g_CfgLogFolderName = "$g_ScriptPath\$g_CfgClientName-Logs"
# $g_CfgLogFilePath = "$g_CfgLogFolderName\$g_CfgLogFileName"

# for local testing, create local network share "\\BlueCare\Configs", give write rights, and run script using -LocalTesting switch
if ( $LocalTesting ) {
    $g_CfgMasterPath = $g_ScriptPath
    $g_CfgSourceFolder = "$g_CfgMasterPath\$g_CfgClientName-Configs"
    $g_CfgDestinationFolder = "BlueCare\Configs"
    $g_CfgLogFolderName = "$g_CfgMasterPath\$g_CfgClientName-Logs"
    $g_CfgLogFilePath = "$g_CfgLogFolderName\$g_CfgLogFileName"
}
#endregion
#region Runspace Function
function Start-ConfigurationDeploy {
    param( [parameter(Position = 0, Mandatory = $true)]$xData, [int]$Throttle = 10)
    Begin {
        Write-Host ""
        Write-Host "OVERWRITING CONFIGURATION FILES:" -ForegroundColor Cyan
        Write-Host ""

        # Define hash table for Get-RunspaceData function
        $runspacehash = @{}

        # Function to perform runspace job cleanup
        Function Get-RunspaceData {
            [cmdletbinding()]param( $dataVariable, [switch]$Wait )
            Do {
                $more = $false
                Foreach ( $runspace in $runspaces ) {
                    If ( $runspace.Runspace.isCompleted ) {
                        $runspace.powershell.EndInvoke( $runspace.Runspace )
                        $runspace.powershell.dispose()
                        $runspace.Runspace = $null
                        $runspace.powershell = $null
                    } ElseIf ( $null -ne $runspace.Runspace ) { $more = $true }
                }
                If ( $more -and $PSBoundParameters['Wait'] ) { Start-Sleep -Milliseconds 100 }
                # Clean out unused runspace jobs
                $temphash = $runspaces.clone()
                $temphash | ? { $null -eq $_.runspace } | % {
                    Write-Verbose ("Removing {0}" -f $_.$dataVariable)
                    $Runspaces.remove($_)
                }
            } while ( $more -and $PSBoundParameters['Wait'] )
        }

        #region ScriptBlock
        $scriptBlock = { param([parameter(Position = 0)]$dataCSVHash)

            $ComputerName = $dataCSVHash.ComputerName
            $PowershellRemoting = $dataCSVHash.PowershellRemoting
            $localComputerName = $dataCSVHash.localComputerName
            $g_CfgClientName = $dataCSVHash.g_CfgClientName
            $g_CfgSourceFolder = $dataCSVHash.g_CfgSourceFolder
            $g_CfgDestinationFolder = $dataCSVHash.g_CfgDestinationFolder

            $CSV_Row = $dataCSVHash.CSV_Row
            $CSV_HeaderColumns = $dataCSVHash.CSV_HeaderColumns

            #Write-Output "Use Powershell remoting for $ComputerName : $($dataCSVHash.PowershellRemoting)"

            # for local testing
            if ( $Env:COMPUTERNAME -eq $localComputerName ) { $ComputerName = $Env:COMPUTERNAME }

            [array]$script:errorsFileCopy = $null

            [array]$g_DataLogSuccess += ( "$ComputerName" + ";" )
            [array]$script:g_DataLogErrors += ( "$ComputerName" + ";" )

            if ( $PowershellRemoting ) {
                # to do: use Win32_Share to resolve local path
                $remoteDestinationFolder = $g_CfgDestinationFolder -replace '^(.*?)\$(.*)', '$1:$2'
            } else {
                $remoteDestinationFolder = "\\$ComputerName\$g_CfgDestinationFolder"
            }

            if ( Test-Path -Path $remoteDestinationFolder -PathType Container -EA 1 ) {
                Write-Output "$($ComputerName):"

                $CSV_HeaderColumns | % {

                    $cfgDestinationFileName = $CSV_ColumnName = $_
                    # Name of the source file from excel, regardless of whether the file exists in the data variable or not
                    $cfgSourceFileName = $CSV_Row."$CSV_ColumnName"
                    # Does xls document have any data in server+column cell? If not, then there wasn't any source file for this server
                    if ( $null -ne $cfgSourceFileName ) {
                        # Only used for detailed log output, not to read the data from the source file
                        $cfgSourceFilePath = $g_CfgSourceFolder + "\" + $CSV_ColumnName + "\" + $cfgSourceFileName
                        $remoteDestinationFilePath = "$remoteDestinationFolder\$cfgDestinationFileName"

                        # for PS 3.0
                        # $remoteDestinationFileData = $SynchronizedDataFiles."$cfgDestinationFileName"."$cfgSourceFileName"
                        # but for PS 2.0 OMG look below at this hell ...

                        # key name of the destination file name set
                        $hashKeyDestinationFileName = $SynchronizedDataFiles | % { $_.Keys } | ? { $_ -contains $cfgDestinationFileName }

                        # Testing of the key name of the source file, from the destination file data set
                        # Does the source file data variable contains the proper key name of the source file, for the corresponding destination file name?
                        # if not then we know that it was not read from the disc because the source file doesn't exist (Get-xData)
                        $cfgSourceFileNameTest = ( $SynchronizedDataFiles.$hashKeyDestinationFileName | % { $_ } | ? { $_.Keys -contains $cfgSourceFileName } ).keys

                        if ( $null -ne $cfgSourceFileNameTest ) {
                            $remoteDestinationFileData = ( $SynchronizedDataFiles.$hashKeyDestinationFileName | % { $_ } | ? { $_.Keys -contains $cfgSourceFileName } ).values
                            if ((( $remoteDestinationFileData -replace " " ) -replace "`r`n" ) -ne "" ) {
                                try {
                                    Out-File -InputObject $remoteDestinationFileData -FilePath $remoteDestinationFilePath -Force -Encoding Default -EA 1
                                    $g_DataLogSuccess += $CSV_ColumnName + " < " + $cfgSourceFileName + ";"
                                    Write-Output " Success: $_ < $cfgSourceFileName"
                                } Catch [System.IO.IOException] {
                                    Write-Output "!  Error: $_"
                                    $script:errorsFileCopy += $CSV_ColumnName ; $script:g_DataLogErrors += "Error: " + $CSV_ColumnName + ":" + $_.FullyQualifiedErrorID + " < " + $_.Exception.Message + ";"
                                }
                            } else {
                                Write-Output "!  Error: $_ < $cfgSourceFileName has no valid data!"
                                $script:errorsFileCopy += $CSV_ColumnName ; $script:g_DataLogErrors += "Error: " + $CSV_ColumnName + " < " + "$cfgSourceFilePath has no valid data!" + ";"
                            }
                        } else {
                            Write-Output "!  Error: $cfgDestinationFileName < $cfgSourceFileName don't exist!"
                            $script:errorsFileCopy += $CSV_ColumnName ; $script:g_DataLogErrors += "Error: " + $CSV_ColumnName + " < " + "$cfgSourceFilePath don't exist!" + ";"
                        }
                    }
                }
            } else {
                try { Get-Item -Path $remoteDestinationFilePath -EA 1 }
                Catch {
                    Write-Output "!  Error: $remoteDestinationFolder < path_does_not_exist/access_denied"
                    $script:g_DataLogErrors += "Error: $remoteDestinationFolder < path_does_not_exist/access_denied"
                }
            }

            if ( !( $g_DataLogSuccess.count -gt 1 ) ) { $g_DataLogSuccess = $null }
            if ( !( $g_DataLogErrors.count -gt 1 ) ) { $g_DataLogErrors = $null }

            [array]$DataLogs = "$g_DataLogSuccess", "$g_DataLogErrors"
            $SynchronizedDataLogs."$ComputerName" = $DataLogs # here I'm adding computername as key and log line as value to the Synchronized hashtable variable

            Write-Output ""
        }
        #endregion

        #region Runspacepool Creation
        Write-Verbose ("Creating runspace pool and session states")
        $sessionstate = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $runspacepool = [runspacefactory]::CreateRunspacePool(1, $Throttle, $sessionstate, $Host)
        $runspacepool.Open()
        Write-Verbose ("Creating empty collection to hold runspace jobs")
        $Script:runspaces = New-Object System.Collections.ArrayList
        #endregion Runspace Creation

    }

    Process {

        $xData | % {

            $dataCSVHash = $_
            # Create the powershell instance and supply the scriptblock with the other parameters
            $powershell = [powershell]::Create()
            $powershell.AddScript($scriptBlock).AddArgument($dataCSVHash) | Out-Null

            # Add the runspace into the powershell instance
            $powershell.RunspacePool = $runspacepool

            # Powershell Remoting, valid credentials for the destination server need to be added to "Credential Manager"

            if ($dataCSVHash.PowershellRemoting) {

                # Create connectionInfo
                $Uri = New-Object System.Uri("http://$($dataCSVHash.ComputerName):5985/wsman")
                $connectionInfo = New-Object System.Management.Automation.Runspaces.WSManConnectionInfo -ArgumentList $Uri
                $connectionInfo.OpenTimeout = 3000

                # Create remote runspace
                $runspace = [runspacefactory]::CreateRunspace($connectionInfo)
            } else {
                # Create local runspace
                $runspace = [runspacefactory]::CreateRunspace()
            }

            # Set variable data that will be synchronized
            $runspace.Open()
            $runspace.SessionStateProxy.SetVariable('SynchronizedDataLogs', $SynchronizedDataLogs)
            $runspace.SessionStateProxy.SetVariable('SynchronizedDataFiles', $SynchronizedDataFiles)
            $powershell.Runspace = $runspace

            # Create a temporary collection for each runspace
            $temp = "" | Select-Object PowerShell, Runspace, dataCSVHash
            $temp.dataCSVHash = $dataCSVHash
            $temp.PowerShell = $powershell

            # Save the handle output when calling BeginInvoke() that will be used later to end the runspace
            $temp.Runspace = $powershell.BeginInvoke()

            Write-Verbose ("Adding {0} collection" -f $temp.dataCSVHash.ComputerName)
            $runspaces.Add($temp) | Out-Null

            Write-Verbose ("Checking status of runspace jobs")
            Get-RunspaceData @runspacehash -dataVariable "dataCSVHash.ComputerName"
        }

    }

    End {
        Write-Verbose ("Finish processing the remaining runspace jobs: {0}" -f ( $runspaces | ? { $null -ne $_.Runspace }).Count )
        $runspacehash.Wait = $true
        Get-RunspaceData @runspacehash

        #region Cleanup Runspace pool
        Write-Verbose ("Closing the runspace pool")
        $runspacepool.close()
        $runspacepool.Dispose()
        #endregion Cleanup Runspace
    }

}
#endregion
#region Functions
function Convert-XLS {
    param( [parameter(Position = 0, Mandatory = $true)]$XLS_FileName )
    .\xls2csv.exe $XLS_FileName 2>&1 | Out-Null
}
function Get-Duplicate {
    param( [parameter(Position = 0, Mandatory = $true)]$array )
    $hash = @{}
    $array | % { $hash[$_] = $hash[$_] + 1 }
    $hash.GetEnumerator() | ? { $_.value -gt 1 } | % { $_.key }
}
function Get-xData {
    param([parameter(Position = 0)]$dataCSV)
    Write-Host "READING BlueCare CONFIGURATION FILES" -ForegroundColor Cyan
    Write-Host ""

    $script:SynchronizedDataFiles = [hashtable]::Synchronized(@{})
    $script:SynchronizedDataLogs = [hashtable]::Synchronized(@{})
    [array]$script:xData = $null
    $dataCSV = $dataCSV | ? { $_.ApplyConfiguration -eq "Yes" -and (( $_.ComputerName -Replace " " ) -ne "" ) }
    $dataCSV | % { if ( $LocalTesting ) { $_.ComputerName = $Env:COMPUTERNAME }
        $ComputerName = $_.ComputerName
        $PowershellRemoting = if (( $_.PowershellRemoting -replace " " ) -ne "Yes" ) {
            $false
        } else {
            if ( $PSVersionTable.PSVersion.Major -lt 3 ) {
                Write-Host "Using Powershell Remoting requires Powershell 3.0 because somehow `$runspace.SessionStateProxy property is empty using Powershell 2.0 :(" ; break 
            }
            $true
        }
        $row = $_

        $CSV_HeaderColumns | % { if ( !( Get-Variable $_ -EA 0 ) ) { New-Variable -Name $_ -Value @() -Scope script } }

        $CSV_HeaderColumns | % {

            $CSV_ColumnName = $_
            $hT = @{}
            $cfgSourceFileName = $row."$CSV_ColumnName"

            if ( $null -ne $cfgSourceFileName ) {
                if ( ( ( (($cfgSourceFileName).trimStart(" ")).trimEnd(" ") ) -replace "`r`n" ) -ne "" ) {
                    $cfgSourceFilePath = $g_CfgSourceFolder + "\" + $CSV_ColumnName + "\" + $cfgSourceFileName

                    # Powershell 2.0 MADNEEEESSS!!
                    # ((( Get-Variable $CSV_ColumnName ) | % { $_.Value } | % { $_.Keys } ) -contains $cfgSourceFileName )
                    # vs
                    # ( ( Get-Variable $CSV_ColumnName ).Value.Keys -contains $cfgSourceFileName )

                    if ( Get-Item -Path $cfgSourceFilePath -EA 0 ) {
                        if ( !( ( ( Get-Variable $CSV_ColumnName ) | % { $_.Value } | % { $_.Keys } ) -eq $cfgSourceFileName ) ) {
                            Write-Host " Reading: $g_CfgClientName-Configs\$CSV_ColumnName\$cfgSourceFileName" -ForegroundColor Green
                            $sourceFileData = [System.IO.File]::ReadAllText( $cfgSourceFilePath )
                            $hT.Add( $cfgSourceFileName , $sourceFileData )
        (Get-Variable -Name $CSV_ColumnName).Value += $hT
                        }
                    } else {
                        Write-Host "!  Error: $cfgSourceFilePath don't exist!" -ForegroundColor Red -BackgroundColor White
                        return 
                    }
                }
            }
        }

        $dataCSVHash = @{
            ComputerName           = $_.ComputerName
            PowershellRemoting     = $PowershellRemoting
            localComputerName      = $localComputerName
            CSV_Row                = $_
            CSV_HeaderColumns      = $CSV_HeaderColumns
            g_CfgClientName        = $g_CfgClientName
            g_CfgSourceFolder      = $g_CfgSourceFolder
            g_CfgDestinationFolder = $g_CfgDestinationFolder
        }
        $script:xData += $dataCSVHash
    }
    $CSV_HeaderColumns | % {
        $CSV_ColumnName = $_
        $SynchronizedDataFiles.Add( (Get-Variable -Name $CSV_ColumnName).Name , (Get-Variable -Name $CSV_ColumnName).Value )
    }

    return $xData
}
function Get-XLSFile {
    param( [parameter(Position = 0)]$XLS_FileName )
    if ( !$XLS_FileName ) { $XLS_FileName = ( Read-Host -Prompt "XLS Filename without extension" ) + ".xls" }
    $XLS_FileName
}
function Get-Data {
    param( [parameter(Position = 0, Mandatory = $true)]$XLS_FileName )
    if ( Test-Path $XLS_FileName ) {
        Convert-XLS $XLS_FileName
        $script:CSV_FileName = ( Get-Item $XLS_FileName ).BaseName + "_" + $XLS_SheetName + ".csv"
        [array]$script:CSV_HeaderColumns = ( ( Get-Content $CSV_FileName -TotalCount 1 ) -replace '"' ) -split ";" | ? { $_ -like "K06_*.param" }
        $CSV_HeaderColumns | % { Remove-Variable $_ -Scope script -ErrorAction SilentlyContinue }
        $dataCSV = Import-Csv -Path $CSV_FileName -Delimiter ";"

        [array]$serverList = @()
        $dataCSV | % { $serverList += $_.ComputerName -Replace " " }
        [array]$duplicatedServers = Get-Duplicate $serverList
        if ( $duplicatedServers ) {
            Write-Host "Error, duplicated servers found in $XLS_FileName!"
            $duplicatedServers
            break
        } else {
            Write-Host "No duplicated servers found."
            Write-Host ""
            $dataCSV
        }
    }
}
function Export-Log {
    param([parameter(Position = 0)]$DataLogs, [parameter(Position = 1)]$LogsFileName)

    if ( $NoLog ) { Write-Host "Logging was disabled." } else {
        if ( !( Test-Path $g_CfgLogFolderName ) ) { New-Item -Path $g_CfgLogFolderName -ItemType Directory }
        Write-Host ""
        Write-Host "PROCESSING LOGS:" -ForegroundColor Cyan
        Write-Host "$g_CfgLogFilePath"
        if ( $DataLogs ) {
            $DataLogs.keys | % {
                $DataLogs.$_ | % { if ( $_ -ne "" ) { Write-Output $_ } }
            } | Out-File -FilePath $LogsFileName -Append -Force
        }
    }
}
function Show-GlobalConfiguration {
    Write-Host ""
    Write-Host "GLOBAL CONFIGURATION:" -ForegroundColor Cyan
    Write-Host "         Client Name: $g_CfgClientName"
    Write-Host "       Config Source: $g_CfgSourceFolder"
    Write-Host "  Config Destination: <ComputerName>\$g_CfgDestinationFolder"
    Write-Host "            Log path: $g_CfgLogFolderName"
    Write-Host ""
}
#endregion

Show-GlobalConfiguration

if ( $verbose ) {
    Start-ConfigurationDeploy ( Get-xData ( Get-Data ( Get-XLSFile $XLS_FileName ))) -verbose
} else { Start-ConfigurationDeploy ( Get-xData ( Get-Data ( Get-XLSFile $XLS_FileName ))) }

Export-Log $SynchronizedDataLogs $g_CfgLogFilePath
