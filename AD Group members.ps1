#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog
#Requires -Modules Toolbox.Remoting

<#
    .SYNOPSIS
        Create a list of AD groups with their members.

    .DESCRIPTION
        Retrieve all the members of an active directory group and create a list
        of the members of the group and the group structure.

    .PARAMETER MaxThreads
        Maximum number of jobs allowed to run at the same time.

    .PARAMETER ImportFile
        A .json file containing the script arguments.

    .PARAMETER LogFolder
        Location for the log files.
#>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String[]]$ImportFile,
    [Int]$MaxThreads = 3,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\Get group members\$ScriptName",
    [String[]]$ScriptAdmin = @(
        $env:POWERSHELL_SCRIPT_ADMIN,
        $env:POWERSHELL_SCRIPT_ADMIN_BACKUP
    )
)

Begin {
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import input file
        $File = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json

        if (-not ($MailTo = $File.MailTo)) {
            throw "Input file '$ImportFile': No 'MailTo' addresses found."
        }

        if (-not ($adGroupNames = $File.AD.GroupNames)) {
            throw "Input file '$ImportFile': No 'AD.GroupNames' found."
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}

Process {
    Try {
        $jobs = $jobResults = @()

        $init = { Import-Module Toolbox.ActiveDirectory }

        ForEach (
            $group in
            (
                $adGroupNames |
                Sort-Object -Property @{Expression = { $_.Trim() } } -Unique
            )
        ) {
            Write-Verbose "Start job for group '$group'"

            $jobs += Start-Job -Name $ScriptName -InitializationScript $init -ScriptBlock {
                Param (
                    $Group
                )

                if (-not
                    ($MembersFlat = Get-ADGroupMemberFlatHC -Identity $Group)
                ) {
                    $MembersFlat = $null
                }

                if (-not
                    ($MemberUsers = Get-ADGroupMember -Identity $Group -Recursive |
                    Select-Object @{L = $Group; E = { $_.Name } })
                ) {
                    $MemberUsers = $null
                }

                [PSCustomObject]@{
                    Name        = $Group
                    MembersFlat = $MembersFlat
                    MemberUsers = $MemberUsers
                }
            } -ArgumentList $group

            Wait-MaxRunningJobsHC -Name $jobs -MaxThreads $MaxThreads
        }

        if ($jobs) {
            $jobResults = $jobs | Wait-Job | Receive-Job
            Write-Verbose "Jobs done"
        }
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
    Finally {
        Get-Job | Remove-Job -Force
    }
}

End {
    Try {
        $excelParams = @{
            Path            = $logFile + ' - Result.xlsx'
            AutoSize        = $true
            BoldTopRow      = $true
            FreezeTopRow    = $true
            AutoFilter      = $true
            ConditionalText = $(
                New-ConditionalText ~* Black Orange
            )
        }

        $mailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = 'Success'
            Message   = 'This report checks security groups for their members.'
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        if ($jobResults) {
            $Sheet = 0
            $jobResults | ForEach-Object {
                if ($_.MembersFlat) {
                    $M = "Group '$($_.Name)' has $($_.MembersFlat.Count) end node members"
                    Write-Verbose $M
                    Write-EventLog @EventVerboseParams -Message $M

                    #region Export to Excel
                    Write-Verbose "Export to Excel file '$($excelParams.Path)'"
                    $Sheet++

                    # Excel worksheet names can only be 31 chars long and
                    # must be unique
                    $_.MembersFlat | Sort-Object -Property * |
                    Update-FirstObjectProperties |
                    Export-Excel @excelParams -WorksheetName ("$Sheet (Flat) " + $_.Name)

                    if ($_.MemberUsers) {
                        $_.MemberUsers | Sort-Object -Property * |
                        Export-Excel @excelParams -WorksheetName ("$Sheet (Users) " + $_.Name)
                    }
                    else {
                        $_.Name | Export-Excel @excelParams -WorksheetName ("$Sheet (Users) " + $_.Name)
                    }
                    #endregion

                    #region Check circular group membership
                    foreach ($group in $_.MembersFlat) {
                        if (
                            $group.PSObject.Members |
                            Where-Object MemberType -EQ NoteProperty |
                            Where-Object Value -Match '\*'
                        ) {
                            $circularGroup = $true
                            break
                        }
                    }
                    #endregion

                    $mailParams.Attachments = $excelParams.Path
                }
                else {
                    $NoGroupMembers = $true
                }

                #region format group names
                if ($circularGroup) {
                    $circularGroup = $false

                    $M = "Group '$($_.Name)' has circular group membership"
                    Write-Warning $M
                    Write-EventLog @EventWarnParams -Message $M
                    $_.Name += ' <b>(Circular group membership found)</b>'
                }

                if ($NoGroupMembers) {
                    $NoGroupMembers = $false

                    $M = "Group '$($_.Name)' has no end node members"
                    Write-Warning $M
                    Write-EventLog @EventWarnParams -Message $M
                    $_.Name += ' <b>(No members found)</b>'
                }
                #endregion
            }

            $mailParams.Message += "<p>Correctly processed group names:</p>"
            $mailParams.Message += $jobResults.Name | ConvertTo-HtmlListHC
            $mailParams.Message += "<p><i>Check the attachment for details.</i></p>"
        }
        else {
            $mailParams.Subject = 'FAILURE'
            $M = 'No groups processed'
            $mailParams.Message += "<p>We couldn't process any group, please check the error message.</p>"
            Write-Warning $M
            Write-EventLog @EventWarnParams -Message $M
        }

        if ($Error) {
            $mailParams.Subject = 'FAILURE'
            $mailParams.Priority = 'High'
            $Error | Get-Unique | ForEach-Object {
                Write-EventLog @EventErrorParams -Message "Error detected:`n`n$_"
            }
            $mailParams.Message += $Error | Get-Unique | ConvertTo-HtmlListHC -Spacing Wide -Header 'Errors detected:'
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
    }
    Catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}