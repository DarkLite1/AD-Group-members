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
param (
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

begin {
    try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        $Error.Clear()

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
        catch {
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
    catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
}

process {
    try {
        $jobs = $jobResults = @()

        $init = { Import-Module Toolbox.ActiveDirectory }

        foreach (
            $group in
            (
                $adGroupNames |
                Sort-Object -Property @{Expression = { $_.Trim() } } -Unique
            )
        ) {
            Write-Verbose "Start job for group '$group'"

            $jobs += Start-Job -Name $ScriptName -InitializationScript $init -ScriptBlock {
                param (
                    $Group
                )

                if (-not
                    ($MembersFlat = Get-ADGroupMemberFlatHC -Identity $Group)
                ) {
                    $MembersFlat = [PSCustomObject]@{
                        GroupName = $Group
                        'Member1' = $null
                        'Member2' = $null
                        'Member3' = $null
                        'Member4' = $null
                        'Member5' = $null
                        'Member6' = $null
                    }
                }

                if (-not
                    (
                        $MemberUsers = Get-ADGroupMember -Identity $Group -Recursive | Where-Object {
                            $_.objectClass -eq 'user'
                        } |
                        Get-ADUser -Property EmailAddress |
                        Select-Object @{
                            Name       = 'GroupName'
                            Expression = { $Group }
                        },
                        @{
                            Name       = 'MemberUserName'
                            Expression = { $_.Name }
                        },
                        @{
                            Name       = 'MemberUserEmailAddress'
                            Expression = { $_.EmailAddress }
                        }
                    )
                ) {
                    $MemberUsers = [PSCustomObject]@{
                        GroupName      = $Group
                        MemberUserName = [PSCustomObject]@{
                            GroupName                = $Group
                            'MemberUserName'         = $null
                            'MemberUserEmailAddress' = $null
                        }
                    }
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
            Write-Verbose 'Jobs done'
        }
    }
    catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
    finally {
        Get-Job | Remove-Job -Force
    }
}

end {
    try {
        $mailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = '{0} AD Groups' -f $adGroupNames.Count
            Message   = 'Check AD groups for their members.'
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
        }

        $circularGroups = @()

        if ($jobResults) {
            $excelSheet = @{
                hierarchicalGroupMembership = @()
                usersInGroup                = @()
            }

            foreach ($group in $jobResults) {
                #region Get data to export to Excel
                $excelSheet.hierarchicalGroupMembership += $group.MembersFlat
                $excelSheet.usersInGroup += $group.MemberUsers
                #endregion

                #region Check circular group membership
                if (
                    $group.MembersFlat.PSObject.Members |
                    Where-Object MemberType -EQ Property |
                    Where-Object Value -Match '\*'
                ) {
                    $circularGroups += $group.Name
                    break
                }
                #endregion
            }

            #region Export to Excel file
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

            Write-Verbose "Export to Excel file '$($excelParams.Path)'"

            if ($excelSheet.usersInGroup) {
                $excelSheet.usersInGroup | Sort-Object -Property * |
                Export-Excel @excelParams -WorksheetName 'usersInGroup'

                $mailParams.Attachment = $excelParams.Path
            }

            if ($excelSheet.hierarchicalGroupMembership) {
                $excelSheet.hierarchicalGroupMembership |
                Sort-Object -Property * |
                Update-FirstObjectProperties |
                Export-Excel @excelParams -WorksheetName 'hierarchicalGroupMembership'

                $mailParams.Attachment = $excelParams.Path
            }
            #endregion
        }

        #region Add circular group names to email
        if ($circularGroups) {
            $M = "Found $($circularGroups.Count) groups with circular group memberships"
            Write-Warning $M
            Write-EventLog @EventWarnParams -Message $M

            $mailParams.Message += "<p>Found $($circularGroups.Count) groups with <b>circular group memberships</b>:<ul>{0}</ul></p>" -f $(
                $circularGroups | ForEach-Object { "<li>$_</li>" }
            )
        }
        #endregion

        $mailParams.Message += '<p>Group names:</p>'
        $mailParams.Message += $jobResults.Name | ConvertTo-HtmlListHC

        if ($mailParams.Attachment) {
            $mailParams.Message += '<p><i>Check the attachment for details.</i></p>'
        }

        if ($Error) {
            $mailParams.Subject += ", $($Error.Count) errors"
            $mailParams.Priority = 'High'
            $Error | Get-Unique | ForEach-Object {
                Write-EventLog @EventErrorParams -Message "Error detected:`n`n$_"
            }
            $mailParams.Message += $Error | Get-Unique | ConvertTo-HtmlListHC -Spacing Wide -Header 'Errors detected:'
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @mailParams
    }
    catch {
        Write-Warning $_
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        $errorMessage = $_; $global:error.RemoveAt(0); throw $errorMessage
    }
    finally {
        Write-EventLog @EventEndParams
    }
}