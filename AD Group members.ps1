#Requires -Version 5.1
#Requires -Modules ActiveDirectory, ImportExcel
#Requires -Modules Toolbox.ActiveDirectory, Toolbox.HTML, Toolbox.EventLog
#Requires -Modules Toolbox.Remoting

<#
    .SYNOPSIS
        Get the members of a group in AD and send the result in an e-mail to 
        the users defined in 'MailTo'.
        
    .DESCRIPTION
        Get the members of a group in active directory and generate one object
        per end node. The result is saved in an Excel file that will be mailed 
        in attachment to the end user. Two worksheets per security group are 
        generated, one storing the group structure with its members and one 
        storing the user members in a single column. The log folder will store 
        the e-mail sent to the end user, a copy of the import file and the 
        Excel file containing the results. All actions are also stored in the 
        Windows Event Log.

    .PARAMETER GroupNames
        Can be one or more AD active directory group names.

    .PARAMETER MailTo
        SMTP mail addresses

    .PARAMETER MaxThreads
        Maximum number of jobs allowed to run at the same time.

    .NOTES
        CHANGELOG
        2017/06/06 Script born
        2017/07/05 Fixed bug where we tried to attach an attachment when the group had no members
                   Added color to Excel cell in case it contains an asterisk symbol
                   Added bold text in the email for circular group membership
        2017/08/23 Added second sheet to also display user members in a single column
        2018/03/20 Removed parameter 'Mode'
        2020/10/13 Removed ImportFile parameter as we take plain arguments now
        2020/10/21 Adjusted catch clause to simply throw

        AUTHOR Brecht.Gijbels@heidelbergcement.com #>

[CmdLetBinding()]
Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String[]]$MailTo,
    [Parameter(Mandatory)]
    [String[]]$GroupNames,
    [ValidateRange(1, 7)]
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
        $Jobs = $JobResults = @()

        $Init = { Import-Module Toolbox.ActiveDirectory }

        ForEach ($G in ($GroupNames | Sort-Object -Property @{E = { $_.Trim() } } -Unique)) {
            $Jobs += Start-Job -Name $ScriptName -InitializationScript $Init -ScriptBlock {
                Param (
                    $Group
                )

                if (-not ($MembersFlat = Get-ADGroupMemberFlatHC -Identity $Group)) {
                    $MembersFlat = $null
                }
                
                if (-not ($MemberUsers = Get-ADGroupMember -Identity $Group -Recursive | 
                        Select-Object @{L = $Group; E = { $_.Name } })) {
                    $MemberUsers = $null
                }

                [PSCustomObject]@{
                    Name        = $Group
                    MembersFlat = $MembersFlat
                    MemberUsers = $MemberUsers
                }
            } -ArgumentList $G
            Write-Verbose "Job started for group '$G'"
            Wait-MaxRunningJobsHC -Name $Jobs -MaxThreads $MaxThreads
        }

        if ($Jobs) {
            $JobResults = $Jobs | Wait-Job | Receive-Job
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
        $ExcelParams = @{
            Path            = $LogFile + ' - Result.xlsx'
            AutoSize        = $true
            BoldTopRow      = $true
            FreezeTopRow    = $true
            AutoFilter      = $true
            ConditionalText = $(
                New-ConditionalText ~* Black Orange
            )
        }

        $MailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = 'Success'
            Message   = 'This report checks security groups for their members.' 
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = $LogFile + ' - Mail.html'
        }

        if ($JobResults) {
            $Sheet = 0
            $JobResults | ForEach-Object {
                if ($_.MembersFlat) {
                    $M = "Group '$($_.Name)' has $($_.MembersFlat.Count) end node members"
                    Write-Verbose $M
                    Write-EventLog @EventVerboseParams -Message "Import file '$($ImportFile.FullName)':`n`n- $M"
                    
                    #region Export to Excel
                    Write-Verbose "Export to Excel file '$($ExcelParams.Path)'"
                    $Sheet++ # Excel worksheet names can only be 31 chars long and must be unique
                    $_.MembersFlat | Sort-Object -Property * | Update-FirstObjectProperties | 
                    Export-Excel @ExcelParams -WorksheetName ("$Sheet (Flat) " + $_.Name)
                    if ($_.MemberUsers) {
                        $_.MemberUsers | Sort-Object -Property * | 
                        Export-Excel @ExcelParams -WorksheetName ("$Sheet (Users) " + $_.Name)
                    }
                    else {
                        $_.Name | Export-Excel @ExcelParams -WorksheetName ("$Sheet (Users) " + $_.Name)
                    }
                    #endregion

                    #region Check circular group membership
                    foreach ($G in $_.MembersFlat) {
                        if ($G.PSObject.Members | Where-Object MemberType -EQ NoteProperty | 
                            Where-Object Value -Match '\*') {
                            $CircularGroup = $true
                            break
                        }
                    }
                    #endregion

                    $MailParams.Attachments = $ExcelParams.Path
                }
                else {
                    $NoGroupMembers = $true
                }

                #region format group names
                if ($CircularGroup) {
                    $CircularGroup = $false

                    $M = "Group '$($_.Name)' has circular group membership"
                    Write-Warning $M
                    Write-EventLog @EventWarnParams -Message "Import file '$($ImportFile.FullName)':`n`n- $M"
                    $_.Name += ' <b>(Circular group membership found)</b>'
                }

                if ($NoGroupMembers) {
                    $NoGroupMembers = $false

                    $M = "Group '$($_.Name)' has no end node members"
                    Write-Warning $M
                    Write-EventLog @EventWarnParams -Message "Import file '$($ImportFile.FullName)':`n`n- $M"
                    $_.Name += ' <b>(No members found)</b>'
                }
                #endregion
            }

            $MailParams.Message += "<p>Correctly processed group names:</p>"
            $MailParams.Message += $JobResults.Name | ConvertTo-HtmlListHC
            $MailParams.Message += "<p><i>Check the attachment for details.</i></p>"
        }
        else {
            $MailParams.Subject = 'FAILURE'
            $M = 'No groups processed'
            $MailParams.Message += "<p>We couldn't process any group, please check the error message.</p>"
            Write-Warning $M
            Write-EventLog @EventWarnParams -Message $M
        }

        if ($Error) {
            $MailParams.Subject = 'FAILURE'
            $MailParams.Priority = 'High'
            $Error | Get-Unique | ForEach-Object { 
                Write-EventLog @EventErrorParams -Message "Error detected:`n`n$_"
            }
            $MailParams.Message += $Error | Get-Unique | ConvertTo-HtmlListHC -Spacing Wide -Header 'Errors detected:'
        }

        Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
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