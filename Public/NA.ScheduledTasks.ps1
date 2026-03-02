function Normalize-NATaskPath {
    [CmdletBinding()]
    param(
        [string]$TaskPath
    )

    if ([string]::IsNullOrWhiteSpace($TaskPath)) {
        return '\'
    }

    $normalized = $TaskPath.Trim() -replace '/', '\'
    $normalized = $normalized.Trim('\')

    if ([string]::IsNullOrWhiteSpace($normalized)) {
        return '\'
    }

    return "\$normalized\"
}


function Register-ScriptScheduledTask {
    <#
    .SYNOPSIS
        Registers a scheduled task that runs a PowerShell script.
    .DESCRIPTION
        Creates a scheduled task in either Standard mode (action/trigger objects)
        or Xml mode (custom XML content/file for advanced scheduling scenarios).
    .PARAMETER TaskName
        Name of the scheduled task.
    .PARAMETER TaskPath
        Task Scheduler path (for example: \ or \Nebula\).
    .PARAMETER Mode
        Scheduling mode: Standard or Xml.
    .PARAMETER ScriptPath
        Script file to execute in Standard mode.
    .PARAMETER PwshPath
        PowerShell executable path in Standard mode. Default: pwsh.exe.
    .PARAMETER ScriptArguments
        Additional script arguments appended after -File in Standard mode.
    .PARAMETER WorkingDirectory
        Working directory used when starting PowerShell in Standard mode.
    .PARAMETER ExecutionPolicy
        Execution policy used in Standard mode. Default: Bypass.
    .PARAMETER StartTime
        Start time for the trigger in Standard mode. Default: now + 5 minutes.
    .PARAMETER ScheduleType
        Trigger type for Standard mode: Daily or Once.
    .PARAMETER RepetitionIntervalMinutes
        Optional repetition interval (minutes) for Standard mode.
    .PARAMETER RepetitionDurationHours
        Optional repetition duration (hours) for Standard mode.
    .PARAMETER TaskXml
        XML content used in Xml mode.
    .PARAMETER TaskXmlPath
        XML file path used in Xml mode.
    .PARAMETER Description
        Optional task description.
    .PARAMETER Credential
        Optional account credential used to run the task.
    .PARAMETER Force
        Recreate the task if it already exists.
    .PARAMETER RunElevated
        Register task with Highest privileges.
    .EXAMPLE
        Register-ScriptScheduledTask -TaskName "MyScript" -ScriptPath "C:\Scripts\job.ps1" -StartTime (Get-Date).Date.AddHours(3)
    .EXAMPLE
        Register-ScriptScheduledTask -TaskName "MyCustomTask" -Mode Xml -TaskXmlPath "C:\Temp\task.xml" -Credential (Get-Credential)
    .EXAMPLE
        Register-ScriptScheduledTask -TaskName "HourlyJob" -TaskPath "\Nebula\" -ScriptPath "C:\Scripts\job.ps1" -ScheduleType Once -StartTime (Get-Date).AddMinutes(30) -Force
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/register-scriptscheduledtask
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TaskName,

        [ValidateNotNullOrEmpty()]
        [string]$TaskPath = '\',

        [ValidateSet('Standard', 'Xml')]
        [string]$Mode = 'Standard',

        [string]$ScriptPath,
        [string]$PwshPath = 'pwsh.exe',
        [string]$ScriptArguments,
        [string]$WorkingDirectory,

        [ValidateSet('Bypass', 'RemoteSigned', 'AllSigned', 'Unrestricted', 'Restricted', 'Undefined', 'Default')]
        [string]$ExecutionPolicy = 'Bypass',

        [datetime]$StartTime = (Get-Date).AddMinutes(5),

        [ValidateSet('Daily', 'Once')]
        [string]$ScheduleType = 'Daily',

        [ValidateRange(0, 1440)]
        [int]$RepetitionIntervalMinutes = 0,

        [ValidateRange(0, 744)]
        [int]$RepetitionDurationHours = 0,

        [string]$TaskXml,
        [string]$TaskXmlPath,

        [string]$Description,

        [System.Management.Automation.PSCredential]$Credential,

        [switch]$RunElevated,
        [switch]$Force,

        [string]$LogLocation
    )

    $result = [pscustomobject]@{
        Success  = $false
        TaskName = $TaskName
        TaskPath = $TaskPath
        Mode     = $Mode
        Message  = $null
    }

    try {
        $TaskPath = Normalize-NATaskPath -TaskPath $TaskPath
        $result.TaskPath = $TaskPath

        if ($Mode -eq 'Standard') {
            if ([string]::IsNullOrWhiteSpace($ScriptPath)) {
                throw 'In Standard mode, ScriptPath is mandatory.'
            }
            if (-not (Test-Path -LiteralPath $ScriptPath)) {
                throw "ScriptPath not found: $ScriptPath"
            }
        }

        if ($Mode -eq 'Xml') {
            $hasXml = -not [string]::IsNullOrWhiteSpace($TaskXml)
            $hasXmlPath = -not [string]::IsNullOrWhiteSpace($TaskXmlPath)
            if (-not $hasXml -and -not $hasXmlPath) {
                throw 'In Xml mode, specify TaskXml or TaskXmlPath.'
            }
            if ($hasXml -and $hasXmlPath) {
                throw 'Specify only one between TaskXml and TaskXmlPath.'
            }
            if ($hasXmlPath -and -not (Test-Path -LiteralPath $TaskXmlPath)) {
                throw "TaskXmlPath not found: $TaskXmlPath"
            }
        }

        $existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
        if ($existingTask) {
            if (-not $Force.IsPresent) {
                throw "Task '$TaskPath$TaskName' already exists. Use -Force to recreate it."
            }
            if ($PSCmdlet.ShouldProcess("$TaskPath$TaskName", 'Unregister existing scheduled task')) {
                Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
                Write-NALog -Message "Existing task '$TaskPath$TaskName' removed because -Force was specified." -Level INFO -LogLocation $LogLocation
            }
        }

        $runLevel = if ($RunElevated.IsPresent) { 'Highest' } else { 'Limited' }

        if ($Mode -eq 'Standard') {
            $resolvedScriptPath = (Resolve-Path -LiteralPath $ScriptPath).Path
            $effectiveWorkingDirectory = if ([string]::IsNullOrWhiteSpace($WorkingDirectory)) {
                Split-Path -Path $resolvedScriptPath -Parent
            } else {
                $WorkingDirectory
            }

            $psArgs = @(
                '-NoProfile',
                '-WindowStyle', 'Hidden',
                '-ExecutionPolicy', $ExecutionPolicy,
                '-WorkingDirectory', ('"{0}"' -f $effectiveWorkingDirectory),
                '-File', ('"{0}"' -f $resolvedScriptPath)
            )

            if (-not [string]::IsNullOrWhiteSpace($ScriptArguments)) {
                $psArgs += $ScriptArguments
            }

            $action = New-ScheduledTaskAction -Execute $PwshPath -Argument ($psArgs -join ' ')
            if ($ScheduleType -eq 'Daily') {
                $trigger = New-ScheduledTaskTrigger -Daily -At $StartTime
            } else {
                $trigger = New-ScheduledTaskTrigger -Once -At $StartTime
            }

            if ($RepetitionIntervalMinutes -gt 0) {
                $trigger.Repetition.Interval = "PT$($RepetitionIntervalMinutes)M"
                if ($RepetitionDurationHours -gt 0) {
                    $trigger.Repetition.Duration = "PT$($RepetitionDurationHours)H"
                }
            }

            $settingsParams = @{
                AllowStartIfOnBatteries = $true
                DontStopIfGoingOnBatteries = $true
                StartWhenAvailable = $true
            }
            $settings = New-ScheduledTaskSettingsSet @settingsParams

            $registerParams = @{
                TaskName    = $TaskName
                TaskPath    = $TaskPath
                Action      = $action
                Trigger     = $trigger
                Settings    = $settings
                RunLevel    = $runLevel
                ErrorAction = 'Stop'
            }
            if (-not [string]::IsNullOrWhiteSpace($Description)) {
                $registerParams.Description = $Description
            }
            if ($Credential) {
                $registerParams.User = $Credential.UserName
                $registerParams.Password = $Credential.GetNetworkCredential().Password
            }

            if ($PSCmdlet.ShouldProcess("$TaskPath$TaskName", 'Register scheduled task (Standard mode)')) {
                Register-ScheduledTask @registerParams | Out-Null
            }
        }
        else {
            $xmlContent = if (-not [string]::IsNullOrWhiteSpace($TaskXmlPath)) {
                Get-Content -LiteralPath $TaskXmlPath -Raw -ErrorAction Stop
            } else {
                $TaskXml
            }

            $registerParams = @{
                TaskName    = $TaskName
                TaskPath    = $TaskPath
                Xml         = $xmlContent
                ErrorAction = 'Stop'
            }
            if ($Credential) {
                $registerParams.User = $Credential.UserName
                $registerParams.Password = $Credential.GetNetworkCredential().Password
            }

            if ($PSCmdlet.ShouldProcess("$TaskPath$TaskName", 'Register scheduled task (Xml mode)')) {
                Register-ScheduledTask @registerParams | Out-Null
            }
        }

        $result.Success = $true
        $result.Message = "Task '$TaskPath$TaskName' registered successfully (Mode: $Mode)."
        Write-NALog -Message $result.Message -Level SUCCESS -LogLocation $LogLocation
    }
    catch {
        $result.Message = $_.Exception.Message
        Write-NALog -Message "Failed to register task '$TaskPath$TaskName': $($result.Message)" -Level ERROR -LogLocation $LogLocation
    }

    return $result
}


function Unregister-ScriptScheduledTask {
    <#
    .SYNOPSIS
        Unregisters a scheduled script task.
    .DESCRIPTION
        Removes an existing task by name/path and returns a status object.
    .PARAMETER TaskName
        Name of the scheduled task.
    .PARAMETER TaskPath
        Task Scheduler path (for example: \ or \Nebula\).
    .EXAMPLE
        Unregister-ScriptScheduledTask -TaskName "MyScript"
    .EXAMPLE
        Unregister-ScriptScheduledTask -TaskName "MyScript" -TaskPath "\Nebula\"
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/unregister-scriptscheduledtask
    #>
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TaskName,

        [ValidateNotNullOrEmpty()]
        [string]$TaskPath = '\',

        [string]$LogLocation
    )

    $result = [pscustomobject]@{
        Success  = $false
        TaskName = $TaskName
        TaskPath = $TaskPath
        Message  = $null
    }

    try {
        $TaskPath = Normalize-NATaskPath -TaskPath $TaskPath
        $result.TaskPath = $TaskPath

        $existingTask = Get-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -ErrorAction SilentlyContinue
        if (-not $existingTask) {
            $result.Success = $true
            $result.Message = "Task '$TaskPath$TaskName' not found. Nothing to remove."
            Write-NALog -Message $result.Message -Level INFO -LogLocation $LogLocation
            return $result
        }

        if ($PSCmdlet.ShouldProcess("$TaskPath$TaskName", 'Unregister scheduled task')) {
            Unregister-ScheduledTask -TaskName $TaskName -TaskPath $TaskPath -Confirm:$false -ErrorAction Stop
        }

        $result.Success = $true
        $result.Message = "Task '$TaskPath$TaskName' unregistered successfully."
        Write-NALog -Message $result.Message -Level SUCCESS -LogLocation $LogLocation
    }
    catch {
        $result.Message = $_.Exception.Message
        Write-NALog -Message "Failed to unregister task '$TaskPath$TaskName': $($result.Message)" -Level ERROR -LogLocation $LogLocation
    }

    return $result
}


function Invoke-ScriptTaskLifecycle {
    <#
    .SYNOPSIS
        Orchestrates script scheduled-task registration or unregistration.
    .DESCRIPTION
        Wrapper helper that delegates to Register-ScriptScheduledTask and
        Unregister-ScriptScheduledTask and optionally handles credential prompt
        and HH:mm task-time parsing for Standard mode scheduling.
    .PARAMETER RegisterTask
        Register the task.
    .PARAMETER UnregisterTask
        Unregister the task.
    .PARAMETER TaskName
        Task name.
    .PARAMETER TaskPath
        Task path.
    .PARAMETER ScriptPath
        Script path to execute when registering.
    .PARAMETER TaskTime
        Daily task time in HH:mm format.
    .PARAMETER ScheduleType
        Daily or Once scheduling mode.
    .PARAMETER Description
        Optional task description.
    .PARAMETER PwshPath
        PowerShell executable path.
    .PARAMETER WorkingDirectory
        Working directory for script execution.
    .PARAMETER ExecutionPolicy
        Execution policy used by scheduled task action.
    .PARAMETER PromptForCredential
        Ask interactively for credentials when Credential is not provided.
    .PARAMETER DefaultUserName
        Username proposed during interactive credential prompt.
    .PARAMETER Credential
        Credential used for task registration.
    .PARAMETER Force
        Recreate task if already present.
    .PARAMETER LogLocation
        Optional log location.
    .EXAMPLE
        Invoke-ScriptTaskLifecycle -RegisterTask -TaskName "MyTask" -TaskPath "\Nebula\" -ScriptPath "C:\Scripts\Job.ps1" -TaskTime "02:00" -PromptForCredential
    .EXAMPLE
        Invoke-ScriptTaskLifecycle -UnregisterTask -TaskName "MyTask" -TaskPath "\Nebula\"
    .EXAMPLE
        Invoke-ScriptTaskLifecycle -RegisterTask -TaskName "MyTask" -TaskPath "\" -ScriptPath "C:\Scripts\Job.ps1" -TaskTime "3:15" -Force
    .LINK
        https://kb.gioxx.org/Nebula/Automations/usage/invoke-scripttasklifecycle
    #>
    [CmdletBinding(DefaultParameterSetName = 'Register')]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = 'Register')]
        [switch]$RegisterTask,

        [Parameter(Mandatory = $true, ParameterSetName = 'Unregister')]
        [switch]$UnregisterTask,

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$TaskName,

        [ValidateNotNullOrEmpty()]
        [string]$TaskPath = '\',

        [Parameter(Mandatory = $true, ParameterSetName = 'Register')]
        [ValidateNotNullOrEmpty()]
        [string]$ScriptPath,

        [string]$TaskTime = '02:00',

        [ValidateSet('Daily', 'Once')]
        [string]$ScheduleType = 'Daily',

        [string]$Description,
        [string]$PwshPath,
        [string]$WorkingDirectory,

        [ValidateSet('Bypass', 'RemoteSigned', 'AllSigned', 'Unrestricted', 'Restricted', 'Undefined', 'Default')]
        [string]$ExecutionPolicy = 'Bypass',

        [switch]$PromptForCredential,
        [Alias('DefaultCredentialUser')]
        [string]$DefaultUserName = "$Env:UserDomain\$Env:UserName",
        [System.Management.Automation.PSCredential]$Credential,

        [switch]$Force,
        [string]$LogLocation
    )

    if ($PSCmdlet.ParameterSetName -eq 'Unregister') {
        $TaskPath = Normalize-NATaskPath -TaskPath $TaskPath
        return Unregister-ScriptScheduledTask -TaskName $TaskName -TaskPath $TaskPath -LogLocation $LogLocation
    }

    $TaskPath = Normalize-NATaskPath -TaskPath $TaskPath

    if ([string]::IsNullOrWhiteSpace($PwshPath)) {
        $PwshPath = Get-Command -Name 'pwsh.exe' -ErrorAction SilentlyContinue |
            Select-Object -First 1 -ExpandProperty Source
        if ([string]::IsNullOrWhiteSpace($PwshPath)) {
            $PwshPath = "$Env:ProgramFiles\PowerShell\7\pwsh.exe"
        }
    }

    if (-not (Test-Path -LiteralPath $PwshPath)) {
        return [pscustomobject]@{
            Success  = $false
            TaskName = $TaskName
            TaskPath = $TaskPath
            Action   = 'Register'
            Message  = "PowerShell executable not found: $PwshPath"
        }
    }

    $normalizedTaskTime = if ($null -eq $TaskTime) { '' } else { $TaskTime.Trim() }
    $taskTimeMatch = [regex]::Match($normalizedTaskTime, '^(?<hour>(?:[01]?\d|2[0-3])):(?<minute>[0-5]\d)$')
    if (-not $taskTimeMatch.Success) {
        return [pscustomobject]@{
            Success  = $false
            TaskName = $TaskName
            TaskPath = $TaskPath
            Action   = 'Register'
            Message  = "TaskTime '$TaskTime' is not valid. Expected format: h:mm or HH:mm"
        }
    }
    $hour = [int]$taskTimeMatch.Groups['hour'].Value
    $minute = [int]$taskTimeMatch.Groups['minute'].Value
    $startTime = [datetime]::Today.AddHours($hour).AddMinutes($minute)

    $effectiveCredential = $Credential
    if (-not $effectiveCredential -and $PromptForCredential.IsPresent) {
        $effectiveCredential = Get-Credential -Message 'Enter user credentials (to execute the scheduled operation)' -UserName $DefaultUserName
        if (-not $effectiveCredential) {
            return [pscustomobject]@{
                Success  = $false
                TaskName = $TaskName
                TaskPath = $TaskPath
                Action   = 'Register'
                Message  = 'Credential prompt canceled by user.'
            }
        }
    }

    $registerParams = @{
        TaskName        = $TaskName
        TaskPath        = $TaskPath
        Mode            = 'Standard'
        ScriptPath      = $ScriptPath
        PwshPath        = $PwshPath
        ExecutionPolicy = $ExecutionPolicy
        StartTime       = $startTime
        ScheduleType    = $ScheduleType
        LogLocation     = $LogLocation
    }

    if (-not [string]::IsNullOrWhiteSpace($Description)) {
        $registerParams.Description = $Description
    }
    if (-not [string]::IsNullOrWhiteSpace($WorkingDirectory)) {
        $registerParams.WorkingDirectory = $WorkingDirectory
    }
    if ($effectiveCredential) {
        $registerParams.Credential = $effectiveCredential
    }
    if ($Force.IsPresent) {
        $registerParams.Force = $true
    }

    return Register-ScriptScheduledTask @registerParams
}

