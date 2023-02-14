﻿# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

. $PSScriptRoot\..\..\..\Shared\Invoke-CatchActionError.ps1

function Register-AuthCertificateRenewalTask {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [string]$TaskName = "Daily Auth Certificate Check",
        [string]$Username,
        [SecureString]$Password,
        [string]$WorkingDirectory,
        [string]$ScriptName,
        [bool]$IgnoreOfflineServers = $false,
        [bool]$IgnoreHybridConfig = $false,
        [ValidatePattern("^\w+([-+.']\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$")]
        [string[]]$SendEmailNotificationTo,
        [switch]$TrustAllCertificates,
        [string]$DailyRuntime = "10am",
        [string]$TaskDescription = "AutoGeneratedViaMonitorExchangeAuthCertificateScript",
        [ScriptBlock]$CatchActionFunction
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $fullPathToScript = [System.IO.Path]::Combine($WorkingDirectory, $ScriptName)
    try {
        $existingScheduledTask = Get-ScheduledTask -TaskName $($TaskName) -ErrorAction Stop | Where-Object {
            ($_.Description -eq $TaskDescription)
        }
    } catch {
        Write-Verbose ("No scheduled task with name: $($TaskName) was found - we don't need to unregister it")
        Invoke-CatchActionError $CatchActionFunction
    }

    if ($null -ne $existingScheduledTask) {
        Write-Verbose ("Scheduled task already exists - will be deleted now to re-create a new one")
        try {
            foreach ($t in $existingScheduledTask) {
                if (-not($WhatIfPreference)) {
                    Unregister-ScheduledTask -TaskPath $($t.TaskPath) -TaskName $($t.TaskName) -Confirm:$false -ErrorAction Stop
                } else {
                    Write-Host ("What if: Will unregister scheduled task: '$($t.TaskName)' by running 'Unregister-ScheduledTask'")
                }
                Write-Verbose ("Scheduled task: $($t.TaskName) successfully unregistered")
            }
        } catch {
            Write-Verbose ("The scheduled task already exists and we were unable to unregister it - Exception $($Error[0].Exception.Message)")
            Invoke-CatchActionError $CatchActionFunction
            return $false
        }
    }

    if (($WhatIfPreference) -or
        (Test-Path -Path $fullPathToScript)) {
        $passwordAsPlaintextString = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto([System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password))

        $schTaskTrigger = New-ScheduledTaskTrigger -Daily -At $DailyRuntime

        $newScheduledTaskParams = @{
            Execute          = "powershell.exe"
            WorkingDirectory = "$($WorkingDirectory)"
        }

        $basicArgumentParameters = "-NonInteractive -NoLogo -NoProfile -Command `".\$($ScriptName) -ValidateAndRenewAuthCertificate `$true -IgnoreUnreachableServers `$$($IgnoreOfflineServers) -IgnoreHybridConfig `$$($IgnoreHybridConfig)"
        if ($null -ne $SendEmailNotificationTo) {
            if ($TrustAllCertificates) {
                $newScheduledTaskParams.Add("Argument", "$($basicArgumentParameters) -SendEmailNotificationTo $([string]::Join(", ", $SendEmailNotificationTo)) -TrustAllCertificates -Confirm:`$false`"")
            } else {
                $newScheduledTaskParams.Add("Argument", "$($basicArgumentParameters) -SendEmailNotificationTo $([string]::Join(", ", $SendEmailNotificationTo)) -Confirm:`$false`"")
            }
        } else {
            $newScheduledTaskParams.Add("Argument", "$($basicArgumentParameters) -Confirm:`$false`"")
        }

        $schTaskAction = New-ScheduledTaskAction @newScheduledTaskParams

        $registerSchTaskParams = @{
            TaskName    = $TaskName
            Trigger     = $schTaskTrigger
            Action      = $schTaskAction
            Description = $TaskDescription
            RunLevel    = "Highest"
            User        = $Username
            Password    = $passwordAsPlaintextString
            Force       = $true
            ErrorAction = "Stop"
        }

        try {
            Write-Verbose ("Scheduled Task: $($TaskName) successfully created")
            if (-not($WhatIfPreference)) {
                Register-ScheduledTask @registerSchTaskParams | Out-Null
            } else {
                Write-Host ("What if: Registering scheduled task with name '$($TaskName)' by running 'Register-ScheduledTask'")
            }
            return $true
        } catch {
            Write-Verbose ("Error while creating Scheduled Task: $($TaskName) - Exception: $($Error[0].Exception.Message)")
            Invoke-CatchActionError $CatchActionFunction
        }
    } else {
        Write-Verbose ("Script: $($fullPathToScript) doesn't exist")
    }

    return $false
}
