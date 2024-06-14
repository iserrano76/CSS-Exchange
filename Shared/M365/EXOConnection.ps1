# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

. $PSScriptRoot\..\ModuleHandle.ps1

function Connect-EXOAdvanced {
    param (
        [Parameter(Mandatory = $false, ParameterSetName = 'SingleSession')]
        [Parameter(Mandatory = $false, ParameterSetName = 'AllowMultipleSessions')]
        [switch]$DoNotShowConnectionDetails,
        [Parameter(Mandatory = $true, ParameterSetName = 'AllowMultipleSessions')]
        [switch]$AllowMultipleSessions,
        [Parameter(Mandatory = $true, ParameterSetName = 'AllowMultipleSessions')]
        [string]$Prefix,
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    #Validate EXO 3.0 is installed and loaded
    $module = $false
    $module = Request-Module -ModuleName "ExchangeOnlineManagement" -MinModuleVersion 3.0.0

    if (-not $module) {
        Write-Host "We cannot continue without ExchangeOnlineManagement Powershell module" -ForegroundColor Red
        return $null
    }

    #Validate EXO is connected or try to connect
    $connections = $null
    $connections = Get-ConnectionInformation -ErrorAction SilentlyContinue
    if ($null -eq $connections) {
        Write-Host "Not connected to Exchange Online" -ForegroundColor Yellow
        if ($Force -or $PSCmdlet.ShouldContinue("Do you want to connect?", "No connection found. We need a ExchangeOnlineManagement connection")) {
            if ($AllowMultipleSessions) {
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction SilentlyContinue -Prefix $Prefix
            } else {
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction SilentlyContinue
            }
            $connection = Get-ConnectionInformation -ErrorAction SilentlyContinue
            if ($null -eq $connection) {
                Write-Host "Connection could not be established" -ForegroundColor Red
            } else {
                $connection.PSObject.Properties | ForEach-Object { Write-Verbose "$($_.Name): $($_.Value)" }
                if (-not $DoNotShowConnectionDetails) {
                    Show-EXOConnection -Connection $connection
                }
                return $connection
            }
        }
        Write-Host "We cannot continue without ExchangeOnlineManagement Powershell session" -ForegroundColor Red
    } else {
        if ($AllowMultipleSessions) {
            Write-Verbose "Found Exchange Online sessions"
            foreach ($connection in $connections) {
                Write-Verbose " "
                $connection.PSObject.Properties | ForEach-Object { Write-Verbose "$($_.Name): $($_.Value)" }
                if (-not $DoNotShowConnectionDetails) {
                    Show-EXOConnection -Connection $connection
                }
            }
            if ($Force -or $PSCmdlet.ShouldContinue("Do you want another connection?", "You have Exchange Online sessions")) {
                Connect-ExchangeOnline -ShowBanner:$false -ErrorAction SilentlyContinue -Prefix $Prefix
                $newconnections = Get-ConnectionInformation -ErrorAction SilentlyContinue
                $newSession = $null
                $newSession = $newconnections | Where-Object { $_ -notin $connections }
                if ($null -eq $newSession) {
                    Write-Host "Connection could not be established" -ForegroundColor Red
                } else {
                    $connection.PSObject.Properties | ForEach-Object { Write-Verbose "$($_.Name): $($_.Value)" }
                    if (-not $DoNotShowConnectionDetails) {
                        Show-EXOConnection -Connection $newSession
                    }
                    return $newSession
                }
            }
        } else {
            if ($connections.count -eq 1) {
                Write-Verbose "You have a Exchange Online session"
                Write-Verbose " "
                $connection.PSObject.Properties | ForEach-Object { Write-Verbose "$($_.Name): $($_.Value)" }
                if (-not $DoNotShowConnectionDetails) {
                    Show-EXOConnection -Connection $connections
                }
                return $connections
            } else {
                Write-Host "You have more than one Exchange Online sessions please use just one session" -ForegroundColor Red
            }
        }
    }
    return $null
}

function Show-EXOConnection {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Exchange.Management.ExoPowershellSnapin.ConnectionInformation]$Connection
    )
    Write-Host "Connected to Exchange Online"
    Write-Host "Session details"
    Write-Host "Tenant Id: $($Connection.TenantId)"
    Write-Host "User: $($Connection.UserPrincipalName)"
}
