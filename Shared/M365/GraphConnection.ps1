# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

. $PSScriptRoot\..\ModuleHandle.ps1

function Connect-GraphAdvanced {
    #[CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Scopes,
        [Parameter(Mandatory = $true)]
        [string[]]$Modules,
        [Parameter(Mandatory = $false)]
        [switch]$DoNotShowConnectionDetails
    )

<#     if ($PSVersionTable.PSVersion.Major -le 5 ) {
        if ($MaximumFunctionCount -lt 7500) {
            $MaximumFunctionCount += 7500
        }
        if ($MaximumVariableCount -lt 7500) {
            $MaximumVariableCount += 7500
        }
    }
 #>
    #Validate Graph is installed and loaded
    $module = $false
    foreach ($module in $Modules) {
        $module = Request-Module -ModuleName $module
        if (-not $module) {
            Write-Host "We cannot continue without Microsoft.Graph Powershell module" -ForegroundColor Red
            return $null
        }
    }

    #Validate Graph is connected or try to connect
    $connection = $null
    $connection = Get-MgContext -ErrorAction SilentlyContinue
    if ($null -eq $connection) {
        Write-Host "Not connected to Graph" -ForegroundColor Yellow
        if ($PSCmdlet.ShouldContinue("Do you want to connect?", "We need a Graph connection")) {
            Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction SilentlyContinue
            $connection = Get-MgContext -ErrorAction SilentlyContinue
            if ($null -eq $connection) {
                Write-Host "Connection could not be established" -ForegroundColor Red
            } else {
                if (Test-GraphContext -Scopes $connection.Scopes -ExpectedScopes $Scopes) {
                    $connection.PSObject.Properties | ForEach-Object { Write-Verbose "$($_.Name): $($_.Value)" }
                    if (-not $DoNotShowConnectionDetails) {
                        Show-GraphContext -Context $connection
                    }
                    return $connection
                } else {
                    Write-Host "We cannot continue without Graph Powershell session non Expected Scopes found" -ForegroundColor Red
                }
            }
        } else {
            Write-Host "We cannot continue without Graph Powershell session" -ForegroundColor Red
        }
    } elseif ($connection.count -eq 1) {
        Write-Verbose "You have a Graph sessions"
        if (Test-GraphContext -Scopes $connection.Scopes -ExpectedScopes $Scopes) {
            $connection.PSObject.Properties | ForEach-Object { Write-Verbose "$($_.Name): $($_.Value)" }
            if (-not $DoNotShowConnectionDetails) {
                Show-GraphContext -Context $connection
            }
            return $connection
        } else {
            Write-Host "We cannot continue without Graph Powershell session non Expected Scopes found" -ForegroundColor Red
        }
    } else {
        Write-Host "You have more than one Graph sessions please use just one session" -ForegroundColor Red
    }
    return $null
}

function Test-GraphContext {
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Scopes,
        [Parameter(Mandatory = $true)]
        [string[]]$ValidScopes
    )

    $missingScopes = Compare-Object -ReferenceObject $ExpectedScopes -DifferenceObject $Scopes

    if ($missingScopes) {
        Write-Host "The following scopes are missing: $($missingScopes | ForEach-Object { $_.InputObject })" -ForegroundColor Red
        return $false
    } else {
        Write-Verbose "All expected scopes are present." -ForegroundColor Green
        return $true
    }
}

function Show-GraphContext {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Authentication.AuthContext]$Context
    )
    Write-Host "Connected to Graph"
    Write-Host "Session details"
    Write-Host "Tenant Id: $($Context.TenantId)"
    Write-Host "User: $($Context.Account)"
}
