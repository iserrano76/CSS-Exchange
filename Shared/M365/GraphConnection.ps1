﻿# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

<#
.SYNOPSIS
This script defines a function `Connect-GraphAdvanced` that establishes a connection to Microsoft Graph.
It ensures that the required modules are installed and loaded.
The function accepts a list of scopes and modules, with optional parameters for tenant ID and connection details display.
If the required modules are not found, the script attempts to install them.
The function returns the connection information or null if the connection fails.

.PARAMETER Scopes
 Mandatory array of strings specifying the scopes for the connection.
.PARAMETER Modules
 Mandatory array of strings specifying the modules required for the connection.
.PARAMETER TenantId
 Optional array of strings specifying the tenant ID(s) for the connection.
.PARAMETER DoNotShowConnectionDetails
 Optional switch to hide connection details.
.PARAMETER MinModuleVersion
 Optional parameter to specify the minimum version of the Graph modules (default and minimum supported version 2.0.0).

.OUTPUTS
Microsoft.Graph.PowerShell.Authentication.AuthContext. The connection information object for the Microsoft Graph session.

.EXAMPLE
$graphConnection = Connect-GraphAdvanced -Scopes User.Read, Mail.Read -Modules Microsoft.Graph
This example establishes a connection to Microsoft Graph with the scopes "User.Read" and "Mail.Read" using the "Microsoft.Graph" module.

.EXAMPLE
$graphConnection = Connect-GraphAdvanced -Scopes Group.Read.All, User.Read.All -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups
This example establishes a connection to Microsoft Graph with the scopes "Group.Read.All" and "User.Read.All" using the "Microsoft.Graph.Users" and "Microsoft.Graph.Groups" modules.

.EXAMPLE
$graphConnection = Connect-GraphAdvanced -Scopes Group.Read.All, User.Read.All -Modules Microsoft.Graph.Users, Microsoft.Graph.Groups -minModuleVersion 2.25.0
This example establishes a connection to Microsoft Graph with the scopes "Group.Read.All" and "User.Read.All" using the "Microsoft.Graph.Users" and "Microsoft.Graph.Groups" modules, and specifies a minimum module version of 2.25.0.
#>

. $PSScriptRoot\..\ModuleHandle.ps1

function Connect-GraphAdvanced {
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Scopes,
        [Parameter(Mandatory = $true)]
        [string[]]$Modules,
        [Parameter(Mandatory = $false)]
        [string[]]$TenantId = $null,
        [Parameter(Mandatory = $false)]
        [switch]$DoNotShowConnectionDetails,
        [ValidateScript({
                if ($_ -lt [System.Version]'2.0.0.0') {
                    throw "Minimum supported version: 2.0.0.0"
                }
                $true
            })]
        [Parameter(Mandatory = $false)]
        [System.Version]$MinModuleVersion = '2.0.0.0'
    )

    #Validate Graph is installed and loaded
    $requestModule = Request-Module -Module $Modules -MinModuleVersion $MinModuleVersion
    if (-not $requestModule) {
        Write-Host "We cannot continue without $Modules Powershell module" -ForegroundColor Red
        return $null
    }

    #Validate Graph is connected or try to connect
    $connection = $null
    try {
        $connection = Get-MgContext -ErrorAction Stop
    } catch {
        Write-Host "We cannot check context. Error:`n$_" -ForegroundColor Red
        return $null
    }

    if ($null -eq $connection) {
        Write-Host "Not connected to Graph" -ForegroundColor Yellow
        $connection = Add-GraphConnection -Scopes $Scopes
    } else {
        Write-Verbose "You have a Graph sessions"
        Write-Verbose "Checking scopes"
        if (-not (Test-GraphScopeContext -Scopes $connection.Scopes -ExpectedScopes $Scopes)) {
            Write-Host "Not connected to Graph with expected scopes" -ForegroundColor Yellow
            $connection = Add-GraphConnection -Scopes $Scopes
        } else {
            Write-Verbose "All scopes are present"
        }
    }

    if ($connection) {
        Write-Verbose "Checking TenantId"
        if ($TenantId) {
            if ($connection.TenantId -ne $TenantId) {
                Write-Host "Connected to $($connection.TenantId). Not expected tenant: $TenantId" -ForegroundColor Red
                return $null
            } else {
                Write-Verbose "TenantId is correct"
            }
        }

        $connection.PSObject.Properties | ForEach-Object { Write-Verbose "$($_.Name): $($_.Value)" }
        if (-not $DoNotShowConnectionDetails) {
            Show-GraphContext -Context $connection
        }
    }
    return $connection
}

function Add-GraphConnection {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "High")]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$Scopes
    )

    if ($PSCmdlet.ShouldProcess("Do you want to connect?", "We need a Graph connection with scopes $Scopes")) {
        Write-Verbose "Connecting to Microsoft Graph API using scopes $Scopes"
        try {
            Connect-MgGraph -Scopes $Scopes -NoWelcome -ErrorAction Stop
        } catch {
            Write-Host "We cannot connect to Graph. Error:`n$_" -ForegroundColor Red
            return $null
        }
        $connection = $null
        try {
            $connection = Get-MgContext -ErrorAction Stop
        } catch {
            Write-Host "We cannot check context. Error:`n$_" -ForegroundColor Red
            return $null
        }
        Write-Verbose "Checking scopes"
        if (-not $connection) {
            Write-Host "We cannot continue without Graph Powershell session" -ForegroundColor Red
            return $null
        }
        if (-not (Test-GraphScopeContext -Scopes $connection.Scopes -ExpectedScopes $Scopes)) {
            Write-Host "We cannot continue without Graph Powershell session without Expected Scopes" -ForegroundColor Red
            return $null
        }
        return $connection
    }
}

function Test-GraphScopeContext {
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string[]]$ExpectedScopes,
        [Parameter(Mandatory = $true)]
        [string[]]$Scopes
    )

    $foundError = $false
    foreach ($expectedScope in $ExpectedScopes) {
        if ($Scopes -notcontains $expectedScope) {
            Write-Host "The following scope is missing: $expectedScope" -ForegroundColor Red
            $foundError = $true
        }
    }

    Write-Verbose "All expected scopes are $(if($foundError){ "NOT "})present."
    return (-not $foundError)
}

function Show-GraphContext {
    param (
        [Parameter(Mandatory = $true)]
        [Microsoft.Graph.PowerShell.Authentication.AuthContext]$Context
    )
    Write-Host "`nConnected to Graph"
    Write-Host "Session details"
    Write-Host "Tenant Id: $($Context.TenantId)"
    if ($graphConnection.AuthType) {
        Write-Host "AuthType: $($graphConnection.AuthType)"
    }
    if ($Context.Account) {
        Write-Host "Account: $($Context.Account)"
    }
}
