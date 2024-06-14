# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

function Request-Module {
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        [Parameter(Mandatory = $false)]
        [System.Version]$MinModuleVersion = $null,
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )

    if ($MinModuleVersion) {
        Write-Verbose "Checking $ModuleName PowerShell Module is loaded with minimum version $MinModuleVersion"
        $loaded = Test-Module -Loaded -ModuleName $ModuleName -MinModuleVersion $MinModuleVersion
        if (-not $loaded) {
            Write-Verbose "$ModuleName PowerShell module not loaded with minimum version $MinModuleVersion"
            $installed = Test-Module -Installed -ModuleName $ModuleName -MinModuleVersion $MinModuleVersion
            if (-not $installed) {
                Write-Verbose "$ModuleName PowerShell Module not installed with minimum version $MinModuleVersion"
                $installed = Add-Module -ModuleName $ModuleName -MinModuleVersion $MinModuleVersion
                if (-not $installed) {
                    Write-Host "$ModuleName Powershell Module failed to install with minimum version $MinModuleVersion"
                    return $false
                }
            }
            Write-Verbose "Importing $ModuleName PowerShell Module with minimum version $MinModuleVersion"
            Import-Module $ModuleName -MinimumVersion  $MinModuleVersion -ErrorAction SilentlyContinue -Force
            $loaded = Test-Module -Loaded -ModuleName $ModuleName -MinModuleVersion $MinModuleVersion
            if (-not $loaded) {
                Write-Host "$ModuleName Powershell Module import failed with minimum version $MinModuleVersion" -ForegroundColor Red
                return $false
            }
            Write-Verbose "$ModuleName Powershell Module loaded with minimum version $MinModuleVersion"
            return $true
        }
        Write-Verbose "$ModuleName Powershell Module already loaded with minimum version $MinModuleVersion"
        return $true
    } else {
        Write-Verbose "Checking $ModuleName Powershell Module is loaded."
        $loaded = Test-Module -Loaded -ModuleName $ModuleName
        if (-not $loaded) {
            Write-Verbose "$ModuleName PowerShell module not loaded"
            $installed = Test-Module -Installed -ModuleName $ModuleName
            if (-not $installed) {
                Write-Verbose "$ModuleName PowerShell Module not installed"
                $installed = Add-Module -ModuleName $ModuleName
                if (-not $installed) {
                    Write-Host "$ModuleName Powershell Module failed to install."
                    return $false
                }
            }
            Write-Verbose "Importing $ModuleName PowerShell Module"
            Import-Module $ModuleName -ErrorAction SilentlyContinue -Force
            $loaded = Test-Module -Loaded -ModuleName $ModuleName
            if (-not $loaded) {
                Write-Host "$ModuleName Powershell Module import failed." -ForegroundColor Red
                return $false
            }
            Write-Verbose "$ModuleName Powershell Module loaded."
            return $true
        }
        Write-Verbose "$ModuleName Powershell Module already loaded."
        return $true
    }
}

function Test-Module {
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true, ParameterSetName = 'Installed')]
        [switch]$Installed,
        [Parameter(Mandatory = $true, ParameterSetName = 'Loaded')]
        [switch]$Loaded,
        [Parameter(Mandatory = $true, ParameterSetName = 'Installed')]
        [Parameter(Mandatory = $true, ParameterSetName = 'Loaded')]
        [string]$ModuleName,
        [Parameter(Mandatory = $false, ParameterSetName = 'Installed')]
        [Parameter(Mandatory = $false, ParameterSetName = 'Loaded')]
        [System.Version]$MinModuleVersion = $null
    )

    $modules = $null
    if ($Installed) {
        Write-Verbose "Checking installation of $ModuleName"
        $modules = Get-Module $ModuleName -ListAvailable | Sort-Object Version -Descending
    } elseif ($Loaded) {
        Write-Verbose "Checking load of $ModuleName"
        $modules = Get-Module $ModuleName | Sort-Object Version -Descending
    }

    if ($modules) {
        Write-Verbose "Detected $ModuleName"
        if ($MinModuleVersion) {
            $foundMinVersion = $false
            foreach ($module in $modules) {
                if (Test-ModuleVersion -MinModuleVersion $MinModuleVersion -ModuleVersion $module.Version) {
                    Write-Verbose "Powershell Module $ModuleName with minimum version $MinModuleVersion': $($module.Version)"
                    $foundMinVersion = $true
                }
            }
            if ($foundMinVersion) {
                Write-Verbose "Powershell Module $ModuleName with minimum version $MinModuleVersion Detected"
                return $true
            } else {
                Write-Host "$ModuleName Powershell Module do not reach minimum version: $MinModuleVersion" -ForegroundColor Red
                return $false
            }
        } else {
            Write-Verbose "$ModuleName Powershell Module found"
            return $true
        }
    }
}

function Add-Module {
    [CmdletBinding(SupportsShouldProcess, ConfirmImpact = 'High')]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        [Parameter(Mandatory = $false)]
        [System.Version]$MinModuleVersion = $null,
        [Parameter(Mandatory = $false)]
        [switch]$Force
    )


    if ($MinModuleVersion) {
        Write-Verbose "Testing $ModuleName with minimum version $minModuleVersion"
        $testModule = (Test-Module -Installed -ModuleName $ModuleName -MinModuleVersion $MinModuleVersion)
    } else {
        Write-Verbose "Testing $ModuleName"
        $testModule = (Test-Module -Installed -ModuleName $ModuleName)
    }

    if (-not $testModule) {
        if ($MinModuleVersion) {
            $message = "Do you want to install min version $MinModuleVersion?"
        } else {
            $message = "Do you want to install?"
        }

        if ($Force -or $PSCmdlet.ShouldContinue($message, "Module $ModuleName")) {
            if ($MinModuleVersion) {
                Write-Verbose "Installing $ModuleName with minimum version $minModuleVersion"
                Install-Module -Name $ModuleName -Force -ErrorAction SilentlyContinue -Scope CurrentUser -MinimumVersion $MinModuleVersion
            } else {
                Write-Verbose "Installing $ModuleName"
                Install-Module -Name $ModuleName -Force -ErrorAction SilentlyContinue -Scope CurrentUser
            }
            if ($MinModuleVersion) {
                Write-Verbose "Verifying installation $ModuleName with minimum version $minModuleVersion"
                $testModule = (Test-Module -Installed -ModuleName $ModuleName -MinModuleVersion $MinModuleVersion)
            } else {
                Write-Verbose "Verifying installation $ModuleName"
                $testModule = (Test-Module -Installed -ModuleName $ModuleName)
            }
            if ($testModule) {
                Write-Verbose "$ModuleName Powershell Module installed"
                return $true
            } else {
                Write-Host "$ModuleName Powershell Module installation failed" -ForegroundColor Red
                return $false
            }
        }
    } else {
        Write-Verbose "$ModuleName Powershell Module already installed"
        return $true
    }
}

function Test-ModuleVersion {
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [System.Version]$MinModuleVersion,
        [Parameter(Mandatory = $true)]
        [System.Version]$ModuleVersion
    )
    return $ModuleVersion -ge $MinModuleVersion
}

