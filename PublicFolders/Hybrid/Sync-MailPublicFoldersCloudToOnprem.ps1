﻿# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# .SYNOPSIS
# Sync-MailPublicFoldersCloudToOnprem.ps1
#    This script imports the new mail public folders as sync mail public folders from Exchange Online to on-premise.
#	 And also synchronizes the properties of existing mail-enabled public folders from Exchange Online to on-premises (thereby overriding the mail public folder properties in on-premise).
#
# Example input to the script:
#
# Sync-MailPublicFoldersCloudToOnprem.ps1 -ConnectionUri <cloud url> -CsvSummaryFile <path for the summary file>
#
# The above example imports new mail public folders objects from Exchange Online as sync mail public folders to on-premise.
param (
    [Parameter(Mandatory=$false)]
    [PSCredential] $Credential,

    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [string] $ConnectionUri = "https://outlook.office365.com/powerShell-liveID",

    [Parameter(Mandatory=$true)]
    [ValidateNotNullOrEmpty()]
    [string] $CsvSummaryFile
)

# cspell:words EXOV2 MEPF

# Writes a dated information message to console
function WriteInfoMessage() {
    param ($message)
    Write-Host "[$($(Get-Date).ToString())]" $message
}

# Writes an error importing a mail public folder to the CSV summary
function WriteErrorSummary() {
    param ($folder, $operation, $errorMessage, $commandText)

    WriteOperationSummary -folder $folder.Guid -operation $operation -result $errorMessage -commandText $commandText
    $script:errorsEncountered++
}

# Writes the operation executed and its result to the output CSV
function WriteOperationSummary() {
    param ($folder, $operation, $result, $commandText)

    $columns = @(
        (Get-Date).ToString(),
        $folder.Guid,
        $operation,
        (EscapeCsvColumn $result),
        (EscapeCsvColumn $commandText)
    )

    Add-Content $CsvSummaryFile -Value ("{0},{1},{2},{3},{4}" -f $columns)
}

#Escapes a column value based on RFC 4180 (http://tools.ietf.org/html/rfc4180)
function EscapeCsvColumn() {
    param ([string]$text)

    if ($text -eq $null) {
        return $text
    }

    $hasSpecial = $false
    for ($i=0; $i -lt $text.Length; $i++) {
        $c = $text[$i]
        if ($c -eq $script:csvEscapeChar -or
            $c -eq $script:csvFieldDelimiter -or
            $script:csvSpecialChars -contains $c) {
            $hasSpecial = $true
            break
        }
    }

    if (-not $hasSpecial) {
        return $text
    }

    $ch = $script:csvEscapeChar.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    return $ch + $text.Replace($ch, $ch + $ch) + $ch
}

## Create a tenant PSSession against Exchange Online using modern auth.
function InitializeExchangeOnlineRemoteSession() {
    WriteInfoMessage $LocalizedStrings.CreatingRemoteSession

    try {
        Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue
        if (Get-Module ExchangeOnlineManagement) {
            $sessionOption = (New-PSSessionOption -SkipCACheck)
            Connect-ExchangeOnline -Credential $Credential -ConnectionURI $ConnectionUri -PSSessionOption $sessionOption -Prefix "Remote" -ErrorAction SilentlyContinue
            $script:isConnectedToExchangeOnline = $true
        } else {
            Write-Warning $LocalizedStrings.EXOV2ModuleNotInstalled
            exit
        }
    } catch {
        Write-Error "Error message: $($_.Exception.Message)"
        WriteInfoMessage ($LocalizedStrings.ConnectExchangeOnlineFailure)
        exit
    }
    WriteInfoMessage ($LocalizedStrings.RemoteSessionCreatedSuccessfully)
}

## Formats the command and its parameters to be printed on console or to file
function FormatCommand() {
    param ([string]$command, [System.Collections.IDictionary]$parameters)

    $commandText = New-Object System.Text.StringBuilder
    [void]$commandText.Append($command)
    foreach ($name in $parameters.Keys) {
        [void]$commandText.AppendFormat(" -{0}:", $name)

        $value = $parameters[$name]
        if ($value -isnot [Array]) {
            [void]$commandText.AppendFormat("`"{0}`"", $value)
        } elseif ($value.Length -eq 0) {
            [void]$commandText.Append("@()")
        } else {
            [void]$commandText.Append("@(")
            foreach ($subValue in $value) {
                [void]$commandText.AppendFormat("`"{0}`",", $subValue)
            }

            [void]$commandText.Remove($commandText.Length - 1, 1)
            [void]$commandText.Append(")")
        }
    }

    return $commandText.ToString()
}

## Get external email address domains which are not configured on the on-premise connector
function GetMissingDomainsFromConnector() {
    param($sendConnectorDomains)

    $missingDomains = @()

    foreach ($domain in $script:ExternalEmailAddressDomains.Keys) {
        if (-not $sendConnectorDomains.ContainsKey($domain)) {
            $missingDomains += $domain
        }
    }

    return $missingDomains
}

## Check if the centralized transport feature is enabled
function IsCentralizedTransportEnabled() {
    $hybridConf = Get-HybridConfiguration

    if ($null -ne $hybridConf -and $null -ne $hybridConf.Features -and $hybridConf.Features.Contains("CentralizedTransport")) {
        return $true
    }

    return $false
}

## Get send connector (on-premise) domains configured for O365
function GetSendConnectorDomains() {
    $connector = Get-SendConnector -Identity $script:SendConnectorToO365 -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
    $domains = @{}

    if ($null -ne $connector -and $null -ne $connector.AddressSpaces) {
        foreach ($addressSpace in $connector.AddressSpaces) {
            $domain = $addressSpace.Address.ToString()

            if (-not $domains.ContainsKey($domain)) {
                $domains.Add($domain, $false)
            }
        }
    }

    return $domains
}

## Store distinct domains used in the external email addresses
function StoreExternalEmailAddressDomain(
    [string] $emailAddress) {
    if (-not [string]::IsNullOrEmpty($emailAddress)) {
        $index = $emailAddress.IndexOf('@')
        $domain = $emailAddress.Substring($index + 1)

        if (-not $script:ExternalEmailAddressDomains.ContainsKey($domain)) {
            $script:ExternalEmailAddressDomains.Add($domain, $false)
        }
    }
}

## Concatenate a list of domains
function ConcatDomains() {
    param($domainList)

    $domains = $script:separator
    $domains += $domainList -join $script:separator
    return $domains
}

## Get external email address from an EXO mail public folder
function GetExternalEmailAddress() {
    param ($remotePublicFolder)

    $primarySmtpAddress = $remotePublicFolder.PrimarySmtpAddress.ToString()

    if ($primarySmtpAddress.EndsWith($script:OnMicrosoftDomain, [StringComparison]::InvariantCultureIgnoreCase)) {
        return $primarySmtpAddress
    }

    $externalEmailAddress = $primarySmtpAddress
    $primarySmtpAddressParts = $primarySmtpAddress.Split($script:proxyAddressSeparators); # alias@domain

    if ($null -ne $remotePublicFolder.EmailAddresses -and $remotePublicFolder.EmailAddresses.Count -gt 0) {
        foreach ($address in $remotePublicFolder.EmailAddresses) {
            if ($address.StartsWith($script:SmtpPrefix, [StringComparison]::InvariantCultureIgnoreCase)) {
                $addressParts = $address.Split($script:proxyAddressSeparators); # smtp:alias@domain

                if ($addressParts.Count -eq 3 -and
                    $addressParts[1].Equals($primarySmtpAddressParts[0], [StringComparison]::InvariantCultureIgnoreCase) -and
                    $addressParts[2].EndsWith($script:OnMicrosoftDomain, [StringComparison]::InvariantCultureIgnoreCase)) {
                    return (RemoveSmtpPrefix $address)
                }
            }
        }
    }

    return $externalEmailAddress
}

## Remove smtp prefix from the email address
function RemoveSmtpPrefix() {
    param ($emailAddress)

    if ([String]::IsNullOrEmpty($emailAddress)) {
        return [String]::Empty
    }

    if ($emailAddress.StartsWith($script:SmtpPrefix, [StringComparison]::InvariantCultureIgnoreCase)) {
        return $emailAddress.Substring($script:SmtpPrefixLength)
    }

    return $emailAddress
}

## Retrieve mail enabled public folders from EXO
function GetRemoteMailPublicFolders() {
    $mailPublicFolders = Get-RemoteMailPublicFolder -ResultSize:Unlimited -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue

    # Return the results
    if ($null -eq $mailPublicFolders -or ([array]($mailPublicFolders)).Count -eq 0) {
        return $null
    }

    return $mailPublicFolders
}

#Get list of MEPFs in OnPrem to be deleted
function GetFoldersToMailDisable(
				[object[]] $localMailPublicFolders,
				[hashtable] $validExternalEmailAddresses) {
    $foldersToMailDisable = @()
    foreach ($syncPublicFolder in $localMailPublicFolders) {
        $localExternalEmailAddress = [String]::Empty

        if ($null -ne $syncPublicFolder.ExternalEmailAddress) {
            $localExternalEmailAddress = RemoveSmtpPrefix $syncPublicFolder.ExternalEmailAddress.ToString()
        }

        if (-not $validExternalEmailAddresses.ContainsKey($localExternalEmailAddress)) {
            $foldersToMailDisable += $syncPublicFolder.Identity
        }
    }
    return $foldersToMailDisable
}

## Sync mail public folders from cloud to on-premise.
function SyncMailPublicFolders {
    [CmdletBinding(SupportsShouldProcess)]
    param(
        [object[]] $mailPublicFolders
    )

    $validExternalEmailAddresses = @{}

    if ($null -ne $mailPublicFolders) {
        foreach ($mailPublicFolder in $mailPublicFolders) {
            # extracting properties
            $alias = $mailPublicFolder.Alias.Trim()
            $identity = $mailPublicFolder.PrimarySmtpAddress.ToString()
            $externalEmailAddress = GetExternalEmailAddress $mailPublicFolder
            $entryId = $mailPublicFolder.EntryId.ToString()
            $name = $mailPublicFolder.Name.Trim()
            $displayName = $mailPublicFolder.DisplayName.Trim()
            $hiddenFromAddressListsEnabled = $mailPublicFolder.HiddenFromAddressListsEnabled

            $windowsEmailAddress = $mailPublicFolder.WindowsEmailAddress.ToString()
            if ($windowsEmailAddress -eq "") {
                $windowsEmailAddress = $externalEmailAddress
            }

            # extracting all the EmailAddress
            $emailAddress = @()
            foreach ($address in $mailPublicFolder.EmailAddresses) {
                $emailAddress += $address.ToString()
            }

            # preserve the ability to reply via Outlook's nickname cache post-migration
            $x500Proxy = ("X500:" + $mailPublicFolder.LegacyExchangeDN)
            if ($x500Proxy -notin $emailAddress) {
                $emailAddress += $x500Proxy
            }

            WriteInfoMessage ($LocalizedStrings.SyncingMailPublicFolder -f $alias)

            $syncPublicFolder = Get-MailPublicFolder -Identity $identity -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue

            if ($null -eq $syncPublicFolder) {
                WriteInfoMessage ($LocalizedStrings.CreatingSyncMailPublicFolder -f $alias)
                try {
                    $newParams = @{}
                    $newParams.Add("Name", $name)
                    $newParams.Add("ExternalEmailAddress", $externalEmailAddress)
                    $newParams.Add("Alias", $alias)
                    $newParams.Add("EntryId", $entryId)
                    $newParams.Add("WindowsEmailAddress", $windowsEmailAddress)
                    $newParams.Add("WarningAction", "SilentlyContinue")
                    $newParams.Add("ErrorAction", "Stop")

                    [string]$createSyncPublicFolder = (FormatCommand $script:NewSyncMailPublicFolderCommand $newParams)

                    # Creating new sync mail public folder
                    $null = &$script:NewSyncMailPublicFolderCommand @newParams

                    WriteOperationSummary -folder $mailPublicFolder -operation $LocalizedStrings.CreateOperationName -result $LocalizedStrings.CsvSuccessResult -commandText $createSyncPublicFolder

                    $setParams = @{}
                    $setParams.Add("Identity", $name)
                    $setParams.Add("EmailAddresses", $emailAddress)
                    $setParams.Add("DisplayName", $displayName)
                    $setParams.Add("HiddenFromAddressListsEnabled", $hiddenFromAddressListsEnabled)
                    $setParams.Add("WarningAction", "SilentlyContinue")
                    $setParams.Add("ErrorAction", "Stop")
                    $setParams.Add("EmailAddressPolicyEnabled", $false)

                    [string]$setOtherProperties = (FormatCommand $script:SetMailPublicFolderCommand $setParams)

                    # Setting other properties to the newly created sync mail public folder
                    &$script:SetMailPublicFolderCommand @setParams

                    WriteOperationSummary -folder $mailPublicFolder -operation $LocalizedStrings.SetOperationName -result $LocalizedStrings.CsvSuccessResult -commandText $setOtherProperties

                    $validExternalEmailAddresses.Add($externalEmailAddress, $false)
                    $script:CreatedPublicFoldersCount++
                }

                catch {
                    WriteErrorSummary -folder $mailPublicFolder -operation $LocalizedStrings.CreateOperationName -errorMessage $_.Exception.Message -commandText ""
                    Write-Error $_
                }
            }

            else {
                WriteInfoMessage ($LocalizedStrings.UpdatingSyncMailPublicFolder -f $syncPublicFolder)
                try {
                    $updateParams = @{}
                    $updateParams.Add("Identity", $syncPublicFolder)
                    $updateParams.Add("EmailAddresses", $emailAddress)
                    $updateParams.Add("HiddenFromAddressListsEnabled", $hiddenFromAddressListsEnabled)
                    $updateParams.Add("DisplayName", $displayName)
                    $updateParams.Add("Name", $name)
                    $updateParams.Add("ExternalEmailAddress", $externalEmailAddress)
                    $updateParams.Add("Alias", $alias)
                    $updateParams.Add("WindowsEmailAddress", $windowsEmailAddress)
                    $updateParams.Add("WarningAction", "SilentlyContinue")
                    $updateParams.Add("ErrorAction", "Stop")

                    [string]$updateProperties = (FormatCommand $script:SetMailPublicFolderCommand $updateParams)

                    # Setting properties to the existing sync mail public folder
                    &$script:SetMailPublicFolderCommand @updateParams

                    WriteOperationSummary -folder $mailPublicFolder -operation $LocalizedStrings.UpdateOperationName -result $LocalizedStrings.CsvSuccessResult -commandText $updateProperties

                    $validExternalEmailAddresses.Add($externalEmailAddress, $false)
                    $script:UpdatedPublicFoldersCount++
                }

                catch {
                    WriteErrorSummary -folder $mailPublicFolder -operation $LocalizedStrings.UpdateOperationName -errorMessage $_.Exception.Message -commandText $updateProperties
                    Write-Error $_
                }
            }

            StoreExternalEmailAddressDomain $externalEmailAddress
            WriteInfoMessage ($LocalizedStrings.DoneSyncingMailPublicFolder -f $alias)
            Write-Host ""
        }
    }

    else {
        WriteInfoMessage ($LocalizedStrings.NoMailPublicFoldersToSync)
        Write-Host ""
    }

    WriteInfoMessage ($LocalizedStrings.DeleteSyncMailPublicFolderTitle)

    $localMailPublicFolders = Get-MailPublicFolder -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue

    $foldersToMailDisable = GetFoldersToMailDisable $localMailPublicFolders $validExternalEmailAddresses
    if ($null -eq $foldersToMailDisable) {
        return
    }

    $foldersToMailDisableFile = Join-Path $PWD "FoldersToMailDisable.txt"
    Set-Content -Path $foldersToMailDisableFile -Value $foldersToMailDisable

    if (-not $PSCmdlet.ShouldProcess($LocalizedStrings.FoldersToMailDisableConfirmation -f $foldersToMailDisableFile)) {
        return
    }

    foreach ($syncPublicFolder in $localMailPublicFolders) {
        $localExternalEmailAddress = [String]::Empty

        if ($null -ne $syncPublicFolder.ExternalEmailAddress) {
            $localExternalEmailAddress = RemoveSmtpPrefix $syncPublicFolder.ExternalEmailAddress.ToString()
        }

        if (-not $validExternalEmailAddresses.ContainsKey($localExternalEmailAddress)) {
            WriteInfoMessage ($LocalizedStrings.DeletingSyncMailPublicFolder -f $syncPublicFolder)
            try {
                $deleteParams = @{}
                $deleteParams.Add("Identity", $syncPublicFolder)
                $deleteParams.Add("Confirm", $false)
                [string]$disableMailPublicFolder = (FormatCommand $script:DeletePublicFolderCommand $deleteParams)

                # Deleting sync mail public folder
                &$script:DeletePublicFolderCommand @deleteParams

                WriteOperationSummary -folder $syncPublicFolder -operation $LocalizedStrings.DeleteOperationName -result $LocalizedStrings.CsvSuccessResult -commandText $disableMailPublicFolder
                $script:RemovedPublicFoldersCount++
            } catch {
                WriteErrorSummary -folder $syncPublicFolder -operation $LocalizedStrings.DeleteOperationName -errorMessage $_.Exception.Message -commandText $disableMailPublicFolder
                Write-Error $_
            }
        }
    }
}

################ DECLARING GLOBAL VARIABLES ################
$script:session = $null

$script:csvSpecialChars = @("`r", "`n")
$script:csvEscapeChar = '"'
$script:csvFieldDelimiter = ','
$script:separator = "`n`t"
$script:NewSyncMailPublicFolderCommand = "New-SyncMailPublicFolder"
$script:SetMailPublicFolderCommand = "Set-MailPublicFolder"
$script:DeletePublicFolderCommand = "Disable-MailPublicFolder"
$script:CreatedPublicFoldersCount = 0
$script:UpdatedPublicFoldersCount = 0
$script:RemovedPublicFoldersCount = 0
$script:ExternalEmailAddressDomains = @{}
$script:SendConnectorToO365 = "Outbound to Office 365"
$script:OnMicrosoftDomain = ".onmicrosoft.com"
$script:SmtpPrefix = "smtp:"
$script:SmtpPrefixLength = $script:SmtpPrefix.Length
[char[]]$script:proxyAddressSeparators = ':', '@'

#load hashtable of localized string
$LocalizedStrings = ConvertFrom-StringData @'
SyncingMailPublicFolder = Syncing mail public folder '{0}'.
CreatingSyncMailPublicFolder = Creating sync mail public folder object '{0}'.
UpdatingSyncMailPublicFolder = Sync mail public folder object '{0}' already exists, hence updating properties.
DoneSyncingMailPublicFolder = Done syncing mail public folder '{0}'.
NoMailPublicFoldersToSync = There aren't any mail public folders in cloud to sync.
DeleteSyncMailPublicFolderTitle = Deleting sync mail public folder, if any, that don't have corresponding mail public folders in the cloud.
DeletingSyncMailPublicFolder = Deleting sync mail public folder for object '{0}', as this is no more in the cloud.
CreateOperationName = Create
SetOperationName = Set
UpdateOperationName = Update
DeleteOperationName = Delete
TimestampCsvHeader = Timestamp
IdentityCsvHeader = Identity
OperationCsvHeader = Operation
ResultCsvHeader = Result
CommandCsvHeader = Command text
CsvSuccessResult = Success
LocalServerVersionNotSupported = You cannot execute this script from your local Exchange server: "{0}". This script can only be executed from Exchange 2013 Management Shell and above.
CreatingRemoteSession = Creating an Exchange Online remote session...
RemoteSessionCreatedSuccessfully = Exchange Online remote session created successfully.
StartedImportingMailPublicFolders = Started import of mail public folders.
CompletedImportingMailPublicFolders = Completed import of mail public folders.
CompletedStatsCount = Total sync mail mail public folders created: {0}, updated: {1} and deleted: {2}.
VerifyConnectorAcceptedDomainsMessage = Please make sure that the following domain(s) are added to the on-premise hybrid connector to avoid a possibility of mail looping: {0}
FoldersToMailDisableConfirmation = You are about to mail-disable the MEPF's in {0} (Y/N).
ConnectExchangeOnlineFailure = Connection to Exchange-Online has failed, terminating the script.
EXOV2ModuleNotInstalled = This script uses modern authentication to connect to Exchange Online and requires EXO V2 module to be installed. Please follow the instructions at https://docs.microsoft.com/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-the-exo-v2-module to install EXO V2 module.
'@

#minimum supported exchange version to run this script
$minSupportedVersion = 15
################ END OF DECLARATION #########################

if (Test-Path $CsvSummaryFile) {
    Remove-Item $CsvSummaryFile -Confirm:$Confirm -Force
}

# Write the output CSV headers
$null = New-Item -Path $CsvSummaryFile -ItemType File -Force -ErrorAction:Stop -Value ("#{0},{1},{2},{3},{4}`r`n" -f $LocalizedStrings.TimestampCsvHeader,
    $LocalizedStrings.IdentityCsvHeader,
    $LocalizedStrings.OperationCsvHeader,
    $LocalizedStrings.ResultCsvHeader,
    $LocalizedStrings.CommandCsvHeader)

$localServerVersion = (Get-ExchangeServer $env:COMPUTERNAME -ErrorAction:Stop).AdminDisplayVersion
# This script can run from Exchange 2013 Management shell and above
if ($localServerVersion.Major -lt $minSupportedVersion) {
    Write-Error ($LocalizedStrings.LocalServerVersionNotSupported -f $localServerVersion) -ErrorAction:Continue
    exit
}

InitializeExchangeOnlineRemoteSession

# Get mail enabled public folders in cloud
WriteInfoMessage ($LocalizedStrings.StartedImportingMailPublicFolders)
Write-Host ""

$mailPublicFoldersEXO = GetRemoteMailPublicFolders

# Create sync mail public folders in on-premise
SyncMailPublicFolders $mailPublicFoldersEXO

Write-Host ""
WriteInfoMessage ($LocalizedStrings.CompletedImportingMailPublicFolders)
WriteInfoMessage ($LocalizedStrings.CompletedStatsCount -f  $script:CreatedPublicFoldersCount, $script:UpdatedPublicFoldersCount, $script:RemovedPublicFoldersCount)
Write-Host ""

if ((IsCentralizedTransportEnabled) -and $script:ExternalEmailAddressDomains.Count -gt 0) {
    $sendConnectorDomains = GetSendConnectorDomains

    if ($null -eq $sendConnectorDomains -or $sendConnectorDomains.Count -eq 0) {
        $domains = ConcatDomains $script:ExternalEmailAddressDomains.Keys
        Write-Warning ($LocalizedStrings.VerifyConnectorAcceptedDomainsMessage -f $domains)
    } else {
        $missingDomains = GetMissingDomainsFromConnector $sendConnectorDomains

        if ($null -ne $missingDomains -and $missingDomains.Count -gt 0) {
            $domains = ConcatDomains $missingDomains
            Write-Warning ($LocalizedStrings.VerifyConnectorAcceptedDomainsMessage -f $domains)
        }
    }
}

# Terminate the PSSession
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue
