# .SYNOPSIS
# Import-PublicFolderMailboxes.ps1
#    Import the public folder mailboxes as mail enabled users from cloud to on-premise 
#
# Example input to the script:
#
# Import-PublicFolderMailboxes.ps1 -ConnectionUri <cloud url>
#
# The above example imports public folder mailbox objects from cloud as mail enabled users to on-premise.
#
# .DESCRIPTION
#
# Copyright (c) 2012 Microsoft Corporation. All rights reserved.
#
# THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK
# OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
param (
    [Parameter(Mandatory=$false)]
    [PSCredential] $Credential,

    [Parameter(Mandatory = $false)]
    [ValidateNotNull()]
    [string] $ConnectionUri = "https://outlook.office365.com/powerShell-liveID"
)

## Create a tenant PSSession.
function CreateTenantSession()
{
    Import-Module ExchangeOnlineManagement -ErrorAction SilentlyContinue;
    if (Get-Module ExchangeOnlineManagement)
    {
        $sessionOption = (New-PSSessionOption -SkipCACheck);
        Connect-ExchangeOnline -Credential $Credential -ConnectionURI $ConnectionUri -PSSessionOption $sessionOption -Prefix "Remote" -ErrorAction SilentlyContinue;
    }
    else
    {
        Write-Warning $LocalizedStrings.EXOV2ModuleNotInstalled;
        Exit;
    }
}

## Writes a dated information message to console
function WriteInfoMessage()
{
    param ($message)
    Write-Host "[$($(Get-Date).ToString())]" $message;
}

## Retrieve public folder mailboxes
function GetPublicFolderMailBoxes()
{
    $publicFolderMailboxes = Get-RemoteMailbox -PublicFolder -ResultSize:Unlimited -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue

    # Return the results     
    if ($publicFolderMailboxes -eq $null -or ([array]($publicFolderMailboxes)).Count -lt 1)
    {
        return $null;
    }

    return $publicFolderMailboxes;
}

## Sync public folder mailboxes from cloud to on-prem.
function SyncPublicFolderMailboxes(
    [object[]] $publicFolderMailboxes)
{
    WriteInfoMessage ($LocalizedStrings.DeletingMailUsersInfo);
    $remoteMailboxes = Get-OrganizationConfig | select RemotePublicFolderMailboxes

    foreach($adObjectId in $remoteMailboxes.RemotePublicFolderMailboxes)
    {
        $mailUser = Get-MailUser $adObjectId -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue
        
        if ($mailUser -ne $null)
        {
                WriteInfoMessage ($LocalizedStrings.RemovingMailUsers -f $mailUser);
                Set-OrganizationConfig -RemotePublicFolderMailboxes @{Remove=$mailUser}

                WriteInfoMessage ($LocalizedStrings.DeleteMailUser -f $mailUser);
                Remove-MailUser $mailUser -Confirm:$false 
        }
    }
    Set-OrganizationConfig -RemotePublicFolderMailboxes:$Null

    $validExternalEmailAddresses = @();
    $mailUserList = @();

    if ($publicFolderMailboxes -ne $null)
    {
        $hasPublicFolderServingHierarchy = $false;
        foreach ($publicFolderMailbox in $publicFolderMailboxes)
        {
            if ($publicFolderMailbox.IsExcludedFromServingHierarchy -eq $false)
            {  
                $hasPublicFolderServingHierarchy = $true;
                $displayName = $publicFolderMailbox.Name.ToString().Trim();
                $name = "RemotePfMbx-" + $displayName + "-" + [guid]::NewGuid();
		        $name = $(if ($name.length -gt 64) { $name.substring(0, 64) } else { $name });
                $externalEmailAddress = $publicFolderMailbox.PrimarySmtpAddress.ToString();

                WriteInfoMessage ($LocalizedStrings.SyncingPublicFolderMailbox -f $displayName);

                $mailUser = Get-MailUser $externalEmailAddress -ErrorAction:SilentlyContinue -WarningAction:SilentlyContinue;

                if ($mailUser -eq $null)
                {
                    WriteInfoMessage ($LocalizedStrings.CreatingMailUser -f $displayName);
		            try
                    {
                        $mailUser = New-MailUser -Name $name -ExternalEmailAddress $externalEmailAddress -DisplayName $displayName;
                        $mailUserList += $mailUser;
                    }
                    catch
                    {
                        Write-Host $error[0];
                    }		    
                }
                else
                {
                    WriteInfoMessage ($LocalizedStrings.MailUserExists -f $mailUser);
                }

                WriteInfoMessage ($LocalizedStrings.ConfiguringMailUser -f $mailUser);               
                
                WriteInfoMessage ($LocalizedStrings.DoneSyncingPublicFolderMailbox -f $displayName);
                Write-Host "";
            }
        }
    }

    if (-not $hasPublicFolderServingHierarchy)
    {
        WriteInfoMessage ($LocalizedStrings.NoHierarchyPublicFolderMailbox);
        Write-Host "";
    }

    foreach ($mailUser in $mailUserList)
    {
        $validExternalEmailAddresses += $mailUser.ExternalEmailAddress;
        Set-OrganizationConfig -RemotePublicFolderMailboxes @{Add=$mailUser};
    }   
}

#load hashtable of localized string
Import-LocalizedData -BindingVariable LocalizedStrings -FileName ImportPublicFolderMailboxes.strings.psd1

# Create a tenant PSSession against Exchange Online with modern auth.
CreateTenantSession;

WriteInfoMessage ($LocalizedStrings.StartedPublicFolderMailboxImport);
Write-Host "";

# Get mail enabled public folders in the organization
$publicFolderMailboxes = GetPublicFolderMailBoxes;

# Create mail enabled users for remote public folder mailboxes
SyncPublicFolderMailboxes $publicFolderMailboxes;

Write-Host "";
WriteInfoMessage ($LocalizedStrings.CompletedPublicFolderMailboxImport);

# Terminate the PSSession
Disconnect-ExchangeOnline -Confirm:$false -ErrorAction SilentlyContinue;