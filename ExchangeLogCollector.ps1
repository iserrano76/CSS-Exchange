<#
.NOTES
    Name: ExchangeLogCollector.ps1
    Author: David Paulson
    Requires: Powershell on an Exchange 2010+ Server with Adminstrator rights
    Version History:
    2.0.0 - Major updates have been made to the script and a new publish of it was done.
    2.0.1 - Missing HTTPErr Logs to the script.
    2.0.2 - Fix Bug with loading EMS and search directory
    2.0.3 - Switch "ClusterLogs" to "High_Availabilty_Logs" and adjust the switch as well to avoid confusion.
            Added a feature that checks to see if you pass some switches or it throws a warning asking are you sure to continue. 
    2.1 - Major Updates. Remote Collection now possible with -Server switch. Moved over to github. 
    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
	BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
	NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
	DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
	OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
.SYNOPSIS
    Collects the requested logs off the Exchange server based off the switches that are used.
.DESCRIPTION
    Collects the requested logs off the Exchange server based off the switches that are used.
.PARAMETER FilePath
    The Location of where you would like the data to be copied over to 
.PARAMETER Servers
    An array of servers that you would like to collect data from.
.PARAMETER EWSLogs
    Will collect the EWS logs from the Exchange Server 
.PARAMETER IISLogs 
    Will Collect the IIS Logs from the Exchange Server 
.PARAMETER DailyPerformanceLogs
    Used to collect Exchange 2013+ Daily Performance Logs 
.PARAMETER ManagedAvailability 
    Used to collect managed Availability Logs from the Exchange 2013+ Server
.PARAMETER Experfwiz 
    Used to collect Experfwiz data from the server 
.PARAMETER RPCLogs
    Used to collect the PRC Logs from the server 
.PARAMETER EASLogs
    Used to collect the Exchange Active Sync Proxy Logs 
.PARAMETER ECPLogs
    Used to collect the ECP Logs from the Exchange server 
.PARAMETER AutoDLogs
    Used to Collect the AutoD Logs from the server 
.PARAMETER OWALogs
    Used to collect the OWA Logs from the server 
.PARAMETER ADDriverLogs
    Used to collect the AD Driver Logs from the server 
.PARAMETER SearchLogs
    Used to collect the Search Logs from the server 
.PARAMETER HighAvailabilityLogs
    Used to collect the High Availability Information from the server 
.PARAMETER MapiLogs
    Used to collect the Mapi Logs from the server 
.PARAMETER MessageTrackingLogs
    Used to collect the Message Tracking Logs from the server 
.PARAMETER HubProtocolLogs
    Used to collect the Hub Protocol Logs from the server
.PARAMETER HubConnectivityLogs
    Used to collect the Hub Connectivity Logs from the server 
.PARAMETER FrontEndConnectivityLogs
    Used to collect the Front End Connectivity Logs from the server 
.PARAMETER FrontEndProtocolLogs
    Used to collect the Front End Protocol Logs from the server 
.PARAMETER MailboxConnectivityLogs
    Used to collect the Mailbox Connectivity Logs from the server 
.PARAMETER MailboxProtocolLogs
    Used to collect the Mailbox Protocol Logs from the server 
.PARAMETER QueueInformationThisServer
    Used to collect the Queue Information from this server 
.PARAMETER ReceiveConnectors
    Used to collect the Receive Connector information from this server 
.PARAMETER SendConnectors
    Used to collect the Send connector information from the Org 
.PARAMETER DAGInformation
    Used to collect the DAG Information for this DAG
.PARAMETER GetVdirs
    Used to collect the Virtual Directories of the environment 
.PARAMETER OrganizationConfig
    Used to collect the Organization Configuration from the environment. 
.PARAMETER TransportConfig
    Used to collect the Transport Configuration from this Exchange Server 
.PARAMETER DefaultTransportLogging
    Used to Get all the default logging that is enabled on the Exchange Server for Transport Information 
.PARAMETER Exmon 
    Used to Collect the Exmon Information 
.PARAMETER ServerInfo
    Used to collect the general Server information from the server 
.PARAMETER ExchangeServerInfo 
    Used to collect Exchange Server data (Get-ExchangeServer, Get-MailboxServer...). Enabled whenever ServerInfo is used as well.
.PARAMETER MSInfo 
    Old switch that was used for collecting the general Server information 
.PARAMETER CollectAllLogsBasedOnDaysWorth
    Used to collect some of the default logging based off Days Worth vs the whole directory 
.PARAMETER DiskCheckOverride
    Used to over the Availalbe Disk space required in order this script to run 
.PARAMETER AppSysLogs
    Used to collect the Application and System Logs. Default is set to true
.PARAMETER AllPossibleLogs
    Switch to enable all default logging enabled on the Exchange server. 
.PARAMETER NoZip
    Used to not zip up the data by default 
.PARAMETER SkipEndCopyOver
    Boolean to prevent the copy over after a remote collection.
#PARAMETER CustomData                             - Might bring this back in later build. 
    Used to collect data from a custom directory 
#PARAMETER CustomDataDirectory 
    Tell which directory you would like to collect data from 
.PARAMETER DaysWorth
    To determine how far back we would like to collect data from 
.PARAMETER ScriptDebug
    To enable Debug Logging for the script to determine what might be wrong with the script 
.PARAMETER DatabaseFailoverIssue
    To enable the common switches to assist with determine the cause of database failover issues 
.PARAMETER Experfwiz_LogmanName
    To be able to set the Experfwiz Logman Name that we would be looking for. By Default "Exchange_Perfwiz"
.PARAMETER Exmon_LogmanName
    To be able to set the Exmon Logman Name that we would be looking for. By Default "Exmon_Trace"
.PARAMETER AcceptEULA
    Switch used to bypass the disclaimer confirmation 

#>

#Parameters 
[CmdletBinding()]
Param (

[string]$FilePath = "C:\MS_Logs_Collection",
[Array]$Servers,
[switch]$EWSLogs,
[switch]$IISLogs,
[switch]$DailyPerformanceLogs,
[switch]$ManagedAvailability,
[switch]$Experfwiz,
[switch]$RPCLogs,
[switch]$EASLogs,
[switch]$ECPLogs,
[switch]$AutoDLogs,
[switch]$OWALogs,
[switch]$ADDriverLogs,
[switch]$SearchLogs,
[switch]$HighAvailabilityLogs,
[switch]$MapiLogs,
[switch]$MessageTrackingLogs,
[switch]$HubProtocolLogs,
[switch]$HubConnectivityLogs,
[switch]$FrontEndConnectivityLogs,
[switch]$FrontEndProtocolLogs,
[switch]$MailboxConnectivityLogs,
[switch]$MailboxProtocolLogs,
[switch]$QueueInformationThisServer,
[switch]$ReceiveConnectors,
[switch]$SendConnectors,
[switch]$DAGInformation,
[switch]$GetVdirs,
[switch]$OrganizationConfig,
[switch]$TransportConfig,
[switch]$DefaultTransportLogging,
[switch]$Exmon,
[switch]$ServerInfo,
[switch]$ExchangeServerInfo,
[switch]$CollectAllLogsBasedOnDaysWorth = $false, 
[switch]$DiskCheckOverride,
[switch]$AppSysLogs = $true,
[switch]$AllPossibleLogs,
[switch]$NoZip,
[bool]$SkipEndCopyOver,
[int]$DaysWorth = 3,
[switch]$DatabaseFailoverIssue,
[string]$Experfwiz_LogmanName = "Exchange_Perfwiz",
[string]$Exmon_LogmanName = "Exmon_Trace",
[switch]$AcceptEULA,
[switch]$ScriptDebug

)

$scriptVersion = 2.6

###############################################
#                                             #
#              Local Functions                #
#                                             #
###############################################

#disclaimer 
Function Display-Disclaimer {
$display = @"

    Exchange Log Collector v{0}

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING
    BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
    NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
    DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

    -This script will copy over data based off the switches provied. 
    -We will check for at least 15 GB of free space at the local target directory BEFORE 
        attempting to copy over the data.
    -Please run this script at your own risk. 

"@ -f $scriptVersion

    Clear-Host
    Write-Host $display
    do{
        if(-not $AcceptEULA)
        {
            $r = Read-Host ("Do you wish to continue ('y' or 'n')")
        }
        else{
            $r = "y"
        }
    }while(($r -ne "y" -and $r -ne "n"))
    if(-not ($AcceptEULA) -and $r -eq "n")
    {
        exit 
    }
    
}

Function Display-FeedBack {
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host "Looks like the script is done. If you ran into any issues or have additional feedback, please feel free to reach out dpaul@microsoft.com."
}


#Function to load the EXShell 
Function Load-ExShell {
    
        if($exinstall -eq $null){
        $testV14 = Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup'
        $testV15 = Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
    
        if($testV14){
            $Script:exinstall = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath	
        }
        elseif ($testV15) {
            $Script:exinstall = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath	
        }
        else{
            Write-Host "It appears that you are not on an Exchange 2010 or newer server. Sorry I am going to quit."
            exit
        }
    
        $script:exbin = $Script:exinstall + "\Bin"
    
        Write-Host "Loading Exchange PowerShell Module..."
        add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010
        }
    }
    

#Function to test if you are an admin on the server 
Function Is-Admin {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )
    If( $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )) {
        return $true
    }
    else {
        return $false
    }
}

Function Display-ScriptDebug{
param(
[Parameter(Mandatory=$true)]$stringdata 
)
    if($ScriptDebug)
    {
        Write-Host("[Script Debug] : {0}" -f $stringdata) -ForegroundColor Cyan
    }
}

Function Get-ZipEnabled {
    
    if($NoZip){return $false}
    else{return $true}
}

Function Get-TransportLoggingInformationPerServer {
param(
[string]$Server,
[int]$version 
)
    Display-ScriptDebug("Function Enter: Get-TransportLoggingInformationPerServer")
    Display-ScriptDebug("Passed - Server: {0} Version: {1}" -f $Server, $version)
    $hubObject = New-Object PSCustomObject
    $tranportLoggingObject = New-Object PSCustomObject
    if($version -ge 15)
    {
        #Hub Transport Layer 
        $data = Get-TransportService -Identity $Server
        $hubObject | Add-Member -MemberType NoteProperty -Name ConnectivityLogPath -Value ($data.ConnectivityLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name MessageTrackingLogPath -Value ($data.MessageTrackingLogPath.PathName) 
        $hubObject | Add-Member -MemberType NoteProperty -Name PipelineTracingPath -Value ($data.PipelineTracingPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name ReceiveProtocolLogPath -Value ($data.ReceiveProtocolLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name SendProtocolLogPath -Value ($data.SendProtocolLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name QueueLogPath -Value ($data.QueueLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name WlmLogPath -Value ($data.WlmLogPath.PathName)
        $tranportLoggingObject | Add-Member -MemberType NoteProperty -Name HubLoggingInfo -Value $hubObject

        #Front End Transport Layer 
        $FETransObject = New-Object PSCustomObject
        $data = Get-FrontendTransportService -Identity $Server
        $FETransObject | Add-Member -MemberType NoteProperty -Name ConnectivityLogPath -Value ($data.ConnectivityLogPath.PathName)
        $FETransObject | Add-Member -MemberType NoteProperty -Name ReceiveProtocolLogPath -Value ($data.ReceiveProtocolLogPath.PathName)
        $FETransObject | Add-Member -MemberType NoteProperty -Name SendProtocolLogPath -Value ($data.SendProtocolLogPath.PathName)
        $FETransObject | Add-Member -MemberType NoteProperty -Name AgentLogPath -Value ($data.AgentLogPath.PathName)
        $tranportLoggingObject | Add-Member -MemberType NoteProperty -Name FELoggingInfo -Value $FETransObject

        #Mailbox Transport Layer 
        $mbxObject = New-Object PSCustomObject
        $data = Get-MailboxTransportService -Identity $Server
        $mbxObject | Add-Member -MemberType NoteProperty -Name ConnectivityLogPath -Value ($data.ConnectivityLogPath.PathName)
        $mbxObject | Add-Member -MemberType NoteProperty -Name ReceiveProtocolLogPath -Value ($data.ReceiveProtocolLogPath.PathName)
        $mbxObject | Add-Member -MemberType NoteProperty -Name SendProtocolLogPath -Value ($data.SendProtocolLogPath.PathName)
        $mbxObject | Add-Member -MemberType NoteProperty -Name PipelineTracingPath -Value ($data.PipelineTracingPath.PathName)
        $mbxObject | Add-Member -MemberType NoteProperty -Name MailboxDeliveryThrottlingLogPath -Value ($data.MailboxDeliveryThrottlingLogPath.PathName)
        $tranportLoggingObject | Add-Member -MemberType NoteProperty -Name MBXLoggingInfo -Value $mbxObject 
        
    }

    elseif($version -eq 14)
    {
        $data = Get-TransportServer -Identity $Server
        $hubObject | Add-Member -MemberType NoteProperty -Name ConnectivityLogPath -Value ($data.ConnectivityLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name MessageTrackingLogPath -Value ($data.MessageTrackingLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name PipelineTracingPath -Value ($data.PipelineTracingPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name ReceiveProtocolLogPath -Value ($data.ReceiveProtocolLogPath.PathName)
        $hubObject | Add-Member -MemberType NoteProperty -Name SendProtocolLogPath -Value ($data.SendProtocolLogPath.PathName)
        $tranportLoggingObject | Add-Member -MemberType NoteProperty -Name HubLoggingInfo -Value $hubObject
    }

    else 
    {
        Write-Host("trying to determine transport information for server {0} and wasn't able to determine the correct version type" -f $Server)
        return     
    }

    Display-ScriptDebug("ReceiveConnectors: {0} QueueInformationThisServer: {1}" -f $ReceiveConnectors, $QueueInformationThisServer)
    if($ReceiveConnectors)
    {
        $value = Get-ReceiveConnector -Server $Server 
        $tranportLoggingObject | Add-Member -MemberType NoteProperty -Name ReceiveConnectorData -Value $value 
    }
    if($QueueInformationThisServer)
    {
        $value = Get-Queue -Server $Server 
        $tranportLoggingObject | Add-Member -MemberType NoteProperty -Name QueueData -Value $value 
    }

    Display-ScriptDebug("Function Exit: Get-TransportLoggingInformationPerServer")
    return $tranportLoggingObject 
}

Function Get-ExchangeBasicServerObject {
param(
[Parameter(Mandatory=$true)][string]$ServerName
)
    Display-ScriptDebug("Function Enter: Get-ExchangeBasicServerObject")
    Display-ScriptDebug("Passed [string]ServerName: {0}" -f $ServerName)
    $oldErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
    $failure = $false
    try {
        $exchServerObject = New-Object PSCustomObject 
        $exchServerObject | Add-Member -MemberType NoteProperty -Name ServerName -Value $ServerName
        $getExchangeServer = Get-ExchangeServer $ServerName -ErrorAction Stop 
        $exchServerObject | Add-Member -MemberType NoteProperty -Name ExchangeServer -Value $getExchangeServer
    }
    catch {
        Write-Host("Failed to detect server {0} as an Exchange Server" -f $ServerName) -ForegroundColor Red
        $failure = $true 
    }
    finally {
        $ErrorActionPreference = $oldErrorAction
    }

    if($failure -eq $true)
    {
        return $failure
    }

    $exchAdminDisplayVersion = $exchServerObject.ExchangeServer.AdminDisplayVersion
    $exchServerRole = $exchServerObject.ExchangeServer.ServerRole 
    Display-ScriptDebug("AdminDisplayVersion: {0} | ServerRole: {1}" -f $exchAdminDisplayVersion.ToString(), $exchServerRole.ToString())

    if($exchAdminDisplayVersion.Major -eq 14)
    {
        $exchVersion = 14
    }
    elseif($exchAdminDisplayVersion.Major -eq 15)
    {
        #determine if 2013/2016/2019
        if($exchAdminDisplayVersion.Minor -eq 0)
        {
            $exchVersion = 15
        }
        elseif($exchAdminDisplayVersion.Minor -eq 1)
        {
            $exchVersion = 16
        }
        else
        {
            $exchVersion = 19
        }
    }
    else
    {
        Write-Host("Failed to determine what version server {0} is. AdminDisplayVersion: {1}." -f $ServerName, $exchAdminDisplayVersion.ToString()) -ForegroundColor Red
        return $true 
    }

    Function Confirm-MailboxServer{
    param([string]$value)
        if($value -like "*Mailbox*"){return $true} else{ return $false}
    }

    Function Confirm-CASServer{
    param([string]$value,[int]$version)
        if(($version -ge 16) -or ($value -like "*ClientAccess*")){return $true} else{return $false}
    }

    Function Confirm-HubServer {
    param([string]$value,[int]$version)
        if(($version -ge 15) -or ($value -like "*HubTransport*")){return $true}{return $false}
    }

    Function Confirm-DAGMember{
    param([bool]$IsMailbox,[string]$ServerName)
        if($IsMailbox)
        {
            if((Get-MailboxServer $ServerName).DatabaseAvailabilityGroup -ne $null){return $true}
            else{return $false}
        }
        else {
            return $false
        }
    }

    $exchServerObject | Add-Member -MemberType NoteProperty -Name Mailbox -Value (Confirm-MailboxServer -value $exchServerRole)
    $exchServerObject | Add-Member -MemberType NoteProperty -Name CAS -Value (Confirm-CASServer -value $exchServerRole -version $exchVersion)
    $exchServerObject | Add-Member -MemberType NoteProperty -Name Hub -Value (Confirm-HubServer -value $exchServerRole -version $exchVersion)
    $exchServerObject | Add-Member -MemberType NoteProperty -Name Version -Value $exchVersion 
    $exchServerObject | Add-Member -MemberType NoteProperty -Name DAGMember -Value (Confirm-DAGMember -IsMailbox $exchServerObject.Mailbox -ServerName $exchServerObject.ServerName)

    Display-ScriptDebug("Confirm-MailboxServer: {0} | Confirm-CASServer: {1} | Confirm-HubServer: {2} | Confirm-DAGMember {3} | Version: {4} | AnyTransportSwitchesEnabled: {5}" -f $exchServerObject.Mailbox,
    $exchServerObject.CAS,
    $exchServerObject.Hub,
    $exchServerObject.DAGMember,
    $exchServerObject.Version,
    $Script:AnyTransportSwitchesEnabled
    )

    return $exchServerObject
}

Function Get-ServerObjects {
param(
[Parameter(Mandatory=$true)][Array]$ValidServers
)
    
    Display-ScriptDebug ("Function Enter: Get-ServerObjects")
    Display-ScriptDebug ("Passed {0} of Servers" -f $ValidServers.Count)
    $svrsObject = @()
    $validServersList = @() 
    $oldErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
    foreach($svr in $ValidServers)
    {
        Display-ScriptDebug -stringdata ("Working on Server {0}" -f $svr)

        $sobj = Get-ExchangeBasicServerObject -ServerName $svr
        if($sobj -eq $true)
        {
            Write-Host("Removing Server {0} from the list" -f $svr) -ForegroundColor Red 
            continue
        }
        else 
        {
            $validServersList += $svr 
        }

        if($Script:AnyTransportSwitchesEnabled -and $sobj.Hub)
        {
            $sobj | Add-Member -Name TransportInfoCollect -MemberType NoteProperty -Value $true 
            $sobj | Add-Member -Name TransportInfo -MemberType NoteProperty -Value (Get-TransportLoggingInformationPerServer -Server $svr -version $sobj.Version )
        }
        else 
        {
            $sobj | Add-Member -Name TransportInfoCollect -MemberType NoteProperty -Value $false    
        }

        $svrsObject += $sobj 
    }
    $ErrorActionPreference = $oldErrorAction
    if (($svrsObject -eq $null) -or ($svrsObject.Count -eq 0))
    {
        Write-Host("Something wrong happened in Get-ServerObjects stopping script") -ForegroundColor Red
        exit 
    }
    #Set the valid servers 
    $Script:ValidServers = $validServersList
    Display-ScriptDebug("Function Exit: Get-ServerObjects")
    Return $svrsObject
}

Function Get-ArgumentList {
param(
[Parameter(Mandatory=$true)][array]$Servers 
)
    
    $obj = New-Object PSCustomObject 
    $obj | Add-Member -Name FilePath -MemberType NoteProperty -Value $FilePath
    $obj | Add-Member -Name RootFilePath -MemberType NoteProperty -Value $Script:RootFilePath
    $obj | Add-Member -Name ManagedAvailability -MemberType NoteProperty -Value $ManagedAvailability
    $obj | Add-Member -Name Zip -MemberType NoteProperty -Value (Get-ZipEnabled)
    $obj | Add-Member -Name AppSysLogs -MemberType NoteProperty -Value $AppSysLogs
    $obj | Add-Member -Name EWSLogs -MemberType NoteProperty -Value $EWSLogs
    $obj | Add-Member -Name DailyPerformanceLogs -MemberType NoteProperty -Value $DailyPerformanceLogs
    $obj | Add-Member -Name RPCLogs -MemberType NoteProperty -Value $RPCLogs 
    $obj | Add-Member -Name EASLogs -MemberType NoteProperty -Value $EASLogs 
    $obj | Add-Member -Name ECPLogs -MemberType NoteProperty -Value $ECPLogs 
    $obj | Add-Member -Name AutoDLogs -MemberType NoteProperty -Value $AutoDLogs
    $obj | Add-Member -Name OWALogs -MemberType NoteProperty -Value $OWALogs
    $obj | Add-Member -Name ADDriverLogs -MemberType NoteProperty -Value $ADDriverLogs
    $obj | Add-Member -Name SearchLogs -MemberType NoteProperty -Value $SearchLogs
    $obj | Add-Member -Name HighAvailabilityLogs -MemberType NoteProperty -Value $HighAvailabilityLogs
    $obj | Add-Member -Name MapiLogs -MemberType NoteProperty -Value $MapiLogs
    $obj | Add-Member -Name MessageTrackingLogs -MemberType NoteProperty -Value $MessageTrackingLogs
    $obj | Add-Member -Name HubProtocolLogs -MemberType NoteProperty -Value $HubProtocolLogs
    $obj | Add-Member -Name HubConnectivityLogs -MemberType NoteProperty -Value $HubConnectivityLogs
    $obj | Add-Member -Name FrontEndConnectivityLogs -MemberType NoteProperty -Value $FrontEndConnectivityLogs
    $obj | Add-Member -Name FrontEndProtocolLogs -MemberType NoteProperty -Value $FrontEndProtocolLogs
    $obj | Add-Member -Name MailboxConnectivityLogs -MemberType NoteProperty -Value $MailboxConnectivityLogs
    $obj | Add-Member -Name MailboxProtocolLogs -MemberType NoteProperty -Value $MailboxProtocolLogs
    $obj | Add-Member -Name QueueInformationThisServer -MemberType NoteProperty -Value $QueueInformationThisServer
    $obj | Add-Member -Name ReceiveConnectors -MemberType NoteProperty -Value $ReceiveConnectors 
    $obj | Add-Member -Name SendConnectors -MemberType NoteProperty -Value $SendConnectors 
    $obj | Add-Member -Name DAGInformation -MemberType NoteProperty -Value $DAGInformation 
    $obj | Add-Member -Name GetVdirs -MemberType NoteProperty -Value $GetVdirs 
    $obj | Add-Member -Name TransportConfig -MemberType NoteProperty -Value $TransportConfig
    $obj | Add-Member -Name DefaultTransportLogging -MemberType NoteProperty -Value $DefaultTransportLogging
    $obj | Add-Member -Name ServerInfo -MemberType NoteProperty -Value $ServerInfo
    $obj | Add-Member -Name CollectAllLogsBasedOnDaysWorth -MemberType NoteProperty -Value $CollectAllLogsBasedOnDaysWorth
    $obj | Add-Member -Name DaysWorth -MemberType NoteProperty -Value $DaysWorth 
    $obj | Add-Member -Name IISLogs -MemberType NoteProperty -Value $IISLogs 
    $obj | Add-Member -Name AnyTransportSwitchesEnabled -MemberType NoteProperty -Value $script:AnyTransportSwitchesEnabled
    $svrobjs = Get-ServerObjects -ValidServers $Servers
    $obj | Add-Member -Name ServerObjects -MemberType NoteProperty -Value $svrobjs
    $obj | Add-Member -Name HostExeServerName -MemberType NoteProperty -Value ($env:COMPUTERNAME)
    $obj | Add-Member -Name Experfwiz -MemberType NoteProperty -Value $Experfwiz
    $obj | Add-Member -Name Experfwiz_LogmanName -MemberType NoteProperty -Value $Experfwiz_LogmanName
    $obj | Add-Member -Name Exmon -MemberType NoteProperty -Value $Exmon
    $obj | Add-Member -Name Exmon_LogmanName -MemberType NoteProperty -Value $Exmon_LogmanName
    $obj | Add-Member -Name ScriptDebug -MemberType NoteProperty -Value $ScriptDebug
    $obj | Add-Member -Name ExchangeServerInfo -MemberType NoteProperty -Value $ExchangeServerInfo
    
    #Collect only if enabled we are going to just keep it on the base of the passed parameter object to make it simple 
    $mbx = $false
    foreach($svr in $svrobjs)
    {
        if($svr.ServerName -eq $env:COMPUTERNAME)
        {
            $mbx = $true
            $checkSvr = $svr
        }
    }
    if(($mbx) -and ($HighAvailabilityLogs) -and ($checkSvr.DAGMember))
    {
        Write-Host("Generating cluster logs for the local server's DAG only")
        Write-Host("Server: {0}" -f $checkSvr.ServerName)
        #Only going to do this for the local server's DAG 
        $cmd = "Cluster log /g"
        Invoke-Expression -Command $cmd | Out-Null
    }
    if($SendConnectors)
    {
        #TODO move this to a different location, but for now this should work. 
        $value = Get-SendConnector 
        $Script:SendConnectorData = $value
        #$obj | Add-Member -MemberType NoteProperty -Name SendConnectorData -Value $value
    }

    
    Return $obj 
}

Function Test-PossibleCommonScenarios {

    #all possible logs 
    if($AllPossibleLogs)
    {
        $Script:EWSLogs = $true 
        $Script:IISLogs = $true 
        $Script:DailyPerformanceLogs = $true
        $Script:ManagedAvailability = $true 
        $Script:RPCLogs = $true 
        $Script:EASLogs = $true 
        $Script:AutoDLogs = $true
        $Script:OWALogs = $true 
        $Script:ADDriverLogs = $true 
        $Script:SearchLogs = $true
        $Script:HighAvailabilityLogs = $true
        $Script:ServerInfo = $true 
        $Script:GetVdirs = $true
        $Script:DAGInformation = $true 
        $Script:DefaultTransportLogging = $true
        $Script:MapiLogs = $true 
        $Script:OrganizationConfig = $true
        $Script:ECPLogs = $true
        $Script:ExchangeServerInfo = $true
    }

    if($DefaultTransportLogging)
    {
        $Script:HubConnectivityLogs = $true 
        $Script:MessageTrackingLogs = $true 
        $Script:QueueInformationThisServer = $true 
        $Script:SendConnectors = $true 
        $Script:ReceiveConnectors = $true 
        $Script:TransportConfig = $true
        $Script:FrontEndConnectivityLogs = $true 
        $Script:MailboxConnectivityLogs = $true 
        $Script:FrontEndProtocolLogs = $true 
    }

    if($DatabaseFailoverIssue)
    {
        $Script:DailyPerformanceLogs = $true
        $Script:HighAvailabilityLogs = $true 
        $Script:ManagedAvailability = $true 
        $Script:DAGInformation = $true
    }
    
    #See if any transport logging is enabled. 
    $Script:AnyTransportSwitchesEnabled = $false
    if($HubProtocolLogs -or 
    $HubConnectivityLogs -or 
    $MessageTrackingLogs -or
    $QueueInformationThisServer -or
    $SendConnectors -or 
    $ReceiveConnectors -or
    $TransportConfig -or
    $FrontEndConnectivityLogs -or 
    $FrontEndProtocolLogs -or 
    $MailboxConnectivityLogs -or
    $MailboxProtocolLogs -or 
    $DefaultTransportLogging){$Script:AnyTransportSwitchesEnabled = $true}

    if($ServerInfo)
    {
        $Script:ExchangeServerInfo = $true 
    }

}

Function Test-NoSwitchesProvided {
    if($EWSLogs -or 
    $IISLogs -or 
    $DailyPerformanceLogs -or 
    $ManagedAvailability -or 
    $Experfwiz -or 
    $RPCLogs -or 
    $EASLogs -or 
    $ECPLogs -or 
    $AutoDLogs -or 
    $SearchLogs -or 
    $OWALogs -or 
    $ADDriverLogs -or
    $HighAvailabilityLogs -or
    $MapiLogs -or 
    $Script:AnyTransportSwitchesEnabled -or
    $DAGInformation -or
    $GetVdirs -or 
    $OrganizationConfig -or
    $Exmon -or 
    $ServerInfo -or
    $ExchangeServerInfo
    ){return}
    else 
    {
        Write-Host ""    
        Write-Warning "Doesn't look like any parameters were provided, are you sure you are running the correct command? This is ONLY going to collect the Application and System Logs."
        Write-Warning "Enter 'y' to continue and 'n' to stop the script"
        do{
            $a = Read-Host "Please enter 'y' or 'n'"
        }while($a -ne 'y' -and $a -ne 'n')
        if($a -eq 'n'){exit}
        else{
            Write-Host "Okay moving on..."
        }
    }
}

Function Test-RemoteExecutionOfServers {
param(
[Parameter(Mandatory=$true)][Array]$Server_List
)
    Display-ScriptDebug("Function Enter: Test-RemoteExecutionOfServers")
    $Servers_up = @() 
    Write-Host "Checking to see if the servers are up in this list:"
    foreach($server in $Server_List) {Write-Host $server}
    Write-Host ""
    Write-Host "Checking their status...."
    foreach($server in $Server_List)
    {
        Write-Host("Checking server {0}....." -f $server) -NoNewline
        if((Test-Connection $server -Quiet))
        {   
            Write-Host "Online" -ForegroundColor Green
            $Servers_up += $server
        }
        else 
        {
            Write-Host "Offline" -ForegroundColor Red
            Write-Host ("Removing Server {0} from the list to collect data from" -f $server)
        }
    }
    #Now we should check to see if can use WRM with invoke-command
    Write-Host ""
    Write-Host "For all the servers that are up, we are going to see if remote execution will work"
    #shouldn't need to test if they are Exchange servers, as we should be doing that locally as well. 
    $valid_Servers = @()
    $oldErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
    foreach($server in $Servers_up)
    {

        try {
            Write-Host("Checking Server {0}....." -f $server) -NoNewLine
            Invoke-Command -ComputerName $server -ScriptBlock { Get-Process | Out-Null}
            #if that doesn't fail, we should be okay to add it to the working list 
            Write-Host("Passed") -ForegroundColor Green
            $valid_Servers += $server
        }
        catch {
            Write-Host("Failed") -ForegroundColor Red
            Write-Host("Removing Server {0} from the list to collect data from" -f $server)
        }
    }
    Display-ScriptDebug("Function Exit: Test-RemoteExecutionOfServers")
    $ErrorActionPreference = $oldErrorAction
    return $valid_Servers 
}



Function Get-VdirsLDAP {

$authTypeEnum = @" 
namespace AuthMethods  
{
    using System;
    [Flags]
    public enum AuthenticationMethodFlags
    {
        None = 0,
        Basic = 1,
        Ntlm = 2,
        Fba = 4,
        Digest = 8,
        WindowsIntegrated = 16,
        LiveIdFba = 32,
        LiveIdBasic = 64,
        WSSecurity = 128,
        Certificate = 256,
        NegoEx = 512,
        // Exchange 2013
        OAuth = 1024,
        Adfs = 2048,
        Kerberos = 4096,
        Negotiate = 8192,
        LiveIdNegotiate = 16384,
    }
}
"@
    
    Write-Host "Collecting Virtual Directory Information..."
    Add-Type -TypeDefinition $authTypeEnum -Language CSharp
    
    $objRootDSE = [ADSI]"LDAP://rootDSE"
    $strConfigurationNC = $objRootDSE.configurationNamingContext
    $objConfigurationNC = New-object System.DirectoryServices.DirectoryEntry("LDAP://$strConfigurationNC")
    $searcher = new-object DirectoryServices.DirectorySearcher
    $searcher.filter = "(&(objectClass=msExchVirtualDirectory)(!objectClass=container))" 
    $searcher.SearchRoot = $objConfigurationNC
    $Searcher.CacheResults = $false  
    $Searcher.SearchScope = "Subtree"
    $Searcher.PageSize = 1000  
    
    # Get all the results
    $colResults  = $searcher.FindAll()
    
    $objects = @()
    
    # Loop through the results and
    foreach ($objResult in $colResults)
    {
        $objItem = $objResult.getDirectoryEntry()
        $objProps = $objItem.Properties
        
        $place = $objResult.Path.IndexOf("CN=Protocols,CN=")
        $ServerDN = [ADSI]("LDAP://" + $objResult.Path.SubString($place,($objResult.Path.Length - $place)).Replace("CN=Protocols,",""))
        [string]$Site = $serverDN.Properties.msExchServerSite.ToString().Split(",")[0].Replace("CN=","")
        [string]$server = $serverDN.Properties.adminDisplayName.ToString()
        [string]$version = $serverDN.Properties.serialNumber.ToString()
        
    
        $obj = New-Object PSObject 
        $obj | Add-Member -MemberType NoteProperty -name Server -value $server
        $obj | Add-Member -MemberType NoteProperty -name Version -value $version
        $obj | Add-Member -MemberType NoteProperty -name Site -value $Site
        [string]$var = $objProps.DistinguishedName.ToString().Split(",")[0].Replace("CN=","")
        $obj | Add-Member -MemberType NoteProperty -name VirtualDirectory -value $var
        [string]$var = $objProps.msExchInternalHostName
        $obj | Add-Member -MemberType NoteProperty -name InternalURL -value $var
        
        if (-not [string]::IsNullOrEmpty($objProps.msExchInternalAuthenticationMethods))
        {
            $obj | Add-Member -MemberType NoteProperty -name InternalAuthenticationMethods -value ([AuthMethods.AuthenticationMethodFlags]$objProps.msExchInternalAuthenticationMethods)
        }
        else
        {
            $obj | Add-Member -MemberType NoteProperty -name InternalAuthenticationMethods -value $null
        }
        
        [string]$var = $objProps.msExchExternalHostName
        $obj | Add-Member -MemberType NoteProperty -name ExternalURL -value $var 
    
        if (-not [string]::IsNullOrEmpty($objProps.msExchExternalAuthenticationMethods))
        {
            $obj | Add-Member -MemberType NoteProperty -name ExternalAuthenticationMethods -value ([AuthMethods.AuthenticationMethodFlags]$objProps.msExchExternalAuthenticationMethods)
        }
        else
        {
            $obj | Add-Member -MemberType NoteProperty -name ExternalAuthenticationMethods -value $null
        }
        
        if (-not [string]::IsNullOrEmpty($objProps.msExch2003Url))
        {
            [string]$var = $objProps.msExch2003Url
            $obj | Add-Member -MemberType NoteProperty -name Exchange2003URL  -value $var
        }
        else
        {
            $obj | Add-Member -MemberType NoteProperty -name Exchange2003URL -value $null
        }
        
        [Array]$objects += $obj
    }
    
    return $objects 
}

Function Get-ExchangeServerDAGName {
param(
[string]$Server 
)
    Display-ScriptDebug("Function Enter: Get-ExchangeServerDAGName")
    $oldErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
    try {
        $dagName = (Get-MailboxServer $Server).DatabaseAvailabilityGroup.Name 
        Display-ScriptDebug("Returning dagName: {0}" -f $dagName)
        Display-ScriptDebug("Function Exit: Get-ExchangeServerDAGName")
        return $dagName
    }
    catch {
        Write-Host("Looks like this server {0} isn't a Mailbox Server. Unable to get DAG Infomration." -f $Server)
        return $null 
    }
    finally
    {
        $ErrorActionPreference = $oldErrorAction 
    }
}

Function Get-MailboxDatabaseInformationFromDAG{
param(
[parameter(Mandatory=$true)]$DAGInfo
)
    Display-ScriptDebug("Function Enter: Get-MailboxDatabaseInformationFromDAG")
    Write-Host("Getting Database information from {0} DAG member servers" -f $DAGInfo.Name)
    $AllDupMDB = @()
    foreach($serverobj in $DAGInfo.Servers)
    {
        foreach($server in $serverobj.Name)
        {
            $AllDupMDB += Get-MailboxDatabase -Server $server -Status 
        }
    }
    #remove all dups 
    $MailboxDBS = @()
    foreach($t_mdb in $AllDupMDB)
    {
        $add = $true
        foreach($mdb in $MailboxDBS)
        {
            if($mdb.Name -eq $t_mdb.Name)
            {
                $add = $false
                break
            }
        }
        if($add)
        {
            $MailboxDBS += $t_mdb
        }
    }

    Write-Host("Found the following databases:")
    foreach($mdb in $MailboxDBS)
    {
        Write-Host($mdb)
    }

    $MailboxDBInfo = @() 

    foreach($mdb in $MailboxDBS)
    {
        $mdb_Name = $mdb.Name 
        $dbObj = New-Object PSCustomObject
        $dbObj | Add-Member -MemberType NoteProperty -Name MDBName -Value $mdb_Name
        $dbObj | Add-Member -MemberType NoteProperty -Name MDBInfo -Value $mdb
        $value = Get-MailboxDatabaseCopyStatus $mdb_Name\*
        $dbObj | Add-Member -MemberType NoteProperty -Name MDBCopyStatus -Value $value
        $MailboxDBInfo += $dbObj
    }
    Display-ScriptDebug("Function Exit: Get-MailboxDatabaseInformationFromDAG")
    return $MailboxDBInfo
}

Function Get-DAGInformation {

    $DAGName = Get-ExchangeServerDAGName -Server $env:COMPUTERNAME #only going to get the local server's DAG info
    if($DAGName -ne $null)
    {
        $dagObj = New-Object PSCustomObject
        $value = Get-DatabaseAvailabilityGroup $DAGName -Status 
        $dagObj | Add-Member -MemberType NoteProperty -Name DAGInfo -Value $value 
        $value = Get-DatabaseAvailabilityGroupNetwork $DAGName 
        $dagObj | Add-Member -MemberType NoteProperty -Name DAGNetworkInfo -Value $value
        $dagObj | Add-Member -MemberType NoteProperty -Name AllMdbs -Value (Get-MailboxDatabaseInformationFromDAG -DAGInfo $dagObj.DAGInfo)
        return $dagObj
    }
}

#Logic for determining the free space on the drive 
Function Get-FreeSpaceFromDrives {
param(
[Parameter(Mandatory=$true)][string]$Root_Full_Path,
[Parameter(Mandatory=$true)][Array]$Drives_WMI
)
    $driveLetter = ($Root_Full_Path.Split("\"))[0]
    $Free_Space = $Drives_WMI | ?{$_.DriveLetter -eq $driveLetter} | select DriveLetter, label, @{LABEL='GBfreespace';EXPRESSION={$_.freespace/1GB}}
    return $Free_Space
}

Function Get-DisksData {
    
    $drives = gwmi win32_volume -Filter 'drivetype = 3'
    $obj = New-Object PSCustomObject
    $obj | Add-Member -MemberType NoteProperty -Name Drives -Value $drives
    $obj | Add-Member -MemberType NoteProperty -Name ServerName -Value ($env:COMPUTERNAME)
    return $obj
}

Function Test-DiskSpace {
param(
[Parameter(Mandatory=$true)][array]$Servers,
[Parameter(Mandatory=$true)][string]$Path,
[Parameter(Mandatory=$true)][int]$CheckSize
)
    Display-ScriptDebug("Function Enter: Test-DiskSpace")
    Display-ScriptDebug("Passed - Path: {0} CheckSize: {1}" -f $Path, $CheckSize)
    Write-Host("Checking the free space on the servers before collecting the data...")

    $script_block = ${Function:Get-DisksData}
    $Servers_Data = Invoke-Command -ComputerName $Servers -ScriptBlock $script_block
    $Passed_Servers = @()

    foreach($server in $Servers_Data)
    {
        $Free_Space = Get-FreeSpaceFromDrives -Root_Full_Path $Path -Drives_WMI $server.Drives
        if($Free_Space.GBfreespace -gt $CheckSize)
        {
            Write-Host("[{0}] : We have more than {1} GB of free space at {2}" -f $server.ServerName, $CheckSize, $Path)
            $Passed_Servers += $server.ServerName
        }
        else 
        {
            Write-Host("[{0}] : We have less than {1} GB of free space on {2}" -f $server.ServerName, $CheckSize, $Path)
        }
    }

    if($Passed_Servers.Count -ne $Servers.Count)
    {
        Write-Host("Looks like all the servers didn't pass the disk space check.")
        Write-Host("We will only collect data from these servers: ")
        foreach($svr in $Passed_Servers)
        {
            Write-Host("{0}" -f $svr)
        }
        do{
            $r = Read-Host("Are you sure you want to continue? 'y' or 'n'")
        }while($r -ne 'y' -and $r -ne 'n')
        if($r -eq 'n')
        {
            exit 
        }
    }
    Display-ScriptDebug("Function Exit: Test-DiskSpace")
    return $Passed_Servers
}

Function Get-RemoteLogLocation {
param(
[parameter(Mandatory=$true)][array]$Servers,
[parameter(Mandatory=$true)][string]$RootPath 
)
    Display-ScriptDebug("Calling: Get-RemoteLogLocation")
    Display-ScriptDebug("Passed: Number of servers {0} | [string]RootPath:{1}" -f $Servers.Count, $RootPath)
    Function Get-ZipLocation 
    {
        param(
        [parameter(Mandatory=$true)][string]$RootPath
        )
        $like = "*-{0}*.zip" -f (Get-Date -Format Md)
        $Item = $RootPath + (Get-ChildItem $RootPath | ?{$_.Name -like $like} | sort CreationTime -Descending)[0]
        
        $obj = New-Object -TypeName PSCustomObject 
        $obj | Add-Member -MemberType NoteProperty -Name ServerName -Value $env:COMPUTERNAME
        $obj | Add-Member -MemberType NoteProperty -Name ZipFolder -Value $Item
        $obj | Add-Member -MemberType NoteProperty -Name Size -Value ((Get-Item $Item).Length)
        return $obj
    }
    
    $script_block = ${function:Get-ZipLocation}
    $LogInfo = Invoke-Command -ComputerName $Servers -ScriptBlock $script_block -ArgumentList $RootPath 
    
    return $LogInfo
}

Function Test-DiskSpaceForCopyOver {
param(
[parameter(Mandatory=$true)][array]$LogPathObject,
[parameter(Mandatory=$true)][string]$RootPath 
)
    Display-ScriptDebug("Function Enter: Test-DiskSpaceForCopyOver")
    foreach($SvrObj in $LogPathObject)
    {
        $iTotalSize += $SvrObj.Size 
    }
    #switch it to GB in size 
    $iTotalSizeGB = $iTotalSize / 1GB
    #Get the local free space again 
    $driveObj = Get-DisksData
    $FreeSpace = Get-FreeSpaceFromDrives -Root_Full_Path $RootPath -Drives_WMI $driveObj.Drives
    $extraSpace = 10
    if($FreeSpace.GBfreespace -gt ($iTotalSizeGB + $extraSpace))
    {
        Write-Host("[{0}] : Looks like we have enough free space at the path to copy over the data" -f $env:COMPUTERNAME)
        Write-Host("[{0}] : FreeSpace: {1} TestSize: {2} Path: {3}" -f $env:COMPUTERNAME, $FreeSpace.GBfreespace, ($iTotalSizeGB + $extraSpace), $RootPath)
        return $true
    }
    else 
    {
        Write-Host("[{0}] : Looks like we don't have enough free space to copy over the data" -f $env:COMPUTERNAME) -ForegroundColor Yellow
        Write-Host("[{0}] : FreeSpace: {1} TestSize: {2} Path: {3}" -f $env:COMPUTERNAME, $FreeSpace.GBfreespace, ($iTotalSizeGB + $extraSpace), $RootPath)
        return $false
    }

}

Function Verify-LocalServerIsUsed {
param(
[Parameter(Mandatory=$true)]$Servers
)
    foreach($server in $Servers)
    {
        if($server -eq $env:COMPUTERNAME)
        {
            Display-ScriptDebug ("Local Server {0} is in the list" -f $server)
            return 
        }
    }

    Write-Host("The server that you are running the script from isn't in the list of servers that we are collecting data from, this is currently not supported. Stopping the script.") -ForegroundColor Yellow
    exit 
}
   

###############################################
#                                             #
#          Possible Remote Functions          #
#                                             #
###############################################

Function Remote-Functions {
param(
[Parameter(Mandatory=$true)][object]$PassedInfo
)

    
    #Template Master https://github.com/dpaulson45/PublicPowerShellScripts/blob/master/Functions/Compress-Folder/Compress-Folder.ps1
    Function Compress-Folder
    {
    [CmdletBinding()]
    param(
    [Parameter(Mandatory=$true)][string]$Folder,
    [Parameter(Mandatory=$false)][bool]$IncludeMonthDay = $false,
    [Parameter(Mandatory=$false)][bool]$IncludeDisplayZipping = $true,
    [Parameter(Mandatory=$false)][scriptblock]$VerboseFunctionCaller,
    [Parameter(Mandatory=$false)][scriptblock]$HostFunctionCaller
    )

    Function Write-VerboseWriter {
    param(
    [Parameter(Mandatory=$true)][string]$WriteString 
    )
        if($VerboseFunctionCaller -eq $null)
        {
            Write-Verbose $WriteString
        }
        else 
        {
            &$VerboseFunctionCaller $WriteString
        }
    }

    Function Write-HostWriter {
    param(
    [Parameter(Mandatory=$true)][string]$WriteString 
    )
        if($HostFunctionCaller -eq $null)
        {
            Write-Host $WriteString
        }
        else
        {
            &$HostFunctionCaller $WriteString    
        }
    }
    Function Enable-IOCompression
    {
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        $successful = $true 
        Write-VerboseWriter("Calling: Enable-IOCompression")
        try 
        {
            Add-Type -AssemblyName System.IO.Compression.Filesystem 
        }
        catch 
        {
            Write-HostWriter("Failed to load .NET Compression assembly. Unable to compress up the data.")
            $successful = $false 
        }
        finally 
        {
            $ErrorActionPreference = $oldErrorAction
        }
        Write-VerboseWriter("Returned: [bool]{0}" -f $successful)
        return $successful
    }   
    Function Confirm-IOCompression 
    {
        Write-VerboseWriter("Calling: Confirm-IOCompression")
        $assemblies = [Appdomain]::CurrentDomain.GetAssemblies()
        $successful = $false
        foreach($assembly in $assemblies)
        {
            if($assembly.Location -like "*System.IO.Compression.Filesystem*")
            {
                $successful = $true 
                break 
            }
        }
        Write-VerboseWriter("Returned: [bool]{0}" -f $successful)
        return $successful
    }

    Function Compress-Now
    {
        Write-VerboseWriter("Calling: Compress-Now ")
        $zipFolder = Get-ZipFolderName -Folder $Folder -IncludeMonthDay $IncludeMonthDay
        if($IncludeDisplayZipping)
        {
            Write-HostWriter("Compressing Folder {0}" -f $Folder)
        }
        $timer = [System.Diagnostics.Stopwatch]::StartNew()
        [System.IO.Compression.ZipFile]::CreateFromDirectory($Folder, $zipFolder)
        $timer.Stop()
        Write-VerboseWriter("Time took to compress folder {0} seconds" -f $timer.Elapsed.TotalSeconds)
        if((Test-Path -Path $zipFolder))
        {
            Write-VerboseWriter("Compress successful, removing folder.")
            Remove-Item $Folder -Force -Recurse 
        }
    }

    Function Get-ZipFolderName {
    param(
    [Parameter(Mandatory=$true)][string]$Folder,
    [Parameter(Mandatory=$false)][bool]$IncludeMonthDay = $false
    )
        Write-VerboseWriter("Calling: Get-ZipFolderName")
        Write-VerboseWriter("Passed - [string]Folder:{0} | [bool]IncludeMonthDay:{1}" -f $Folder, $IncludeMonthDay)
        if($IncludeMonthDay)
        {
            $zipFolderNoEXT = "{0}-{1}" -f $Folder, (Get-Date -Format Md)
        }
        else 
        {
            $zipFolderNoEXT = $Folder
        }
        Write-VerboseWriter("[string]zipFolderNoEXT: {0}" -f $zipFolderNoEXT)
        $zipFolder = "{0}.zip" -f $zipFolderNoEXT
        if(Test-Path $zipFolder)
        {
            [int]$i = 1
            do{
                $zipFolder = "{0}-{1}.zip" -f $zipFolderNoEXT,$i 
                $i++
            }while(Test-Path $zipFolder)
        }
        Write-VerboseWriter("Returned: [string]zipFolder {0}" -f $zipFolder)
        return $zipFolder
    }
    $passedVerboseFunctionCaller = $false
    $passedHostFunctionCaller = $false
    if($VerboseFunctionCaller -ne $null){$passedVerboseFunctionCaller = $true}
    if($HostFunctionCaller -ne $null){$passedHostFunctionCaller = $true}

    Write-VerboseWriter("Calling: Compress-Folder")
    Write-VerboseWriter("Passed - [string]Folder: {0} | [bool]IncludeDisplayZipping{1} | [scriptblock]VerboseFunctionCaller:{2} | [scriptblock]HostFunctionCaller:{3}" -f $Folder, 
    $IncludeDisplayZipping,
    $passedVerboseFunctionCaller,
    $passedHostFunctionCaller)

    if(Test-Path $Folder)
    {
        if(Confirm-IOCompression)
        {
            Compress-Now
        }
        else
        {
            if(Enable-IOCompression)
            {
                Compress-Now
            }
            else
            {
                Write-HostWriter("Unable to compress folder {0}" -f $Folder)
                Write-VerboseWriter("Unable to enable IO compression on this system")
            }
        }
    }
    else
    {
        Write-HostWriter("Failed to find the folder {0}" -f $Folder)
    }
    }

    Function New-FolderCreate {
    param(
    [string]$Folder
    )
        if(-not (Test-Path -Path $Folder))
        {
            Write-Host("[{0}] : Creating Directory {1}" -f $env:COMPUTERNAME, $Folder)
            [System.IO.Directory]::CreateDirectory($Folder) | Out-Null
        }
        else 
        {
            Write-Host("[{0}] : Directory {1} is already created!" -f $env:COMPUTERNAME, $Folder)
        }

    }

    Function Remote-DisplayScriptDebug {
    param(
    [Parameter(Mandatory=$true)]$stringdata 
    )
        if($PassedInfo.ScriptDebug)
        {
            Write-Host("[{0} - Script Debug] : {1}" -f $env:COMPUTERNAME, $stringdata) -ForegroundColor Cyan
        }
    }

    Function Remote-DisplayHostWriter{
    param(
    [Parameter(Mandatory=$true)][string]$WriteString 
    )
        Write-Host("[{0}] : {1}" -f $env:COMPUTERNAME, $WriteString)
    }

    
    Function Zip-Folder {
    param(
    [string]$Folder,
    [bool]$ZipItAll
    )
        if($ZipItAll)
        {
            Compress-Folder -Folder $Folder -IncludeMonthDay $true -VerboseFunctionCaller ${Function:Remote-DisplayScriptDebug} -HostFunctionCaller ${Function:Remote-DisplayHostWriter}
            
        }
        else 
        {
            Compress-Folder -Folder $Folder -VerboseFunctionCaller ${Function:Remote-DisplayScriptDebug} -HostFunctionCaller ${Function:Remote-DisplayHostWriter}
        }
    
    }
    

    Function Copy-FullLogFullPathRecurse {
    param(
    [Parameter(Mandatory=$true)][string]$LogPath,
    [Parameter(Mandatory=$true)][string]$CopyToThisLocation
    )   
        Remote-DisplayScriptDebug("Function Enter: Copy-FullLogFullPathRecurse")
        Remote-DisplayScriptDebug("Passed - LogPath: {0} CopyToThisLocation: {1}" -f $LogPath, $CopyToThisLocation)
        New-FolderCreate -Folder $CopyToThisLocation 
        if(Test-Path $LogPath)
        {
            Copy-Item $LogPath\* $CopyToThisLocation -Recurse
            Zip-Folder $CopyToThisLocation
        }
        else 
        {
            Remote-DisplayHostWriter("No Folder at {0}. Unable to copy this data." -f $LogPath)
            New-Item -Path ("{0}\NoFolderDetected.txt" -f $CopyToThisLocation) -ItemType File -Value $LogPath
        }
        Remote-DisplayScriptDebug("Function Exit: Copy-FullLogFullPathRecurse")
    }

    Function Copy-LogsBasedOnTime {
    param(
    [Parameter(Mandatory=$true)][string]$LogPath,
    [Parameter(Mandatory=$true)][string]$CopyToThisLocation
    )
        Remote-DisplayScriptDebug("Function Enter: Copy-LogsBasedOnTime")
        Remote-DisplayScriptDebug("Passed - LogPath: {0} CopyToThisLocation: {1}" -f $LogPath, $CopyToThisLocation)
        New-FolderCreate -Folder $CopyToThisLocation

        Function No-FilesInLocation {
        param(
        [Parameter(Mandatory=$true)][string]$CopyFromLocation,
        [Parameter(Mandatory=$true)][string]$CopyToLocation 
        )
            Write-Warning("[{0}] : It doesn't look like you have any data in this location {1}." -f $Script:LocalServerName, $CopyFromLocation)
            Write-Warning("[{0}] : You should look into the reason as to why, because this shouldn't occur." -f $Script:LocalServerName)
            #Going to place a file in this location so we know what happened
            $tempFile = $CopyToLocation + "\NoFilesDetected.txt"
            New-Item $tempFile -ItemType File -Value $LogPath 
            Start-Sleep 1
        }

        $date = (Get-Date).AddDays(0-$PassedInfo.DaysWorth)
        $copyFromDate = "$($Date.Month)/$($Date.Day)/$($Date.Year)"
        Remote-DisplayScriptDebug("Copy From Date: {0}" -f $copyFromDate)
        $SkipCopy = $false 
        #We are not copying files recurse so we need to not include possible directories or we will throw an error 
        $Files = Get-ChildItem $LogPath | Sort-Object LastWriteTime -Descending | ?{$_.LastWriteTime -ge $copyFromDate -and $_.Mode -notlike "d*"}
        #if we don't have any logs, we want to attempt to copy something 
        if($Files -eq $null)
        {
            #Write-Warning("[{0}] : Oops! Looks like I wasn't able to find what you are looking for, so I am going to attempt to collect the newest log for you" -f $Script:LocalServerName)
            <#
                There are a few different reasons to get here
                1. We don't have any files in the timeframe request in the directory that we are looking at
                2. We have sub directories that we need to look into and look at those files (Only if we don't have files in the currently location so we aren't pulling files like the index files from message tracking)
            #>
            #Debug
            Remote-DisplayScriptDebug("Copy-LogsBasedOnTime: Failed to find any logs in the directory provided, need to do a deeper look to find some logs that we want.")
            $allFiles = Get-ChildItem $LogPath | Sort-Object LastWriteTime -Descending
            Remote-DisplayScriptDebug("Displaying all items in the directory: {0}" -f $LogPath)
            foreach($file in $allFiles)
            {
                Remote-DisplayScriptDebug("File Name: {0} Last Write Time: {1}" -f $file.Name, $file.LastWriteTime)
            }
            
            #Let's see if we have any files in this location while having directories 
            $directories = $allFiles | ?{$_.Mode -like "d*"}
            $filesInDirectory = $allFiles | ?{$_.Mode -notlike "d*"}

            if(($directories -ne $null) -and ($filesInDirectory -eq $null))
            {
                #This means we should be looking in the sub directories not the current directory so let's re-do that logic to try to find files in that timeframe. 
                foreach($dir in $directories)
                {
                    $newLogPath = $dir.FullName
                    $newCopyToThisLocation = "{0}\{1}" -f $CopyToThisLocation, $dir.Name
                    New-FolderCreate -Folder $newCopyToThisLocation
                    $Files = Get-ChildItem $newLogPath| Sort-Object LastWriteTime -Descending | ?{$_.LastWriteTime -ge $copyFromDate -and $_.Mode -notlike "d*"}
                    if($Files -eq $null)
                    {
                        No-FilesInLocation -CopyFromLocation $newLogPath -CopyToLocation $newCopyToThisLocation
                    }
                    else 
                    {
                        Remote-DisplayScriptDebug("Found {0} number of files at the location {1}" -f $Files.Count, $newLogPath)
                        $FilesFullPath = @()
                        $Files | %{$FilesFullPath += $_.VersionInfo.FileName}
                        Copy-BulkItems -CopyToLocation $newCopyToThisLocation -ItemsToCopyLocation $FilesFullPath
                        Zip-Folder -Folder $newCopyToThisLocation
                    }
                }
                Remote-DisplayScriptDebug("Function Exit: Copy-LogsBasedOnTime")
                return 
            }

            #If we get here, we want to find the latest file that isn't a directory.
            $Files = $allFiles | ?{$_.Mode -notlike "d*"} | Select-Object -First 1 

            #If we are still null, we want to let them know 
            If($Files -eq $null)
            {
                $SkipCopy = $true 
                No-FilesInLocation -CopyFromLocation $LogPath -CopyToLocation $CopyToThisLocation
            }
        }
        Remote-DisplayScriptDebug("Found {0} number of files at the location {1}" -f $Files.Count, $LogPath)
        #ResetFiles to full path 
        $FilesFullPath = @()
        $Files | %{$FilesFullPath += $_.VersionInfo.FileName}

        if(-not ($SkipCopy))
        {
            Copy-BulkItems -CopyToLocation $CopyToThisLocation -ItemsToCopyLocation $FilesFullPath
            Zip-Folder -Folder $CopyToThisLocation
        }
        Remote-DisplayScriptDebug("Function Exit: Copy-LogsBasedOnTime")
    }

    Function Copy-BulkItems {
    param(
    [string]$CopyToLocation,
    [Array]$ItemsToCopyLocation
    )
        #Create the folder 
        New-FolderCreate -Folder $CopyToLocation 
        foreach($item in $ItemsToCopyLocation)
        {
            Copy-Item -Path $item -Destination $CopyToLocation
        }
    }

    Function Remove-EventLogChar {
    param(
        [string]$location 
    )
        Get-ChildItem $location | Rename-Item -NewName {$_.Name -replace "%4","-"}
    }

    Function Test-IISMultiW3SVCDirectores {
        if($Script:this_ServerObject.Version -ge 15){$i = 3}
        else{$i = 2}

        if((Get-ChildItem $Script:IISLogDirectory | ?{$_.Name -like "W3SVC*"}).Count -ge $i ){return $true}
        return $false
    }

    Function Test-CommandExists {
    param(
    [string]$command
    )
        $oldAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"

        try {
            if(Get-Command $command){return $true}
        }
        catch {
            return $false
        }
        finally{
            $ErrorActionPreference = $oldAction
        }
    }

    Function Set-IISDirectoryInfo {
        Remote-DisplayScriptDebug("Function Enter: Set-IISDirectoryInfo")

        Function Get-IISDirectoryFromGetWebSite 
        {
            Remote-DisplayScriptDebug("Get-WebSite command exists")
            foreach($WebSite in $(Get-WebSite))
            {
                $logFile = "$($Website.logFile.directory)\W3SVC$($website.id)".replace("%SystemDrive%",$env:SystemDrive)
                $Script:IISLogDirectory += $logFile + ";"
                Remote-DisplayScriptDebug("Found Directory: {0}" -f $logFile)
            }
            #remove the last ; 
            $Script:IISLogDirectory = $Script:IISLogDirectory.Substring(0, $Script:IISLogDirectory.Length - 1)
            #$Script:IISLogDirectory = ((Get-WebConfigurationProperty "system.applicationHost/sites/siteDefaults" -Name logFile).directory).Replace("%SystemDrive%",$env:SystemDrive) 
            Remote-DisplayScriptDebug("Set IISLogDirectory: {0}" -f $Script:IISLogDirectory)
        }

        Function Get-IISDirectoryFromDefaultSettings 
        {
            $Script:IISLogDirectory = "C:\inetpub\logs\LogFiles\" #Default location for IIS Logs 
            Remote-DisplayScriptDebug("Get-WebSite command doesn't exists. Set IISLogDirectory to: {0}" -f $Script:IISLogDirectory)
        }

        if((Test-CommandExists -command "Get-WebSite"))
        {
            Get-IISDirectoryFromGetWebSite
        }
        else 
        {
            #May need to load the module 
            try 
            {
                Remote-DisplayScriptDebug("Going to attempt to load the WebAdministration Module")
                Import-Module WebAdministration
                Remote-DisplayScriptDebug("Successful loading the module")
                if((Test-CommandExists -command "Get-WebSite"))
                {
                    Get-IISDirectoryFromGetWebSite
                }
            }
            catch 
            {
                Get-IISDirectoryFromDefaultSettings
            }
            
        }
        #Test out the directories that we found. 
        foreach($directory in $Script:IISLogDirectory.Split(";"))
        {
            if(-not (Test-Path $directory))
            {
                Remote-DisplayScriptDebug("Failed to find a valid path for at least one of the IIS directories. Test path: {0}" -f $directory)
                Remote-DisplayScriptDebug("Function Exit: Set-IISDirectoryInfo - Failed")
                Write-Host("[{0}] : Failed to determine where the IIS Logs are located at. Unable to collect them." -f $Script:LocalServerName)
                return $false
            }
        }

        Remote-DisplayScriptDebug("Function Exit: Set-IISDirectoryInfo - Passed")
        return $true 
    }

    ####### Collect Logs Functions #####################
    Function Collect-ServerInfo {
        Remote-DisplayScriptDebug("Function Enter: Collect-ServerInfo")
        $copyTo = $Script:RootCopyToDirectory + "\General_Server_Info"
        New-FolderCreate -Folder $copyTo 

        #Get MSInfo from server 
        msinfo32.exe /nfo $copyTo\msinfo.nfo 
        Write-Warning("[{0}] : Waiting for msinfo32.exe process to end before moving on..." -f $Script:LocalServerName)
        while((Get-Process | ?{$_.ProcessName -eq "msinfo32"}).ProcessName -eq "msinfo32")
        {
            sleep 5;
        }

        #Running Processes #35 
        $runningProcesses = Get-Process
        Save-DataInfoToFile -dataIn $runningProcesses -SaveToLocation ("{0}\Running_Processes" -f $copyTo) -FormatList $false

        #Services Information #36
        $services = Get-Service 
        Save-DataInfoToFile -dataIn $services -SaveToLocation ("{0}\Services_Information" -f $copyTo) -FormatList $false

        #VSSAdmin Information #39
        $vssWriters = vssadmin list Writers
        $vssWriters > "$copyTo\VSS_Writers.txt"

        #Driver Information #34
        $drivers = Get-ChildItem ("{0}\System32\drivers" -f $env:SystemRoot) | Where-Object{$_.Name -like "*.sys"}
        Save-DataInfoToFile -dataIn $drivers -SaveToLocation ("{0}\System32_Drivers" -f $copyTo)

        Get-HotFix | Select Source, Description, HotFixID, InstalledBy, InstalledOn | Export-Clixml "$copyTo\HotFixInfo.xml"
        
        #TCPIP Networking Information #38
        ipconfig /all > "$copyTo\IPConfiguration.txt"

        netstat -anob > "$copyTo\Netstat_ANOB.txt"

        route print > "$copyTo\Network_Routes.txt"

        arp -a > "$copyTo\Network_ARP.txt"

        netstat -nato > "$copyTo\Netstat_NATO.txt"

        netstat -es > "$copyTo\Netstat_ES.txt" 

        #IPsec 
        netsh ipsec dynamic show all > "$copyTo\IPsec_netsh_dynamic.txt"

        netsh ipsec static show all > "$copyTo\IPsec_netsh_static.txt"

        #FLTMC
        fltmc > "$copyTo\FilterDrivers.txt"
        fltmc volumes > "$copyTo\FLTMC_Volumes.txt"
        fltmc instances > "$copyTo\FLTMC_Instances.txt"

        
        $hiveKey = Get-ChildItem HKLM:\SOFTWARE\Microsoft\Exchange\ -Recurse
        $hiveKey += Get-ChildItem HKLM:\SOFTWARE\Microsoft\ExchangeServer\ -Recurse
        $hiveKey | Export-Clixml "$copyTo\Exchange_Registry_Hive.xml"

        gpresult /R /Z > "$copyTo\GPResult.txt"
        gpresult /H "$copyTo\GPResult.html"

        #Storage Information 
        if(Test-CommandExists -command "Get-Volume")
        {
            $volume = Get-Volume
            $volume | fl * > "$copyTo\Volume.txt"
            $volume | Export-Clixml "$copyTo\Volume.xml"
        }
        else 
        {
            Remote-DisplayScriptDebug("Get-Volume isn't a valid command")    
        }

        if(Test-CommandExists -command "Get-Disk")
        {
            $disk = Get-Disk 
            $disk | fl * > "$copyTo\Disk.txt"
            $disk | Export-Clixml "$copyTo\Disk.xml"
        }
        else 
        {
            Remote-DisplayScriptDebug("Get-Disk isn't a valid command")    
        }


        Zip-Folder -Folder $copyTo
        Remote-DisplayScriptDebug("Function Exit: Collect-ServerInfo")
    }

    Function Get-HighAvailabilityLogs_V15 
    {
        $Logs = @() 
        $root = $script:LocalsysRoot

        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4AppLogMirror.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4BlockReplication.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4Debug.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4Monitoring.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4Network.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4Operational.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4Seeding.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4TruncationDebug.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-MailboxDatabaseFailureItems%4Operational.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-MailboxDatabaseFailureItems%4Debug.evtx"

        return $Logs 
    }

    Function Get-HighAvailabilityLogs_V14
    {
        $Logs = @()
        $root = $script:LocalsysRoot

        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4BlockReplication.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4Debug.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4Operational.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4SeedingDebug.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-HighAvailability%4TruncationDebug.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-MailboxDatabaseFailureItems%4Operational.evtx"
        $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-MailboxDatabaseFailureItems%4Debug.evtx"
        
        return $Logs 
    }

    Function Collect-HighAvailabilityLogs 
    {
        if($Script:this_ServerObject.Mailbox)
        {
            $copyTo = $Script:RootCopyToDirectory + "\High_Availability_logs"
            $Logs = @() 
            if($Script:this_ServerObject.DAGMember)
            {
                #Cluster log /g for some reason, we can't run this within invoke-command as we get a permission issue not sure why, as everything else works. 
                #going to run this cmdlet outside of invoke-command like all the other exchange cmdlets 
                $test = $script:LocalsysRoot + "\Cluster\Reports\Cluster.log"
                if(Test-Path -Path $test)
                {
                    $Logs += $test
                }
            }
            if($Script:this_ServerObject.Version -ge 15)
            {
                $Logs += Get-HighAvailabilityLogs_V15
            }
            elseif($Script:this_ServerObject.Version -eq 14)
            {
                $Logs += Get-HighAvailabilityLogs_V14 
            }
            else 
            {
                Write-Host("[{0}] : unknown server version: {1}" -f $Script:LocalServerName, $Script:this_ServerObject.Version) -ForegroundColor Red
                return 
            }
            Copy-BulkItems -CopyToLocation $copyTo -ItemsToCopyLocation $Logs 
            Remove-EventLogChar -location $copyTo
            Zip-Folder -Folder $copyTo
        }
        else 
        {
            Write-Host("[{0}] : Doesn't look like this server has the Mailbox Role Installed. Not going to collect the High Availability Logs" -f $Script:LocalServerName)
        }
    }

    Function Collect-AppSysLogs {

        $root = $script:LocalsysRoot
        $Logs = @()
        $Logs += $root + "\System32\Winevt\Logs\Application.evtx"
        $Logs += $root + "\System32\Winevt\Logs\system.evtx"
        $Logs += $root + "\System32\Winevt\Logs\MSExchange Management.evtx"

        $copyTo = $Script:RootCopyToDirectory + "\App_Sys_Logs"
        Copy-BulkItems -CopyToLocation $copyTo -ItemsToCopyLocation $Logs 

        Zip-Folder -Folder $copyTo

    }

    Function Collect-ManagedAvailabilityLogs {
    
            $root = $script:LocalsysRoot
            $Logs = @()
            $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-ActiveMonitoring%4ProbeResult.evtx" #Probe Results Logs 
            $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-ManagedAvailability%4RecoveryActionResults.evtx" #Recovery Action Results Logs 
            $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-ManagedAvailability%4RecoveryActionLogs.evtx" #Recovery Action Logs 
            $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-ActiveMonitoring%4ResponderDefinition.evtx" #Responder Definition Logs 
            $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-ActiveMonitoring%4ResponderResult.evtx" #Responder Results Logs 
            $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-ActiveMonitoring%4MonitorDefinition.evtx" #Monitor Definition Logs 
            $Logs += $root + "\System32\Winevt\Logs\Microsoft-Exchange-ManagedAvailability%4Monitoring.evtx" #Monitoring Logs 

            $copyTo = $Script:RootCopyToDirectory + "\MA_Logs"
            Copy-BulkItems -CopyToLocation $copyTo -ItemsToCopyLocation $Logs
            Remove-EventLogChar -location $copyTo 
            Zip-Folder -Folder $copyTo 

    }

    Function Enable-ZipAssembly {
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        try 
        {
            Add-Type -AssemblyName System.IO.Compression.Filesystem 
        }
        catch 
        {
            Write-Host("[{0}] : Failed to load .NET Compression assembly. Disable the ability to zip data" -f $Script:LocalServerName)
            $PassedInfo.Zip = $false
        }
        finally
        {
            $ErrorActionPreference = $oldErrorAction
        }

    }

    Function Get-ThisServerObject {

        foreach($srv in $PassedInfo.ServerObjects)
        {
            if($srv.ServerName -eq $Script:LocalServerName)
            {
                $Script:this_ServerObject = $srv 
            }
        }
        if($Script:this_ServerObject -eq $null -or $Script:this_ServerObject.ServerName -ne $Script:LocalServerName)
        {
            #Something went wrong.... 
            Write-Host("[{0}] : Something went wrong trying to find the correct Server Object for this server. Stopping this instance of Execution"-f $Script:LocalServerName)
            exit 
        }
    }

    #This is in two different location. Make changes to both. 
    Function Set-RootCopyDirectory{
        $date = Get-Date -Format yyyyMd
        $str = "{0}{1}" -f $PassedInfo.RootFilePath, $Script:LocalServerName
        return $str
    }

    Function Get-ExchangeInstallDirectory
    {
        $installDirectory = [string]::Empty
        if((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup'))
        {
            $installDirectory = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath	
        }
        elseif((Test-Path 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'))
        {
            $installDirectory = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath	
        }
        else 
        {
            Write-Host("[{0}] : Something went wrong trying to find the Exchange install path on this server. Stopping this instance of Execution" -f $env:COMPUTERNAME) 
            exit    
        }
        return $installDirectory 
    }

    Function Set-InstanceRunningVars
    {
        $Script:LocalServerName = $env:COMPUTERNAME
        $script:LocalsysRoot = $env:SystemRoot

        $Script:RootCopyToDirectory = Set-RootCopyDirectory
        #Set the local Server Object Information 
        Get-ThisServerObject 
                
        $Script:this_Exinstall = Get-ExchangeInstallDirectory
        #shortcut to Exbin directory (probably not really needed)
        $Script:this_ExBin = $Script:this_Exinstall + "Bin\"

    }

    Function Save-DataInfoToFile {
        param(
        $dataIn,
        $SaveToLocation,
        $FormatList = $true
        )
            
            $xmlOut = $SaveToLocation + ".xml"
            $txtOut = $SaveToLocation + ".txt"
            if($dataIn -ne $null)
            {
                $dataIn | Export-Clixml $xmlOut -Encoding UTF8
                if($FormatList)
                {
                    $dataIn | Format-List * | Out-File $txtOut
                }
                else 
                {
                    $dataIn | Format-Table -AutoSize | Out-File $txtOut
                }
            }
    }

    ###################################
    #                                 #
    #         Logman Functions        #
    #                                 #
    ###################################
    
    Function Start-Logman {
    param(
    [Parameter(Mandatory=$true)][string]$logmanName,
    [Parameter(Mandatory=$true)][string]$Server_Name
    )
        Write-Host("Starting Data Collection {0} on server {1}" -f $logmanName,$Server_Name)
        logman start -s $Server_Name $logmanName
    }
    
    Function Stop-Logman {
    param(
    [Parameter(Mandatory=$true)][string]$logmanName,
    [Parameter(Mandatory=$true)][string]$Server_Name
    )
        Write-Host("Stopping Data Collection {0} on server {1}" -f $logmanName,$Server_Name)
        logman stop -s $Server_Name $logmanName
    }
    
    
    Function Copy-LogmanData{
    param(
    [Parameter(Mandatory=$true)]$objLogman
    )
        switch ($objLogman.LogmanName)
        {
            "Exchange_Perfwiz" {$FolderName = "ExPerfWiz_Data"; break}
            "Exmon_Trace" {$FolderName = "ExmonTrace_Data"; break}
            default {$FolderName = "Logman_Data"; break}
        }
    
        $strDirectory = $objLogman.RootPath
        $copyTo = $Script:RootCopyToDirectory + "\" + $FolderName
        New-FolderCreate -Folder $copyTo
        if(Test-Path $strDirectory)
        {
            $wildExt = "*" + $objLogman.Ext
            $filterDate = $objLogman.StartDate
            $files = Get-ChildItem $strDirectory | ?{($_.Name -like $wildExt) -and ($_.CreationTime -ge $filterDate)}
            if($files -ne $null)
            {
                foreach($file in $files)
                {
                    Write-Host("[{0}] : Copying over file {1}..." -f $Script:LocalServerName, $file.VersionInfo.FileName)
                    copy $file.VersionInfo.FileName $copyTo
                }
                Zip-Folder -Folder $copyTo
            }
            else 
            {
                Write-Host ("[{0}] : Failed to find any files in the directory: '{1}' that was greater than or equal to this time: {2}" -f $Script:LocalServerName, $strDirectory, $filterDate) -ForegroundColor Yellow
                Write-Host ("[{0}] : Going to try to see if there are any files in this directory for you..." -f $Script:LocalServerName) -NoNewline
                $files = Get-ChildItem $strDirectory | ?{$_.Name -like $wildExt}
                if($files -ne $null)
                {
                    #only want to get lastest data 
                    $newestFilesTime = ($files | Sort CreationTime -Descending)[0].CreationTime.AddDays(-1)
                    $newestFiles = $files | ?{$_.CreationTime -ge $newestFilesTime}
                    foreach($file in $newestFiles)
                    {
                        Write-Host("[{0}] : Copying over file {1}..." -f $Script:LocalServerName, $file.VersionInfo.FileName)
                        copy $file.VersionInfo.FileName $copyTo
                    }
                    Zip-Folder -Folder $copyTo
                }
                else 
                {
                    Write-Warning ("[{0}] : Failed to find any files in the directory: '{1}'" -f $Script:LocalServerName, $strDirectory)      
                    $tempFile = $copyTo + "\NoFiles.txt"    
                    New-Item $tempFile -ItemType File -Value $strDirectory
                }
                
                
            }
        }
        else 
        {
            Write-Warning ("[{0}] : Doesn't look like this Directory is valid. {1}" -f $Script:LocalServerName, $strDirectory)
            $tempFile = $copyTo + "\NotValidDirectory.txt"
            New-Item $tempFile -ItemType File -Value $strDirectory
        }
    
    }

    
    Function Get-LogmanData {
    param(
    [Parameter(Mandatory=$true)][string]$logmanName,
    [Parameter(Mandatory=$true)][string]$Server_Name
    )
        $objLogman = Get-LogmanObject -logmanName $logmanName -Server_Name $Server_Name
        if($objLogman -ne $null)
        {
            switch ($objLogman.Status) 
            {
                "Running" {
                            Write-Host ("[{0}] : Looks like logman {1} is running...." -f $Script:LocalServerName, $logmanName)
                            Write-Host ("[{0}] : Going to stop {1} to prevent corruption...." -f $Script:LocalServerName, $logmanName)
                            Stop-Logman -logmanName $logmanName -Server_Name $Server_Name
                            Copy-LogmanData -objLogman $objLogman
                            Write-Host("[{0}] : Starting Logman {1} again for you...." -f $Script:LocalServerName, $logmanName)
                            Start-Logman -logmanName $logmanName -Server_Name $Server_Name
                            Write-Host ("[{0}] : Done starting Logman {1} for you" -f $Script:LocalServerName, $logmanName)
                            break;
                            }
                "Stopped" {
                            Write-Host ("[{0}] : Doesn't look like Logman {1} is running, so not going to stop it..." -f $Script:LocalServerName, $logmanName)
                            Copy-LogmanData -objLogman $objLogman
                            break;
                        }
                Default {
                            Write-Host ("[{0}] : Don't know what the status of Logman '{1}' is in" -f $Script:LocalServerName, $logmanName)
                            Write-Host ("[{0}] : This is the status: {1}" -f $Script:LocalServerName, $objLogman.Status)
                            Write-Host ("[{0}] : Going to try stop it just in case..." -f $Script:LocalServerName)
                            Stop-Logman -logmanName $logmanName -Server_Name $Server_Name
                            Copy-LogmanData -objLogman $objLogman
                            Write-Host ("[{0}] : Not going to start it back up again...." -f $Script:LocalServerName)
                            Write-Warning ("[{0}] : Please start this logman '{1}' if you need to...." -f $Script:LocalServerName, $logmanName)
                            break; 
                        }
            }
        }
        else 
        {
            Write-Host("[{0}] : Can't find {1} on {2} ..... Moving on." -f $Script:LocalServerName, $logmanName, $Server_Name)    
        }
    
    }
    
    Function Get-LogmanStatus {
    param(
    [Parameter(Mandatory=$true)]$rawLogmanData 
    )
        $status = "Status:"
        $stop = "Stopped"
        $run = "Running"
            
        if(-not($rawLogmanData[2].Contains($status)))
        {
            $i = 0
            while((-not($rawLogmanData[$i].Contains($status))) -and ($i -lt ($rawLogmanData.count - 1)))
            {
                $i++
            }
        }
        else {$i = 2}
        $strLine = $rawLogmanData[$i]
    
        if($strLine.Contains($stop)){$currentStatus = $stop}
        elseif($strLine.Contains($run)){$currentStatus = $run}
        else{$currentStatus = "unknown"}
        return $currentStatus
    }
    
    Function Get-LogmanRootPath {
    param(
    [Parameter(Mandatory=$true)]$rawLogmanData
    )
        $Root_Path = "Root Path:"
        if(-not($rawLogmanData[3].Contains($Root_Path)))
        {
            $i = 0
            while((-not($rawLogmanData[$i].Contains($Root_Path))) -and ($i -lt ($rawLogmanData.count - 1)))
            {
                $i++
            }
        }
        else {$i = 3}
    
        $strRootPath = $rawLogmanData[$i]
        $replace = $strRootPath.Replace("Root Path:", "")
        [int]$Index = $replace.IndexOf(":") - 1
        $strReturn = $replace.SubString($Index)
        return $strReturn
    }
    
    Function Get-LogmanStartDate {
    param(
    [Parameter(Mandatory=$true)]$rawLogmanData
    )
        $strStart_Date = "Start Date:"
        if(-not($rawLogmanData[11].Contains($strStart_Date)))
        {
            $i = 0
            while((-not($rawLogmanData[$i].Contains($strStart_Date))) -and ($i -lt ($rawLogmanData.count - 1)))
            {
                $i++
            }
            #Circular Log collection doesn't contain Start Date
            if(-not($rawLogmanData[$i].Contains($strStart_Date)))
            {
                $strReturn = (Get-Date).AddDays(-1).ToString()
                return $strReturn
            }
        }
        else {$i = 11}
        $strLine = $rawLogmanData[$i]
    
        [int]$index = $strLine.LastIndexOf(" ") + 1 
        $strReturn = $strLine.SubString($index)
        return $strReturn
    }
    
    Function Get-LogmanExt {
    param(
    [Parameter(Mandatory=$true)]$rawLogmanData 
    )
        $strLocation = "Output Location:"
        if(-not($rawLogmanData[15].Contains($strLocation)))
        {
            $i = 0
            while((-not($rawLogmanData[$i].Contains($strLocation))) -and ($i -lt ($rawLogmanData.Count - 1)))
            {
                $i++
            }
        }
        else{ $i = 15}
    
        $strLine = $rawLogmanData[$i]
        [int]$index = $strLine.LastIndexOf(".")
        if($index -ne -1)
        {
            $strExt = $strLine.SubString($index)
        }
        else {
            $strExt = $null
        }
        return $strExt
    }
    
    Function Get-LogmanObject {
    param(
    [Parameter(Mandatory=$true)][string]$logmanName,
    [Parameter(Mandatory=$true)][string]$Server_Name
    )
        $rawDataResults = logman -s $Server_Name $logmanName
        if($rawDataResults[$rawDataResults.Count - 1].Contains("Set was not found."))
        {
            return $null
        }
        else 
        {
            $objLogman = New-Object -TypeName psobject
            $objLogman | Add-Member -MemberType NoteProperty -Name LogmanName -Value $logmanName
            $objLogman | Add-Member -MemberType NoteProperty -Name Status -Value (Get-LogmanStatus -rawLogmanData $rawDataResults)
            $objLogman | Add-Member -MemberType NoteProperty -Name RootPath -Value (Get-LogmanRootPath -rawLogmanData $rawDataResults)
            $objLogman | Add-Member -MemberType NoteProperty -Name StartDate -Value (Get-LogmanStartDate -rawLogmanData $rawDataResults)
            $objLogman | Add-Member -MemberType NoteProperty -Name Ext -Value (Get-LogmanExt -rawLogmanData $rawDataResults)
            $objLogman | Add-Member -MemberType NoteProperty -Name RestartLogman -Value $false
            $objLogman | Add-Member -MemberType NoteProperty -Name ServerName -Value $Server_Name
            $objLogman | Add-Member -MemberType NoteProperty -Name RawData -Value $rawDataResults
            $objLogman | Add-Member -MemberType NoteProperty -Name SaveRootLocation -Value $FilePath
    
            return $objLogman
        }
    
    }

    Function  Collect-LogmanExperfwiz
    {
        Get-LogmanData -logmanName $PassedInfo.Experfwiz_LogmanName -Server_Name $Script:LocalServerName
    }

    Function Collect-LogmanExmon
    {
        Get-LogmanData -logmanName $PassedInfo.Exmon_LogmanName -Server_Name $Script:LocalServerName
    }

    Function Remote-Main {
        Remote-DisplayScriptDebug("Function Enter: Remote-Main")
        

        Set-InstanceRunningVars


        if($PassedInfo.Zip)
        {
            Enable-ZipAssembly
        }

        $cmdsToRun = @() 
        #############################################
        #                                           #
        #              Exchange 2013 +              #
        #                                           #
        #############################################
        $copyInfo = "-LogPath '{0}' -CopyToThisLocation '{1}'"
        if($Script:this_ServerObject.Version -ge 15)
        {
            Remote-DisplayScriptDebug("Server Version greater than 15")
            if($PassedInfo.EWSLogs)
            {
                if($Script:this_ServerObject.Mailbox)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\EWS"),($Script:RootCopyToDirectory + "\EWS_BE_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth){ $cmdsToRun += ("Copy-LogsBasedOnTime {0}" -f $info)}
                    else {$cmdsToRun += ("Copy-FullLogFullPathRecurse {0}" -f $info)}
                    
                }
                if($Script:this_ServerObject.CAS)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\HttpProxy\Ews"),($Script:RootCopyToDirectory + "\EWS_Proxy_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth){ $cmdsToRun += ("Copy-LogsBasedOnTime {0}" -f $info)}
                    else{$cmdsToRun += ("Copy-FullLogFullPathRecurse {0}" -f $info)}
                    
                }
            }

            if($PassedInfo.RPCLogs)
            {
                if($Script:this_ServerObject.Mailbox)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\RPC Client Access"), ($Script:RootCopyToDirectory + "\RCA_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth){ $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else{$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                    
                }
                if($Script:this_ServerObject.CAS)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\HttpProxy\RpcHttp"), ($Script:RootCopyToDirectory + "\RCA_Proxy_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth){ $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else{$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                    
                }

                $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\RpcHttp"), ($Script:RootCopyToDirectory + "\RPC_Http_Logs"))
                if($PassedInfo.CollectAllLogsBasedOnDaysWorth) {$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info }
                else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
            }

            if($Script:this_ServerObject.CAS -and $PassedInfo.EASLogs)
            {
                $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\HttpProxy\Eas"), ($Script:RootCopyToDirectory + "\EAS_Proxy_Logs"))
                if($PassedInfo.CollectAllLogsBasedOnDaysWorth) {$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
            }
            
            if($PassedInfo.AutoDLogs)
            {
                if($Script:this_ServerObject.Mailbox)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\Autodiscover"), ($Script:RootCopyToDirectory + "\AutoD_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth) {$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else { $cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                }
                if($Script:this_ServerObject.CAS)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\HttpProxy\Autodiscover"), ($Script:RootCopyToDirectory + "\AutoD_Proxy_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth) {$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else { $cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info }
                }
            }

            if($PassedInfo.OWALogs)
            {
                if($Script:this_ServerObject.Mailbox)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\OWA"), ($Script:RootCopyToDirectory + "\OWA_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth) {$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                }
                if($Script:this_ServerObject.CAS)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\HttpProxy\OwaCalendar"), ($Script:RootCopyToDirectory + "\OWA_Proxy_Calendar_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth) {$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else { $cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}

                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\HttpProxy\Owa"), ($Script:RootCopyToDirectory + "\OWA_Proxy_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth) {$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info }
                    else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                }
            }

            if($PassedInfo.ADDriverLogs)
            {
                $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\ADDriver"), ($Script:RootCopyToDirectory + "\AD_Driver_Logs"))
                if($PassedInfo.CollectAllLogsBasedOnDaysWorth){ $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
            }

            if($PassedInfo.MapiLogs)
            {
                if($Script:this_ServerObject.Mailbox -and $Script:this_ServerObject.Version -eq 15)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\MAPI Client Access"), ($Script:RootCopyToDirectory + "\MAPI_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth) {$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                }
                elseif($Script:this_ServerObject.Mailbox)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\MapiHttp\Mailbox"), ($Script:RootCopyToDirectory + "\MAPI_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth) {$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info} 
                }

                if($Script:this_ServerObject.CAS)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\HttpProxy\Mapi"), ($Script:RootCopyToDirectory + "\MAPI_Proxy_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth){$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                }
            }

            if($PassedInfo.ECPLogs)
            {
                if($Script:this_ServerObject.Mailbox)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\ECP"), ($Script:RootCopyToDirectory + "\ECP_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth){$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                }
                if($Script:this_ServerObject.CAS)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\HttpProxy\Ecp"), ($Script:RootCopyToDirectory + "\ECP_Proxy_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth){$cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else {$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                }
            }

            if($Script:this_ServerObject.Mailbox -and $PassedInfo.SearchLogs)
            {
                $info = ($copyInfo -f ($Script:this_ExBin + "Search\Ceres\Diagnostics\Logs"), ($Script:RootCopyToDirectory + "\Search_Diag_Logs"))
                $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info 
                $info = ($copyInfo -f ($Script:this_ExBin + "Search\Ceres\Diagnostics\ETLTraces"), ($Script:RootCopyToDirectory + "\Search_Diag_ETLs"))
                $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
            }
            
            if($PassedInfo.DailyPerformanceLogs)
            {
                #Daily Performace Logs are always by days worth 
                $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\Diagnostics\DailyPerformanceLogs"), ($Script:RootCopyToDirectory + "\Daily_Performance_Logs"))
                $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info 
            }

            if($PassedInfo.ManagedAvailability)
            {  
                $cmdsToRun += 'Collect-ManagedAvailabilityLogs'
            }
   
        }
        
        ############################################
        #                                          #
        #              Exchange 2010               #
        #                                          #
        ############################################
        if($Script:this_ServerObject.Version -eq 14)
        {
            if($Script:this_ServerObject.CAS)
            {
                if($PassedInfo.RPCLogs)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\RPC Client Access"), ($Script:RootCopyToDirectory + "\RCA_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth){ $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info}
                    else{$cmdsToRun += "Copy-FullLogFullPathRecurse {0}" -f $info}
                }
                if($PassedInfo.EWSLogs)
                {
                    $info = ($copyInfo -f ($Script:this_Exinstall + "Logging\EWS"),($Script:RootCopyToDirectory + "\EWS_BE_Logs"))
                    if($PassedInfo.CollectAllLogsBasedOnDaysWorth){ $cmdsToRun += ("Copy-LogsBasedOnTime {0}" -f $info)}
                    else {$cmdsToRun += ("Copy-FullLogFullPathRecurse {0}" -f $info)}
                }
            }
        }

        ############################################
        #                                          #
        #          All Exchange Versions           #
        #                                          #
        ############################################
        if($PassedInfo.AnyTransportSwitchesEnabled -and $Script:this_ServerObject.TransportInfoCollect)
        {
            if($PassedInfo.MessageTrackingLogs)
            {
                $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.HubLoggingInfo.MessageTrackingLogPath), ($Script:RootCopyToDirectory + "\Message_Tracking_Logs"))
                $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
            }

            if($PassedInfo.HubProtocolLogs)
            {
                $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.HubLoggingInfo.ReceiveProtocolLogPath), ($Script:RootCopyToDirectory + "\Hub_Receive_Protocol_Logs"))
                $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
                $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.HubLoggingInfo.SendProtocolLogPath), ($Script:RootCopyToDirectory + "\Hub_Send_Protocol_Logs"))
            }
            if($PassedInfo.HubConnectivityLogs)
            {
                $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.HubLoggingInfo.ConnectivityLogPath), ($Script:RootCopyToDirectory + "\Hub_Connectivity_Logs"))
                $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
            }
            if($PassedInfo.QueueInformationThisServer)
            {
                #current queue data 
                $data = $Script:this_ServerObject.TransportInfo.QueueData
                $create = $Script:RootCopyToDirectory + "\Queue_Data"
                New-FolderCreate $create 
                $saveLocation = $create + "\Current_Queue_Info"
                Save-DataInfoToFile -dataIn $data -SaveToLocation $saveLocation
                if($Script:this_ServerObject.Version -ge 15)
                {
                    $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.HubLoggingInfo.QueueLogPath), ($Script:RootCopyToDirectory + "\Queue_V15_Data"))
                    $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
                }
            }
            if($PassedInfo.ReceiveConnectors)
            {
                $data = $Script:this_ServerObject.TransportInfo.ReceiveConnectorData
                $create = $Script:RootCopyToDirectory + "\Connectors"
                New-FolderCreate $create
                $saveLocation = ($create + "\{0}_Receive_Connectors") -f $Script:LocalServerName
                Save-DataInfoToFile -dataIn $data -SaveToLocation $saveLocation
            }
            if($PassedInfo.TransportConfig)
            {
                if($Script:this_ServerObject.Version -ge 15)
                {
                    $items = @()
                    $items += $Script:this_ExBin + "\EdgeTransport.exe.config" 
                    $items += $Script:this_ExBin + "\MSExchangeFrontEndTransport.exe.config" 
                    $items += $Script:this_ExBin + "\MSExchangeDelivery.exe.config" 
                    $items += $Script:this_ExBin + "\MSExchangeSubmission.exe.config"

                }
                else 
                {
                    $items = @()
                    $items += $Script:this_ExBin + "\EdgeTransport.exe.config"
                }

                Copy-BulkItems -CopyToLocation ($Script:RootCopyToDirectory + "\Transport_Configuration") -ItemsToCopyLocation $items
            }
            #Exchange 2013+ only 
            if($Script:this_ServerObject.Version -ge 15)
            {
                if($PassedInfo.FrontEndConnectivityLogs)
                {
                    $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.FELoggingInfo.ConnectivityLogPath), ($Script:RootCopyToDirectory + "\FE_Connectivity_Logs"))
                    $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
                }
                if($PassedInfo.FrontEndProtocolLogs)
                {
                    $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.FELoggingInfo.ReceiveProtocolLogPath), ($Script:RootCopyToDirectory + "\FE_Receive_Protocol_Logs"))
                    $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
                    $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.FELoggingInfo.SendProtocolLogPath), ($Script:RootCopyToDirectory + "\FE_Send_Protocol_Logs"))
                    $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
                }
                if($PassedInfo.MailboxConnectivityLogs)
                {
                    $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.MBXLoggingInfo.ConnectivityLogPath + "\Delivery"), ($Script:RootCopyToDirectory + "\MBX_Delivery_Connectivity_Logs"))
                    $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
                    $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.MBXLoggingInfo.ConnectivityLogPath + "\Submission"), ($Script:RootCopyToDirectory + "\MBX_Submission_Connectivity_Logs"))
                    $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
                }
                if($PassedInfo.MailboxProtocolLogs)
                {
                    $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.MBXLoggingInfo.ReceiveProtocolLogPath), ($Script:RootCopyToDirectory + "\MBX_Receive_Protocol_Logs"))
                    $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
                    $info = ($copyInfo -f ($Script:this_ServerObject.TransportInfo.MBXLoggingInfo.SendProtocolLogPath), ($Script:RootCopyToDirectory + "\MBX_Send_Protocol_Logs"))
                    $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
                }
            }

        }

        if($PassedInfo.IISLogs -and (Set-IISDirectoryInfo))
        {
            foreach($directory in $Script:IISLogDirectory.Split(";"))
            {
                $copyTo = "{0}\IIS_{1}_Logs" -f $Script:RootCopyToDirectory, ($directory.Substring($directory.LastIndexOf("\") + 1))
                $info = ($copyInfo -f $directory, $copyTo) 
                $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info
            }

            $info = ($copyInfo -f ($script:LocalsysRoot +"\System32\LogFiles\HTTPERR"), ($Script:RootCopyToDirectory + "\HTTPERR_Logs"))
            $cmdsToRun += "Copy-LogsBasedOnTime {0}" -f $info 
        }

        if($PassedInfo.HighAvailabilityLogs)
        {
            $cmdsToRun += "Collect-HighAvailabilityLogs"
        }
        if($PassedInfo.ServerInfo)
        {
            $cmdsToRun += "Collect-ServerInfo"
        }

        if($PassedInfo.AppSysLogs)
        {
            $cmdsToRun += 'Collect-AppSysLogs'
        }

        if($PassedInfo.Experfwiz)
        {
            $cmdsToRun += "Collect-LogmanExperfwiz"
        }

        if($PassedInfo.Exmon)
        {
            $cmdsToRun += "Collect-LogmanExmon"
        }

        #Execute the cmds 
        foreach($cmd in $cmdsToRun)
        {
            Remote-DisplayScriptDebug("cmd: {0}" -f $cmd)
            Invoke-Expression $cmd
        }



        <#Dump out the data that only needs to be collected once, on the server that hosted the execution of the script
        if($Script:LocalServerName -eq ($PassedInfo.HostExeServerName))
        {
            Remote-DisplayScriptDebug("Writting only once data")
            if($PassedInfo.GetVdirs)
            {
                $target = $Script:RootCopyToDirectory + "\ConfigNC_msExchVirtualDirectory_All.CSV"
                $PassedInfo.VDirsInfo | Sort-Object -Property Server | Export-Csv $target -NoTypeInformation
            }

            if($PassedInfo.DAGInformation)
            {
                
                $data = $PassedInfo.DAGInfoData 
                $dagName = $data.DAGInfo.Name 
                $create =  $Script:RootCopyToDirectory + "\" + $dagName + "_DAG_MDB_Information"
                New-FolderCreate -Folder $create 
                $saveLocation = $create + "\{0}"
                                
                Save-DataInfoToFile -dataIn ($data.DAGInfo) -SaveToLocation ($saveLocation -f ($dagName +"_DAG_Info"))
                
                Save-DataInfoToFile -dataIn ($data.DAGNetworkInfo) -SaveToLocation ($saveLocation -f ($dagName + "DAG_Network_Info"))
                
                foreach($mdb in $data.AllMdbs)
                {
                    Save-DataInfoToFile -dataIn ($mdb.MDBInfo) -SaveToLocation ($saveLocation -f ($mdb.MDBName + "_DB_Info"))
                    Save-DataInfoToFile -dataIn ($mdb.MDBCopyStatus) -SaveToLocation ($saveLocation -f ($mdb.MDBName + "_DB_CopyStatus"))
                }

            }

            if($PassedInfo.SendConnectors)
            {
                $data = $PassedInfo.SendConnectorData
                $create = $Script:RootCopyToDirectory + "\Connectors"
                New-FolderCreate $create
                $saveLocation = $create + "\Send_Connectors"
                Save-DataInfoToFile -dataIn $data -SaveToLocation $saveLocation
            }
        }
        #>
        
        if((-not($PassedInfo.ExchangeServerInfo)) -and $Script:LocalServerName -ne ($PassedInfo.HostExeServerName))
        {
            #Zip it all up 
            Zip-Folder -Folder $Script:RootCopyToDirectory -ZipItAll $true
        }
        
    }

    $oldErrorAction = $ErrorActionPreference
    $ErrorActionPreference = "Stop"
    try 
    {
        if($PassedInfo.ByPass -ne $true)
        {
            Remote-Main
        }
        else 
        {
            Write-Host "Loading common functions"    
        }
        
    }
    catch 
    {
        Write-Host("[{0}] : An error occurred in Remote-Functions" -f $env:COMPUTERNAME) -ForegroundColor Red
        Write-Host("Error Exception: {0}" -f $Error[0].Exception) -ForegroundColor Red
        Write-Host("Error Stack: {0}" -f $Error[0].ScriptStackTrace) -ForegroundColor Red
    }
    finally
    {
        $ErrorActionPreference = $oldErrorAction
    }
}

Function Get-ExchangeObjectServerData{
param(
[Parameter(Mandatory=$true)][array]$Servers 
)
    Display-ScriptDebug("Enter Function: Get-ExchangeObjectServerData")
    $serverObjects = @()
    foreach($server in $Servers)
    {
        $obj = Get-ExchangeBasicServerObject -ServerName $server 

        if($obj.Hub)
        {
            if($obj.Version -ge 15)
            {
                $hubInfo = Get-TransportService $server 
            }
            else 
            {
                $hubInfo = Get-TransportServer $server
            }
            $obj | Add-Member -MemberType NoteProperty -Name TransportServerInfo -Value $hubInfo
        }
        if($obj.CAS)
        {
            if($obj.Version -ge 15)
            {
                $casInfo = Get-ClientAccessService $server
            }
            else 
            {
                $casInfo = Get-ClientAccessServer $server 
            }
            $obj | Add-Member -MemberType NoteProperty -Name CAServerInfo -Value $casInfo
        }
        if($obj.Mailbox)
        {
            $obj | Add-Member -MemberType NoteProperty -Name MailboxServerInfo -Value (Get-MailboxServer $server)
        }
        if($obj.Version -ge 15)
        {
            $obj | Add-Member -MemberType NoteProperty -Name HealthReport -Value (Get-HealthReport $server) 
            $obj | Add-Member -MemberType NoteProperty -Name ServerComponentState -Value (Get-ServerComponentState $server)
        }

        $serverObjects += $obj 
    }

    return $serverObjects 
}

#Template Master: https://github.com/dpaulson45/PublicPowerShellScripts/blob/master/Functions/Start-JobManager/Start-JobManager.ps1
Function Start-JobManager {
    [CmdletBinding()]
    param(
    [Parameter(Mandatory=$true)][array]$ServersWithArguments,
    [Parameter(Mandatory=$true)][scriptblock]$ScriptBlock,
    [Parameter(Mandatory=$false)][bool]$DisplayReceiveJob = $true,
    [Parameter(Mandatory=$false)][bool]$NeedReturnData = $false,
    [Parameter(Mandatory=$false)][scriptblock]$VerboseFunctionCaller,
    [Parameter(Mandatory=$false)][scriptblock]$HostFunctionCaller
    )
    
    #Function Version 1.0
    Function Write-VerboseWriter {
    param(
    [Parameter(Mandatory=$true)][string]$WriteString 
    )
        if($VerboseFunctionCaller -eq $null)
        {
            Write-Verbose $WriteString
        }
        else 
        {
            &$VerboseFunctionCaller $WriteString
        }
    }
    
    Function Write-HostWriter {
    param(
    [Parameter(Mandatory=$true)][string]$WriteString 
    )
        if($HostFunctionCaller -eq $null)
        {
            Write-Host $WriteString
        }
        else
        {
            &$HostFunctionCaller $WriteString    
        }
    }
    
    $passedVerboseFunctionCaller = $false
    $passedHostFunctionCaller = $false
    if($VerboseFunctionCaller -ne $null){$passedVerboseFunctionCaller = $true}
    if($HostFunctionCaller -ne $null){$passedHostFunctionCaller = $true}
    
    Function Start-Jobs {
        Write-VerboseWriter("Calling Start-Jobs")
        foreach($serverObject in $ServersWithArguments)
        {
            $server = $serverObject.ServerName
            $argumentList = $serverObject.ArgumentList
            Write-VerboseWriter("Starting job on server {0}" -f $server)
            Invoke-Command -ComputerName $server -ScriptBlock $ScriptBlock -ArgumentList $argumentList -AsJob -JobName $server | Out-Null
        }
    }
    
    Function Confirm-JobsPending {
        $jobs = Get-Job
        if($jobs -ne $null)
        {
            return $true 
        }
        return $false
    }
    
    Function Wait-JobsCompleted {
        Write-VerboseWriter("Calling Wait-JobsCompleted")
        [System.Diagnostics.Stopwatch]$timer = [System.Diagnostics.Stopwatch]::StartNew()
        while(Confirm-JobsPending)
        {
            $completedJobs = Get-Job | Where-Object {$_.State -ne "Running"}
            if($completedJobs -eq $null)
            {
                Start-Sleep 1 
                continue 
            }
    
            $returnData = @{}
            foreach($job in $completedJobs)
            {
                Write-VerboseWriter("Job {0} received. State: {1} | HasMoreData: {2}" -f $job.Name, $job.State,$job.HasMoreData)
                if($NeedReturnData -eq $false -and $DisplayReceiveJob -eq $false -and $job.HasMoreData -eq $true)
                {
                    Write-VerboseWriter("This job has data and you provided you didn't want to return it or display it.")
                }
                $receiveJob = Receive-Job $job 
                Remove-Job $job
                if($DisplayReceiveJob)
                {
                    $receiveJob
                }
                if($NeedReturnData)
                {
                    $returnData.Add($job.Name, $receiveJob)
                }
            }
        }
        $timer.Stop()
        Write-VerboseWriter("Waiting for jobs to complete took {0} seconds" -f $timer.Elapsed.TotalSeconds)
        if($NeedReturnData)
        {
            return $returnData
        }
        return $null 
    }
    
    [System.Diagnostics.Stopwatch]$timerMain = [System.Diagnostics.Stopwatch]::StartNew()
    Write-VerboseWriter("Calling Start-JobManager")
    Write-VerboseWriter("Passed: [bool]DisplayReceiveJob: {0} | [bool]NeedReturnData:{1} | [scriptblock]VerboseFunctionCaller: {2} | [scriptblock]HostFunctionCaller: {3}" -f $DisplayReceiveJob,
    $NeedReturnData,
    $passedVerboseFunctionCaller,
    $passedHostFunctionCaller)
    
    if($NeedReturnData -and $DisplayReceiveJob)
    {
        Write-VerboseWriter("Unable to display the job as well as return data. Setting DisplayReceiveJob to false")
        $DisplayReceiveJob = $false
    }
    
    Start-Jobs
    $data = Wait-JobsCompleted
    $timerMain.Stop()
    Write-VerboseWriter("Exiting: Start-JobManager | Time in Start-JobManager: {0} seconds" -f $timerMain.Elapsed.TotalSeconds)
    if($NeedReturnData)
    {
        return $data
    }
    return $null
}

Function Write-ExchangeDataOnMachines {

    Function Write-ExchangeData {
        param(
        [Parameter(Mandatory=$true)][object]$PassedInfo
        )

                $server = $PassedInfo.ServerObject 
                $location = $PassedInfo.Location 
                Function Write-Data{
                param(
                [Parameter(Mandatory=$true)][object]$DataIn, 
                [Parameter(Mandatory=$true)][string]$FilePathNoEXT
                )
                    $DataIn | Format-List * > "$FilePathNoEXT.txt"
                    $DataIn | Export-Clixml "$FilePathNoEXT.xml"
                }
                $exchBin = "{0}\Bin" -f $PassedInfo.InstallDirectory
                $configFiles = Get-ChildItem $exchBin | Where-Object{$_.Name -like "*.config"}
                $copyTo = "{0}\Config" -f $location 
                $configFiles | ForEach-Object{ Copy-Item $_.VersionInfo.FileName $copyTo}

                Write-Data -DataIn $server.ExchangeServer -FilePathNoEXT "$location\ExchangeServer"

                Get-Command exsetup | ForEach-Object{$_.FileVersionInfo} > "$location\GCM.txt"

                if($server.Hub)
                {
                    Write-Data -DataIn $server.TransportServerInfo -FilePathNoEXT "$location\TransportServer"
                }
                if($server.CAS)
                {
                    Write-Data -DataIn $server.CAServerInfo -FilePathNoEXT "$location\ClientAccessServer"
                }
                if($server.Mailbox)
                {
                    Write-Data -DataIn $server.MailboxServerInfo -FilePathNoEXT "$location\MailboxServer"
                }
                if($server.Version -ge 15)
                {
                    Write-Data -DataIn $server.HealthReport -FilePathNoEXT "$location\HealthReport"
                    Write-Data -DataIn $server.ServerComponentState -FilePathNoEXT "$location\ServerComponentState"
                }
        }


    $exchangeServerData = Get-ExchangeObjectServerData -Servers $Script:ValidServers 
    #if single server or Exchange 2010 where invoke-command doesn't work 
    if($Script:ValidServers.count -gt 1)
    {
        #Need to have install directory run through the loop first as it could be different on each server
        $serversObjectListInstall = @() 
        foreach($server in $exchangeServerData)
        {
            $serverObject = New-Object PSCustomObject 
            $serverObject | Add-Member -MemberType NoteProperty -Name ServerName -Value $server.ServerName
            $serverObject | Add-Member -MemberType NoteProperty -Name ArgumentList -Value ([string]::Empty)
            $serversObjectListInstall += $serverObject
        }
        $serverInstallDirectories = Start-JobManager -ServersWithArguments $serversObjectListInstall -ScriptBlock ${Function:Get-ExchangeInstallDirectory} -VerboseFunctionCaller ${Function:Display-ScriptDebug} -NeedReturnData $true 
    
    
        $serverListCreateDirectories = @() 
        $serverListDumpData = @() 
        $serverListZipData = @() 
    
        foreach($server in $exchangeServerData)
        {
            $location = "{0}{1}\Exchange_Server_Data" -f $Script:RootFilePath, $server.ServerName 

            #Create Directory 
            $serverCreateDirectory = New-Object PSCustomObject 
            $serverCreateDirectory | Add-Member -MemberType NoteProperty -Name ServerName -Value $server.ServerName
            $serverCreateDirectory | Add-Member -MemberType NoteProperty -Name ArgumentList -Value ("{0}{1}\Exchange_Server_Data\Config" -f $Script:RootFilePath, $server.ServerName)
            $serverListCreateDirectories += $serverCreateDirectory

            #Write Data 
            $argumentList = New-Object PSCustomObject 
            $argumentList | Add-Member -MemberType NoteProperty -Name ServerObject -Value $server
            $argumentList | Add-Member -MemberType NoteProperty -Name Location -Value $location
            $argumentList | Add-Member -MemberType NoteProperty -Name InstallDirectory -Value $serverInstallDirectories[$server.ServerName]
            $serverDumpData = New-Object PSCustomObject 
            $serverDumpData | Add-Member -MemberType NoteProperty -Name ServerName -Value $server.ServerName 
            $serverDumpData | Add-Member -MemberType NoteProperty -Name ArgumentList -Value $argumentList
            $serverListDumpData += $serverDumpData

            #Zip data if not local cause we might have more stuff to run 
            if($server.ServerName -ne $env:COMPUTERNAME)
            {
                $folder = "{0}{1}" -f $Script:RootFilePath, $server.ServerName
                $parameters = $folder, $true, $false
                $serverZipData = New-Object PSCustomObject 
                $serverZipData | Add-Member -MemberType NoteProperty -Name ServerName -Value $server.ServerName
                $serverZipData | Add-Member -MemberType NoteProperty -Name ArgumentList -Value $parameters  
                $serverListZipData += $serverZipData 
            }
        }


        Display-ScriptDebug("Calling job for folder creation")
        Start-JobManager -ServersWithArguments $serverListCreateDirectories -ScriptBlock ${Function:New-FolderCreate} -VerboseFunctionCaller ${Function:Display-ScriptDebug}
        Display-ScriptDebug("Calling job for Exchange Data Write")
        Start-JobManager -ServersWithArguments $serverListDumpData -ScriptBlock ${Function:Write-ExchangeData} -VerboseFunctionCaller ${Function:Display-ScriptDebug} -DisplayReceiveJob $false 
        Display-ScriptDebug("Calling job for Zipping the data")
        Start-JobManager -ServersWithArguments $serverListZipData -ScriptBlock ${Function:Compress-Folder} -VerboseFunctionCaller ${Function:Display-ScriptDebug} 

    }
    else 
    {
        if($exinstall -eq $null)
        {
            $exinstall = Get-ExchangeInstallDirectory
        }
        $location = "{0}{1}\Exchange_Server_Data" -f $Script:RootFilePath, $exchangeServerData.ServerName
        New-FolderCreate -Folder ("{0}\Config" -f $location)
        $passInfo = New-Object PSCustomObject 
        $passInfo | Add-Member -MemberType NoteProperty -Name ServerObject -Value $exchangeServerData 
        $passInfo | Add-Member -MemberType NoteProperty -Name Location -Value $location
        $passInfo | Add-Member -MemberType NoteProperty -Name InstallDirectory -Value $exinstall 
        Display-ScriptDebug("Writing out the Exchange data")
        Write-ExchangeData -PassedInfo $passInfo 
        $folder = "{0}{1}" -f $Script:RootFilePath, $exchangeServerData.ServerName
                
    }
}
Function Write-DataOnlyOnceOnLocalMachine {
    Display-ScriptDebug("Enter Function: Write-DataOnlyOnceOnLocalMachine")
    Display-ScriptDebug("Writting only once data")

    #This is in two different location. Make changes to both. 
    Function Set-LocalRootCopyDirectory{
        $date = Get-Date -Format yyyyMd
        $str = "{0}\{1}\{2}" -f $FilePath, $date, $env:COMPUTERNAME
        return $str
    }

    #This is in two different location. Make changes to both. 
    Function New-LocalFolderCreate {
        param(
        [string]$Folder
        )
            if(-not (Test-Path -Path $Folder))
            {
                Write-Host("[{0}] : Creating Directory {1}" -f $env:COMPUTERNAME, $Folder)
                [System.IO.Directory]::CreateDirectory($Folder) | Out-Null
            }
            else 
            {
                Write-Host("[{0}] : Directory {1} is already created!" -f $env:COMPUTERNAME, $Folder)
            }
    
    }

     #This is in two different location. Make changes to both. 
    Function Save-LocalDataInfoToFile {
        param(
        $dataIn,
        $SaveToLocation 
        )
            
            $xmlOut = $SaveToLocation + ".xml"
            $txtOut = $SaveToLocation + ".txt"
            if($data -ne $null)
            {
                $dataIn | Export-Clixml $xmlOut -Encoding UTF8
                $dataIn | fl * | Out-File $txtOut
            }
    }

    Function Enable-LocalZipAssembly {
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        try 
        {
            $Script:LocalZip = $true
            Add-Type -AssemblyName System.IO.Compression.Filesystem 
        }
        catch 
        {
            Write-Host("[{0}] : Failed to load .NET Compression assembly. Disable the ability to zip data" -f $Script:LocalServerName)
            $Script:LocalZip = $false
        }
        finally
        {
            $ErrorActionPreference = $oldErrorAction
        }

    }
    
    Function Zip-LocalFolder {
        param(
        [string]$Folder,
        [bool]$ZipItAll
        )
    
            if($Script:LocalZip)
            {
                if(-not($ZipItAll))
                {
                    #Zip location 
                    $zipFolder = $Folder + ".zip"
                    if(Test-Path -Path $zipFolder)
                    {
                        #Folder exist for some reason 
                        [int]$i = 1
                        do{
                            $zipFolder = $Folder + "-" + $i + ".zip"
                            $i++
                        }while(Test-Path -Path $zipFolder)
                    }
                }
                else 
                {
                    $zipFolder = "{0}-{1}.zip" -f $Folder, (Get-Date -Format Md)
                    if(Test-Path -Path $zipFolder)
                    {
                        [int]$i = 1
                        $date = Get-Date -Format Md
                        do{
                            $zipFolder = "{0}-{1}-{2}.zip" -f $Folder, $date, $i
                            $i++
                        }while(Test-Path -Path $zipFolder)
                    }
    
                }
    
                if(-not($ZipItAll)){Write-Host("[{0}] : Zipping up the folder {1}" -f $env:COMPUTERNAME, $Folder)}
                else{Write-Host("[{0}] : Zipping up all the data for the server...." -f $env:COMPUTERNAME)}
                [System.IO.Compression.ZipFile]::CreateFromDirectory($Folder, $zipFolder)
    
                if((Test-Path -Path $zipFolder))
                {
                    Remove-Item $Folder -Force -Recurse
                }
            }
        }

    Enable-LocalZipAssembly
    $RootCopyToDirectory = Set-LocalRootCopyDirectory

    if($GetVdirs)
    {
        $target = $RootCopyToDirectory  + "\ConfigNC_msExchVirtualDirectory_All.CSV"
        $data = (Get-VdirsLDAP)
        $data | Sort-Object -Property Server | Export-Csv $target -NoTypeInformation
    }

    if($OrganizationConfig)
    {
        $target = $RootCopyToDirectory + "\OrganizationConfig"
        $data = Get-OrganizationConfig
        Save-LocalDataInfoToFile -dataIn $data -SaveToLocation $target
    }

    if($DAGInformation)
    {
        $data = Get-DAGInformation
        $dagName = $data.DAGInfo.Name 
        $create =  $RootCopyToDirectory  + "\" + $dagName + "_DAG_MDB_Information"
        New-FolderCreate -Folder $create 
        $saveLocation = $create + "\{0}"
                        
        Save-LocalDataInfoToFile -dataIn ($data.DAGInfo) -SaveToLocation ($saveLocation -f ($dagName +"_DAG_Info"))
        
        Save-LocalDataInfoToFile -dataIn ($data.DAGNetworkInfo) -SaveToLocation ($saveLocation -f ($dagName + "DAG_Network_Info"))
        
        foreach($mdb in $data.AllMdbs)
        {
            Save-LocalDataInfoToFile -dataIn ($mdb.MDBInfo) -SaveToLocation ($saveLocation -f ($mdb.MDBName + "_DB_Info"))
            Save-LocalDataInfoToFile -dataIn ($mdb.MDBCopyStatus) -SaveToLocation ($saveLocation -f ($mdb.MDBName + "_DB_CopyStatus"))
        }

        Zip-LocalFolder -Folder $create
    }

    if($SendConnectors)
    {
        $data = Get-SendConnector 
        $create = $RootCopyToDirectory + "\Connectors"
        New-FolderCreate $create
        $saveLocation = $create + "\Send_Connectors"
        Save-LocalDataInfoToFile -dataIn $data -SaveToLocation $saveLocation
    }

    Zip-Folder -Folder $RootCopyToDirectory -ZipItAll $true
    Display-ScriptDebug("Exiting Function: Write-DataOnlyOnceOnLocalMachine")
}


##################Main###################
Function Main {

    Display-Disclaimer
    Test-PossibleCommonScenarios
    Test-NoSwitchesProvided
    if(-not (Is-Admin))
    {
        Write-Warning "Hey! The script needs to be executed in elevated mode. Start the Exchange Mangement Shell as an Administrator."
        exit 
    }
    Load-ExShell

    <#
    Added the ability to call functions from within a bundled function so i don't have to duplicate work. 
    Loading the functions into memory by using the '.' allows me to do this, 
    providing that the calling of that function doesn't do anything of value when doing this. 
    #>
    $obj = New-Object PSCustomObject 
    $obj | Add-Member -MemberType NoteProperty -Name ByPass -Value $true 
    . Remote-Functions -PassedInfo $obj
    $Script:RootFilePath = "{0}\{1}\" -f $FilePath, (Get-Date -Format yyyyMd)

    if($Servers -ne $null)
    {
        
        #possible to return null or only a single server back (localhost)
        $Script:ValidServers = Test-RemoteExecutionOfServers -Server_List $Servers
        if($Script:ValidServers -ne $null)
        {
            $Script:ValidServers = Test-DiskSpace -Servers $Script:ValidServers -Path $FilePath -CheckSize 15
            $remote_ScriptingBlock = ${Function:Remote-Functions}
            Verify-LocalServerIsUsed $Script:ValidServers

            $argumentList = Get-ArgumentList -Servers $Script:ValidServers
            #I can do a try catch here, but i also need to do a try catch in the remote so i don't end up failing here and assume the wrong failure location
            try 
            {
                Invoke-Command -ComputerName $Script:ValidServers -ScriptBlock $remote_ScriptingBlock -ArgumentList $argumentList -ErrorAction Stop
            }
            catch 
            {
                Write-Error "An error has occurred attempting to call Invoke-Command to do a remote collect all at once. Please notify dpaul@microsoft.com of this issue. Stopping the script."
                exit
            }
            
            #Write out Exchange Data 
            if($ExchangeServerInfo)
            {
                [System.Diagnostics.Stopwatch]$timer = [System.Diagnostics.Stopwatch]::StartNew()
                Write-ExchangeDataOnMachines
                $timer.Stop()
                Display-ScriptDebug("Write-ExchangeDataOnMachines total time took {0} seconds" -f $timer.Elapsed.TotalSeconds)
            }

            Write-DataOnlyOnceOnLocalMachine
            $LogPaths = Get-RemoteLogLocation -Servers $Script:ValidServers -RootPath $Script:RootFilePath
            if((-not($SkipEndCopyOver)) -and (Test-DiskSpaceForCopyOver -LogPathObject $LogPaths -RootPath $Script:RootFilePath))
            {
                Write-Host("")
                Write-Host("Copying over the data may take some time depending on the network")
                foreach($svr in $LogPaths)
                {
                    #Don't want to do the local host
                    if($svr.ServerName -ne $env:COMPUTERNAME)
                    {
                        $remoteCopyLocation = "\\{0}\{1}" -f $svr.ServerName, ($svr.ZipFolder.Replace(":","$"))
                        Write-Host("[{0}] : Copying File {1}...." -f $svr.ServerName, $remoteCopyLocation) 
                        Copy-Item -Path $remoteCopyLocation -Destination $Script:RootFilePath
                        Write-Host("[{0}] : Done copying file" -f $svr.ServerName)
                    }
                    
                }

            }
            else 
            {
                Write-Host("")
                Write-Host("Please collect the following files from these servers and upload them: ")
                foreach($svr in $LogPaths)
                {
                    Write-Host("Server: {0} Path: {1}" -f $svr.ServerName, $svr.ZipFolder) 
                }
            }
        }
        else 
        {
            #We have failed to do invoke-command on all the servers.... so we are going to do the same logic locally
            Write-Host("Failed to do remote collection for all the servers in the list...") -ForegroundColor Yellow
            do{
                $read = Read-Host("Do you want me to collect from the local server only? 'y' or 'n'")
            }while($read -ne "y" -and $read -ne "n")
            if($read -eq "y")
            {
                Remote-Functions -PassedInfo (Get-ArgumentList -Servers $env:COMPUTERNAME)
                $Script:ValidServers = @($env:COMPUTERNAME)
                Write-ExchangeDataOnMachines
                Write-DataOnlyOnceOnLocalMachine
            }
            
        }
    }

    else 
    {
        Write-Host("Note: Remote Collection is now possible for Windows Server 2012 and greater on the remote machine. Just use the -Servers paramater with a list of Exchange Server names") -ForegroundColor Yellow
        Write-Host("Going to collect the data locally")
        Remote-Functions -PassedInfo (Get-ArgumentList -Servers $env:COMPUTERNAME)
        $Script:ValidServers = @($env:COMPUTERNAME)
        Write-ExchangeDataOnMachines
        Write-DataOnlyOnceOnLocalMachine
    }

    Display-FeedBack
        
}

Main 