ConvertFrom-StringData @'
###PSLOC
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
EXOV2ModuleNotInstalled = This script uses modern authenticaion to connect to Exchange Online and requires EXO V2 module to be installed. Please follow the instructions at https://docs.microsoft.com/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-the-exo-v2-module to install EXO V2 module.
###PSLOC
'@