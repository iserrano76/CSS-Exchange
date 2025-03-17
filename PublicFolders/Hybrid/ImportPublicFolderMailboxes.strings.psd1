ConvertFrom-StringData @'
###PSLOC
SyncingPublicFolderMailbox = Syncing public folder mailbox '{0}'.
CreatingMailUser = Creating mailuser object '{0}'.
MailUserExists = Mailuser object '{0}' already exists for this public folder mailbox.
ConfiguringMailUser = Adding '{0}' to RemotePublicFolderMailboxes.
DoneSyncingPublicFolderMailbox = Done syncing public folder mailbox '{0}'
NoHierarchyPublicFolderMailbox = There aren't any public folder mailboxes, serving hierarchy, to import.
DeletingMailUsersInfo = Deleting mailusers, if any, that don't have corresponding public folder mailboxes in the cloud, serving hierarchy.
RemovingMailUsers = Removing '{0}' from RemotePublicFolderMailboxes.
DeleteMailUser = Deleting mailuser object '{0}'.
IncorrectCredentials = Please provide correct credentials to establish remote session.
StartedPublicFolderMailboxImport = Started import of public folder mailboxes.
CompletedPublicFolderMailboxImport = Completed import of public folder mailboxes.
EXOV2ModuleNotInstalled = This script uses modern authenticaion to connect to Exchange Online and requires EXO V2 module to be installed. Please follow the instructions at https://docs.microsoft.com/powershell/exchange/exchange-online-powershell-v2?view=exchange-ps#install-the-exo-v2-module to install EXO V2 module.
###PSLOC
'@