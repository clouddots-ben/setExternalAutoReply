# Create Log
$date = get-date -Format yyyyMMddHHmmss
$logPath = join-path -Path $([System.IO.Path]::GetTempPath()) -childpath "$($date)OOOLog.csv"
start-transcript -path $logPath -append
# ExchangeOnline Module Check
write-information "Checking for installation of Module ExchangeOnlineManagement..."
if (!(get-module -listavailable -name ExchangeOnlineManagement)) {
    Install-Module -name ExchangeOnlineManagement -Force -ErrorAction Stop
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    write-information "ExchangeOnlineManagement Module was Installed and Imported"
}
elseif (!(get-module ExchangeOnlineManagement)) {
    Import-Module -name ExchangeOnlineManagement
    write-information "ExchangeOnlineManagement Module Imported"
}
# Import configuration File
$configPath = join-path -path $PSScriptRoot -childpath config.json
$config = get-content -path $configPath | convertfrom-json
# Authentication
Connect-ExchangeOnline -UserPrincipalName $config.EXOAdminUsername
# Import Users from Declared Path
$userImportPath = join-path -path $PSScriptRoot -childpath userImport.csv
$oooCSV = Import-Csv $userImportPath -ErrorAction Stop
# Get Current Mailbox OOO configuration
$mailboxOOOCurrentConfig = $oooCSV | foreach-object { 
    Get-MailboxAutoReplyConfiguration -Identity $_.user
}
# export existingforward array as CSV to the temp directory
if ($mailboxOOOCurrentConfig.count -gt 0) {
    $mailboxOOOCurrentConfigPath = join-path -Path $([System.IO.Path]::GetTempPath()) -childpath "$($date)mailboxOOOCurrentConfig.csv"
    $mailboxOOOCurrentConfig | Export-Csv -path $mailboxOOOCurrentConfigPath
    write-information -messagedata "List of accounts with existing forwards has been created at $mailboxOOOCurrentConfigPath" -informationaction Continue
}
# Set Enable OOO and Set external Message to message in $externalMessage
$externalMessagePath = join-path -path $PSScriptRoot -childpath externalMessage.txt
$externalMessage = get-content -Path $externalMessagePath
$count = 0
$oooCSV | foreach-object { 
    $count++
    write-information -MessageData "$count $($_.user)" -informationaction Continue
    Set-MailboxAutoReplyConfiguration -Identity $_.user -autoreplystate Enabled -externalmessage $ExecutionContext.InvokeCommand.ExpandString($externalMessage)
}
Disconnect-ExchangeOnline -Confirm:$false
Stop-Transcript