# Powershell script for Zabbix agents.

# Version 2.1 - for Zabbix agent 5x

## This script will check for pending Windows Updates, report them to Zabbix, and optionally install the updates.

### If you do not wish the script to install updates, look for the comment in the script that tells you how to disable that function.

#### Check https://github.com/SpookOz/zabbix-winupdates for the latest version of this script

# ------------------------------------------------------------------------- #
# Variables
# ------------------------------------------------------------------------- #

# Change $reportpath to wherever you want your update reports to go.

#$reportpath = "C:\IT\WinUpdates"

# Change $ZabbixInstallPath to wherever your Zabbix Agent is installed

$ZabbixInstallPath = "$Env:Programfiles\Zabbix Agent 2"
$ZabbixConfFile = "$Env:Programfiles\Zabbix Agent 2"

# Do not change the following variables unless you know what you are doing

$htReplace = New-Object hashtable
foreach ($letter in (Write-Output ä ae ö oe ü ue Ä Ae Ö Oe Ü Ue ß ss)) {
    $foreach.MoveNext() | Out-Null
    $htReplace.$letter = $foreach.Current
}
$pattern = "[$(-join $htReplace.Keys)]"
$returnStateOK = 0
$returnStateWarning = 1
$returnStateCritical = 2
$returnStateUnknown = 3
$returnStateOptionalUpdates = $returnStateWarning
$Sender = "$ZabbixInstallPath\zabbix_sender.exe"
$Senderarg1 = '-vv'
$Senderarg2 = '-c'
$Senderarg3 = "$ZabbixConfFile\zabbix_agent2.conf"
$Senderarg4 = '-i'
$SenderargUpdateReboot = '\updatereboot.txt'
$Senderarglastupdated = '\lastupdated.txt'
$Senderargcountcritical = '\countcritical.txt'
$SenderargcountOptional = '\countOptional.txt'
$SenderargcountHidden = '\countHidden.txt'
$Countcriticalnum = '\countcriticalnum.txt'
$Senderarg5 = '-k'
$Senderargupdating = 'Winupdates.Updating'
$Senderarg6 = '-o'
$Senderarg7 = '0'
$Senderarg8 = '1'


$test= get-content "C:\Program Files\Zabbix Agent 2\zabbix_agent2.conf" | select -first 1 -skip 133
$ok=$test.Substring(9)






#If(!(test-path $reportpath))
#{
#     New-Item -ItemType Directory -Force -Path $reportpath
#}

# ------------------------------------------------------------------------- #
# This part gets the date Windows Updates were last applied and writes it to temp file
# ------------------------------------------------------------------------- #

$windowsUpdateObject = New-Object -ComObject Microsoft.Update.AutoUpdate
Write-Output "- Winupdates.LastUpdated $($windowsUpdateObject.Results.LastInstallationSuccessDate)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderarglastupdated

# ------------------------------------------------------------------------- #
# This part get the reboot status and writes to test file
# ------------------------------------------------------------------------- #

if (Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"){ 
	Write-Output "- Winupdates.Reboot 1" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargUpdateReboot
    Write-Host "`t There is a reboot pending" -ForeGroundColor "Red"
}else {
	Write-Output "- Winupdates.Reboot 0" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargUpdateReboot
    Write-Host "`t No reboot pending" -ForeGroundColor "Green"
		}
# ------------------------------------------------------------------------- #		
# This part checks available Windows updates
# ------------------------------------------------------------------------- #

$updateSession = new-object -com "Microsoft.Update.Session"
$updates=$updateSession.CreateupdateSearcher().Search(("IsInstalled=0 and Type='Software'")).Updates

$criticalTitles = "";
$countCritical = 0;
$countOptional = 0;
$countHidden = 0;

# ------------------------------------------------------------------------- #
# If no updates are required - it writes the info to a temp file, sends it to Zabbix server and exits
# ------------------------------------------------------------------------- #

if ($updates.Count -eq 0) {

	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$Countcriticalnum
	Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden
    Write-Host "`t There are no pending updates" -ForeGroundColor "Green"
	
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7 -s $ok
	
	exit $returnStateOK
}

# ------------------------------------------------------------------------- #
# This part counts the number of updates to be applied
# ------------------------------------------------------------------------- #

foreach ($update in $updates) {
	if ($update.IsHidden) {
		$countHidden++
	}
	elseif ($update.AutoSelectOnWebSites) {
		$criticalTitles += $update.Title + " `n"
		$countCritical++
	} else {
		$countOptional++
	}
}

# ------------------------------------------------------------------------- #
# This part writes the number of each update required to a temp file and sends it to Zabbix
# ------------------------------------------------------------------------- #

if (($countCritical + $countOptional) -gt 0) {

	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$Countcriticalnum
	Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden
    Write-Host "`t There are $($countCritical) critical updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countOptional) optional updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countHidden) hidden updates available" -ForeGroundColor "Yellow"
	
    & $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7 -s $ok
}   

# ------------------------------------------------------------------------- #


# ------------------------------------------------------------------------- #
# If Hidden Updates are found, this part will write the info to a temp file, send to Zabbix server and exit
# ------------------------------------------------------------------------- #

if ($countHidden -gt 0) {
	
	$countCritical | Out-File -Encoding "ASCII" -FilePath $env:temp$Countcriticalnum
	Write-Output "- Winupdates.Critical $($countCritical)" | Out-File -Encoding "ASCII" -FilePath $env:temp$Senderargcountcritical
	Write-Output "- Winupdates.Optional $($countOptional)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountOptional
	Write-Output "- Winupdates.Hidden $($countHidden)" | Out-File -Encoding "ASCII" -FilePath $env:temp$SenderargcountHidden
    Write-Host "`t There are $($countCritical) critical updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countOptional) optional updates available" -ForeGroundColor "Yellow"
    Write-Host "`t There are $($countHidden) hidden updates available" -ForeGroundColor "Yellow"
    
	
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargUpdateReboot -s $ok 
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderarglastupdated -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$Senderargcountcritical -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountOptional -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg4 $env:temp$SenderargcountHidden -s $ok
	& $Sender $Senderarg1 $Senderarg2 $Senderarg3 $Senderarg5 $Senderargupdating $Senderarg6 $Senderarg7 -s $ok
	
	exit $returnStateOK
}

exit $returnStateUnknown
