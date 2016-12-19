
# ******************************************************************************************************
$szdate = get-date -Format MMM-dd
$szLogFileName = ".\Logs\LogFile2_$szdate.txt"

# **** Log Functions ****
function Log([string]$szMsg) 
{ 
  write-host $szMsg
  Log2File $szMsg
}
function LogInfo([string]$szMsg) 
{ 
  write-host $szMsg -foregroundcolor green
  Log2File $szMsg
}
function LogWarning([string]$szMsg) 
{ 
  write-host $szMsg -foregroundcolor yellow
  Log2File $szMsg
}
function LogError([string]$szMsg) 
{ 
  write-host $szMsg -foregroundcolor red
  Log2File $szMsg
}
function LogDebug([string]$szMsg) 
{
  #write-host $szMsg -foregroundcolor blue
  #Log2File $szMsg
}
function LogPSCommand([string]$szMsg) 
{ 
  write-host $szMsg -foregroundcolor Cyan
  Log2File $szMsg
}
function Log2File([string]$szMsg) 
{
  if( $Script:szLogFileName -ne "" )
  {
    $szlongdate = get-date -format s
	Add-content $Script:szLogFileName -value "$szlongdate  $szMsg"
	#Add-content $Script:szLogFileName -value $szMsg
  }
}
# **** End Log Functions ****

function Check4Exchange{
#Add Exchange 2010 snapin if not already loaded in the PowerShell session
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	LogDebug "Loading the Exchange Server PowerShell snapin"

	try{
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction STOP
	}
	catch{
		#Snapin was not loaded
		LogWarning "The Exchange Server PowerShell snapin did not load."
		LogWarning $_.Exception.Message
		EXIT
	}
	. $env:ExchangeInstallPath\bin\RemoteExchange.ps1
	Connect-ExchangeServer -auto -AllowClobber
	}
}

# ******************************************************************************************************
# ******************************************************************************************************

# **** Begin Script ****

#Parameter for the number of messages a sender needs to send to trigger the email
[int]$recipientcount=$args[0]
#Parameter for the number of hours to go back to search the logs
[int]$hours=$args[1]

# if the recipient count is not provided in the parameters, ask. The default is 100
if (!$recipientcount){
	[int]$recipientcount=read-host "`rEnter the min number of recipients (100)"
}
if (!$recipientcount){
	[int]$recipientcount = 100
}

# if the hours parameter is not provided, ask. The default is 6
if (!$hours){
	[int]$hours=read-host "`rEnter number of hours back to search (6)"
}
if (!$hours){
	[int]$hours = 6
}

# check to see if Exchange 2010 module is loaded
Check4Exchange

# Gather message tracking logs and start building object
$endtime = [datetime]::now
$starttime = [datetime]::now.addhours(-$hours)

$servers = Get-TransportServer
$foundmessages = @()

foreach ($server in $servers){
	LogInfo "Retrieving logs from $server"
	$messages = get-messagetrackinglog -Server $server.name -Start $starttime -End $endtime -eventid RECEIVE -resultsize unlimited
	foreach ($message in $messages) {
		if ($message.recipientcount -ge $recipientcount -and $message.sender -like '*@uwaterloo.ca') {
			$foundmessages += $message
		}
	}
}

LogInfo "`n-------------------------------------------------"
LogInfo "     Sent Mail with: "
LogInfo "        Recipient Count $recipientcount or more"
LogInfo "        in the last $hours hours"
LogInfo "-------------------------------------------------"

If ($foundmessages.count -eq 0){
	LogWarning "`nNo messages found`n"
}
Else{
	$foundmessages | sort recipientcount -Descending | ft @{Expression={$_.recipientcount};Label="Count"},sender,messagesubject -auto
}




