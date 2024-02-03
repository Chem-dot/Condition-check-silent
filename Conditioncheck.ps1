#Places the log file in same directory as the script is running from
$Logfile = "$PSScriptRoot\logs.log"

Function LogWrite
{
   Param ([string]$logstring)
   $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
   $logWithTimestamp = "${timestamp}: $logstring"
   Add-content $Logfile -value $logWithTimestamp
}

# Initialize the counter and a flag variable
$tries = 0
$conditionsMet = $false


# Loop up to 10 times
LogWrite "Starting..."
while ($tries -lt 10 -and -not $conditionsMet) {
    # Increment the counter
    $tries++
    
    # This gets the output from 'netsh wlan show interfaces' and stores it in a variable
    $wlanOutput = netsh wlan show interfaces
    # This will filter the output to find the line with 'SSID' and then split it on ':' to extract the SSID name
    $connectedSSID = $wlanOutput | Select-String -Pattern 'SSID\s+:\s+(.+)$' | ForEach-Object { $_.Matches.Groups[1].Value.Trim() }
    # Now we check if the trimmed SSID matches the one we're looking for
    # Make sure this matches the exact SSID of your network
    $targetSSID = "My place" 
    $targetSSID1 = "Wifi1"

    # VPN check
    $targetProfileName = "COMPANYDOMAIN.NET" # The network profile name to check for
    # Get the list of network profiles on the system
    $netProfiles = Get-NetConnectionProfile
    # Check if the desired network profile is currently connected
    $connectedProfile = $netProfiles | Where-Object { $_.Name -eq $targetProfileName }

    # Check if the Outlook process is running
    $process = Get-Process Outlook -ErrorAction SilentlyContinue

    if ($process) {
        LogWrite "Outlook is running. Checking for network connections" 
        if ($connectedProfile -or $connectedSSID -eq $targetSSID -or $connectedSSID -eq $targetSSID1){
            LogWrite "Connection found continueing..." 
            # Set the flag to true as both conditions are met
            $conditionsMet = $true
            # Break the loop since the conditions are met
            break                          
        } else {
            LogWrite "Network not found. Attempt $tries of 10. Trying again in 60 seconds..."         
            Start-Sleep -Seconds 60
            # If the conditions are not met the script will sleep for 60 seconds before trying the loop again.
        } 
    } else {
        LogWrite "Outlook is not running. Attempt $tries of 10. Trying again in 60 seconds..."
        Start-Sleep -Seconds 60
    }
    }

# Check the flag after the loop to perform actions if both conditions were met
if ($conditionsMet) {
    LogWrite "Both conditions were met. Performing the actions..." 
    #Replace \\NETWORKSHARE... with the actual share you want to run.
    & "\\NETWORKSHARE\FOLDER\EXAMPLE.SCRIPT"
}
else {
    LogWrite "Conditions not met. I've looked $tries times! I give up :("
}

# if more actions have to be taken continue from here.