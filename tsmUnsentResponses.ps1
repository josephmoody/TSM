# Date:   2017-04-17
# Description:
#   This script performs a web request and check for any responses
#   If the script determines that there are responses on the TSM
#   we perform an additional webrequest to then transmit the responses
#   There is also lots of logging for debugging later.

# Include tsmHostVars File
# Use the tsmHostVars file for loading your TSMs
# This will prevent needing to alter the file
# when updates are made to the main functions below
. .\tsmHostVars.ps1


# These vars have been left for visibility
# TSM Hosts. IP or DNS Name (without domain) or a combination
#$tsmHosts = @("drc-ces-01", "drc-chs-01", "drc-cms-01") ###### CHANGE ME ######

# TSM IP addresses example
#$tsmHosts = @("10.2.5.119", "10.2.5.112")

# TSM OU Name Example
#$tsmHosts = Get-ADComputer -Filter * -SearchBase "OU=TSM,OU=Servers,DC=Test,DC=local" | Sort name | where Name -NE TSM-Access | select -ExpandProperty name

# TSM Domain. Change to your domain.
#$tsmDomain = "polk.k12.ga.us" ###### CHANGE ME ######


# Set log file location. Default is current users Desktop
#$tsmLogs = "$($env:USERPROFILE)\Desktop\$(get-date -format "yyyy-MM-dd")_tsmResponses.log"
$tsmLogs = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -name Desktop).Desktop + "\" + (Get-Date -Format "yyyy-MM-dd") + "_tsmResponses.log"

# Amount of time to before webrequest timesout
$reqTimeOut = 15

# URL to check for UnSent responses
$tsmUnSentURL = ":8080/studentResponse/unsent"

# URL to Transmit Responses
$tsmTransmittURL = ":8080/studentResponse/transmitResponses"

# Get date Log format
function getLogDate {

    return (get-date -Format "yyyy-MM-dd HH:mm:ss")

}

# Function to parse the content from the WebRequest
# Thanks to Joseph Moody for the RegEx
function getUnsentCount($htmlContent) {

    $UnSentCount = $htmlContent.ToString() -split "[`r`n]" | Select-String -SimpleMatch 'unsentCount'
    $UnSentCount = ($UnSentCount) -replace '\D+(\d+)\D+','$1'

    # Return Number of Unsent Responses
    return $UnSentCount

}


function tsmWebRequest($tsmHostname, $tsmTrans) {

    # Construct Url link
    if($tsmTrans -gt 0) {
        $link = "http://" + $tsmHostname + $tsmTransmittURL
    } else {
        $link = "http://" + $tsmHostname + $tsmUnSentURL
    }

    
        
    # Try creating WebRequest and log errors
    try {
        $html = Invoke-WebRequest -Uri $link -TimeoutSec $reqTimeOut -DisableKeepAlive
    } catch [System.Net.WebException] {
            
        # Timed out Exception
        if($_.Exception.ToString() -like "*operation has timed out*") {
                
            $result = (getLogDate) + " - Error: " + $tsm + " - WebRequest timed out"
        } else {
            # Write Unhandled exceptions
            $result = (getLogDate) + " - Unhandled Error: " + $tsm + " - " + $_.Exception.ToString()
        }

        # Write result to host
        Write-Host -ForegroundColor Red $result

        # Write to Log file
        Out-File -FilePath $tsmLogs -Append -InputObject $result

        return $null

    }

    # Return WebRequest
    return $html

}

# Check for Responses
function tsmCheckResponses {
    foreach ($tsm in $tsmHosts) {

        # Determine wether uri is IP or DNS Name
        try {
            $hostname = [ipaddress]$tsm
        } catch {
            $hostname = $tsm + "." + $tsmDomain
        }
        
        $tsmStatus = tsmWebRequest $hostname 0

        # Process Data if html var exists
        
        if($tsmStatus -ne $null) {
            $resNum = 0
            # Get number of responses        
            $resNum = getUnsentCount($tsmStatus.Content)

            #$resNum = 1 # test to submit responses

            $result = (getLogDate) + " - WebStatusCode: " + $tsmStatus.StatusCode + " - " + $tsm + " has $($resNum) responses"

            #Clear-Variable $tsmStatus

            if($resNum -gt 0) {
                
                # Write Number of Responses
                Write-Host -ForegroundColor Green $result
                Out-File -FilePath $tsmLogs -Append -InputObject $result

                # Transmit Responses
                $tsmTransmit = tsmWebRequest $hostname 1
                
                if($tsmTransmit -ne $null) {
                
                    $transmitResponses = getUnsentCount($tsmTransmit.Content)
                    $result = (getLogDate) + " - WebStatusCode: " + $tsmTransmit.StatusCode + " - " + $tsm + " has $($transmitResponses) responses that have not been transmitted"

                    # Display number of responses that did not transmit
                    if($transmitResponses -gt 0) {

                        Write-Host -ForegroundColor Red $result
                        Out-File -FilePath $tsmLogs -Append -InputObject $result
                    } 

                    # Display if all responses where transmitted
                    else {
                        $result = (getLogDate) + " - WebStatusCode: " + $tsmTransmit.statusCode + " - " + $tsm + " transmitted responses successfully"
                        Write-Host -ForegroundColor Green $result
                    }
                } 

                # Could not complete web request
                else {
                    $result = (getLogDate) + " - " + $tsm + " Invoke WebRequest Failed. Could not transmit responses"
                    Write-Host -ForegroundColor Red $result

                    Out-File -FilePath $tsmLogs -Append -InputObject $result
                }

            } else {
                Write-Host $result
            }

        } else { # If tsmWebResponse is null

            $result = (getLogDate) + " - " + $tsm + " Invoke WebRequest Failed"
            Write-Host -ForegroundColor Red $result

            Out-File -FilePath $tsmLogs -Append -InputObject $result
        }
        
        
     }

}


# Perform TSM Check Responses Loop

while(1) {

    tsmCheckResponses # Check for responses
    Start-Sleep 8     # Sleep for number of seconds

}
