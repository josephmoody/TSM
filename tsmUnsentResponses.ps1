# Date:   2017-04-17
# Authors: David Thompson, Joseph Moody
# Description:
#   This script performs a web request and check for any responses
#   If the script determines that there are responses on the TSM
#   we perform an additional webrequest to then transmit the responses
#   There is also lots of logging for debugging later.

# Include tsmHostVars File
# Use the tsmHostVars file for loading your TSMs
# This will prevent needing to alter the file
# when updates are made to the main functions below

$tsmVarFile = ".\tsmHostVars.ps1"
if(Test-Path $tsmVarFile) {
    . $tsmVarFile
} else {
    write-host "Could not locate TSM Var File. Please change file tsmVarFile variable to file location or from powershell navigate to folder containing files."
    pause
    exit
}


# These vars have been left for visibility. Configuration should be done is tsmHostVars file.
# TSM Hosts. IP or DNS Name (without domain) or a combination
#$tsmHosts = @("drc-ces-01", "drc-chs-01", "drc-cms-01") ###### CHANGE ME ######

# TSM IP addresses example
#$tsmHosts = @("10.2.5.119", "10.2.5.112")

# TSM OU Name Example
#$tsmHosts = Get-ADComputer -Filter * -SearchBase "OU=TSM,OU=Servers,DC=Test,DC=local" | Sort name | where Name -NE TSM-Access | select -ExpandProperty name

# TSM Domain. Change to your domain.
#$tsmDomain = "polk.k12.ga.us" ###### CHANGE ME ######


# Create Object with webfailure
$tsmWeb = @()

foreach ($t in $tsmHosts) {

    $tsmSetup = @{"TSM"=$t;"WebFail"=0}
    $tsmWeb += New-Object PSCustomObject -Property $tsmSetup

}

# studentResponses
$tsmStudentResponses = @()

# Number of UnSent Responses before alert
$previousunSentResponseAlert = 3

# Set log file location. Default is current users Desktop
$tsmLogs = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -name Desktop).Desktop + "\" + (Get-Date -Format "yyyy-MM-dd") + "_tsmResponses.log"

# Log for storing responses
$tsmLogStudent = (Get-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders\" -name Desktop).Desktop + "\" + (Get-Date -Format "yyyy-MM-dd") + "_tsmStudentResponses.log"

# Amount of time to before webrequest timesout
$reqTimeOut = 20

# URL to check for UnSent responses
$tsmUnSentURL = ":8080/studentResponse/unsent"

# URL to Transmit Responses
$tsmTransmittURL = ":8080/studentResponse/transmitResponses"

# Get date Log format
function getLogDate {

    return (get-date -Format "yyyy-MM-dd HH:mm:ss")

}

# Check if log files exists
If ((Test-Path -Path $tsmLogs) -eq 0) {
    $logText = (getLogDate) + " - Start TSM Log"
    Out-File -FilePath $tsmLogs -Append -InputObject $logText
}

If ((Test-Path -Path $tsmLogStudent) -eq 0) {
    $logText = (getLogDate) + " - Start TSM Log"
    Out-File -FilePath $tsmLogStudent -Append -InputObject $logText
}

# Function to parse the content from the WebRequest
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

        # Add exception to errored tsm
        $addWebRequest = $tsmWeb | where {$_.TSM -eq $tsm}
        $addWebRequest.WebFail = ($addWebRequest.WebFail + 1)
            
        # Timed out Exception
        if($_.Exception.ToString() -like "*operation has timed out*") {
                
            $result = (getLogDate) + " - Error: " + $tsm + " - " + $addWebRequest.WebFail + " WebRequests have timed out."
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

            $result = (getLogDate) + " - WebStatusCode: " + $tsmStatus.StatusCode + " - " + $tsm + " has $($resNum) responses"

            if($resNum -gt 0) {
               
                # Write Number of Responses
                Write-Host -ForegroundColor Green $result
                Out-File -FilePath $tsmLogs -Append -InputObject $result 
          
                # Begin Convert Responoses Table Body from String to Object 
                $unSentDataString = $tsmStatus.AllElements | where id -EQ responsesTableBody | select -ExpandProperty innerhtml

      
                #Check for Multiple Table Entries and Spilt
                $MultipleTabeEntryObjects = 1
                while ($MultipleTabeEntryObjects -le $resNum) {
                
                    $MultipleTableEntriesIndex = $unSentDataString.IndexOf("</tr>")
                
                
                    Set-Variable -Name "unSentDataString$MultipleTabeEntryObjects" -Value  $unSentDataString.Substring(0,$MultipleTableEntriesIndex)

                    $CurrentStringVariable = Get-Variable -Name "unSentDataString$MultipleTabeEntryObjects" -ValueOnly

                    $CurrentStringVariable = $CurrentStringVariable.Replace("<tr>",'')
                    $CurrentStringVariable = $CurrentStringVariable.Replace("</tr>",'')
  
                    $SchoolIndex = $CurrentStringVariable.IndexOf("<td>")
                    $CurrentStringVariable = $CurrentStringVariable.Substring($SchoolIndex)
                
                
                    Set-Variable -name "unSentDataObject$MultipleTabeEntryObjects" -Value $CurrentStringVariable
                    $CurrentDataVariable = Get-Variable -Name "unSentDataObject$MultipleTabeEntryObjects" -ValueOnly
                
                    $CurrentDataVariable = $CurrentStringVariable | ConvertFrom-String -Delimiter "<td>" -PropertyNames NA,School,TestSession,Student,GTID,EarliestResponse

                    $CurrentDataVariable.PSObject.Properties.Remove('NA')
               
                    $unSentDataSchoolIndex = $CurrentDataVariable.School.IndexOf("<")
                    $CurrentDataVariable.School = $CurrentDataVariable.School.Substring(0,$unSentDataSchoolIndex)

                    $unSentDataStudentIndex = $CurrentDataVariable.Student.IndexOf("<")
                    $CurrentDataVariable.Student = $CurrentDataVariable.Student.Substring(0,$unSentDataStudentIndex)

                    $unSentDataGTIDIndex = $CurrentDataVariable.GTID.IndexOf("<")
                    $CurrentDataVariable.GTID = $CurrentDataVariable.GTID.Substring(0,$unSentDataGTIDIndex)

                    $unSentDataTestSessionIndex = $CurrentDataVariable.TestSession.IndexOf("<")
                    $CurrentDataVariable.TestSession = $CurrentDataVariable.TestSession.Substring(0,$unSentDataTestSessionIndex)

                    $unSentDataEarliestResponseIndex = $CurrentDataVariable.EarliestResponse.IndexOf("<")
                    $CurrentDataVariable.EarliestResponse = $CurrentDataVariable.EarliestResponse.Substring(0,$unSentDataEarliestResponseIndex)

                    $intCurrentGtidCount = 1
                    if(($tsmStudentResponses | measure).Count -gt 0) {
                        $currentGtidCount = $tsmStudentResponses | where { $_.GTID -eq $CurrentDataVariable.GTID }

                        if(($currentGtidCount | measure).count -gt 0) {
                            $intCurrentGtidCount = $currentGtidCount.responseCount + 1
                            $currentGtidCount.GTID = $intCurrentGtidCount # Add an additional response to GTID
                        } else {
                            # Add GTID to object if student does not exist
                            
                        }
                    } else {

                        # Add new object if object is null
                        $addCurrentDataVar = @{"School"=$CurrentDataVariable.School;"Student"=$CurrentDataVariable.Student;"GTID"=$CurrentDataVariable.GTID;"TestSession"=$CurrentDataVariable.TestSession;"EarliestResponse"=$CurrentDataVariable.EarliestResponse;"responseCount"=$intCurrentGtidCount}
                        $tsmStudentResponses += New-Object PSCustomObject -Property $addCurrentDataVar

                    }

                    #Search TSM Unsent Resonses Log for previous GTID entry
                    
                    $previousunSentResponseAlertforGTID = 0
                    $previousunSentResponseAlertforGTID = $intCurrentGtidCount # Use already counted gtids

                    if ($previousunSentResponseAlertforGTID -ge $previousunSentResponseAlert){
                    $previousGTIDFound = $CurrentDataVariable.School + ": " + $CurrentDataVariable.Student + " has " + $previousunSentResponseAlertforGTID + " previous unsent responses ending at " + $CurrentDataVariable.EarliestResponse + ". Teacher name: " + $CurrentDataVariable.TestSession
                    Write-Host -ForegroundColor Yellow $previousGTIDFound
                    Out-File -FilePath $tsmLogs -Append -InputObject $previousGTIDFound
                }

                # Transmit Responses
                $tsmTransmit = tsmWebRequest $hostname 1

                #Output unsent data to TSM Student Log
                Out-File -FilePath $tsmLogStudent -Append -InputObject $CurrentDataVariable
             
                $MultipleTabeEntryObjects++
                $unSentDataString = $unSentDataString.Remove(0,$MultipleTableEntriesIndex+5)
            }


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