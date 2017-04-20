#Report should only be used against one type of TSM (WIDA or EOC/EOG).#First TSM should be online - if offline, only name and status column will populate.#Use only one of the two "$TSMServers =" input lines.    #Gather List of TSM Server Names through Active Directory. The last #| section shows how to exclude certain machine names.#$TSMServers = Get-ADComputer -Filter * -SearchBase "OU=TSM,OU=Servers,DC=TEST,DC=local" | sort Name #| where name -NE TSM-WIDA-01. .\tsmHostVars.ps1#$TSMServers = New-Object -TypeName psobject -Property ([ordered] @{#                    Name = "drc-ces-01.polk.k12.ga.us"#                    })#Gather List of TSM Server Names through CSV Import (header should be: Name)#$TSMServers = import-csv .\TSMServer.csv$TSMsInfo = @()# Start looping through each serverforeach ($TSMServer in $tsmHosts){    Clear-Variable StrVersion,StrDomain,StrIP,Content*,Uri,TSMInfo,Session -ErrorAction SilentlyContinue    # Construct Server URI    $URI = "http://" + $TSMServer + ":8080"    write-host Now processing $TSMServer    # Create WebRequest to Server    $Session = Invoke-WebRequest -Uri $URI -SessionVariable TSMSession    #Populate Bad Session Object    if ($Session -eq $null) {        $TSMInfo = New-Object -TypeName psobject -Property (                    [ordered] @{                        Name = $TSMServer                        Status = "Bad"                        }                    )        # Add Server to object        $TSMsInfo += $TSMInfo        Clear-Variable TSMInfo,Session -ErrorAction SilentlyContinue        continue    }        #Populate Good Session Object    if ($Session -ne $Null) {        # Convert Session to string and Split once        $strSession = $Session.ToString() -split "[`r`n]"        # Get Version Number        $Version = $strSession | Select-String "app version"        $StrVersion = $Version.ToString()        $StrVersion = $StrVersion.replace("    app version        ",'')            # Get Server Name        $Name = $strSession | Select-String -SimpleMatch '"TSM Name" value='        $StrName = $Name.ToString()        $StrName = $StrName.replace('						<input type="text" id="inputTSMName" class="input-xlarge" placeholder="TSM Name" value="','')        $StrName = $StrName.replace('" maxlength="40">','')        # Get DRC CentralOffice ServerName        $Domain = $strSession | Select-String -SimpleMatch 'TSM Server Domain:'        $StrDomain = $Domain.ToString()        $StrDomain = $StrDomain.Replace('			  		<label class="control-label"><strong>TSM Server Domain:</strong> ','')        $StrDomain = $StrDomain.Replace('.drc-centraloffice.com</label>','') # Removed DRC Central Office domain and trailing Label tag        # Get IP address        $IP = $strSession | Select-String -SimpleMatch 'TSM Server IP:'        $StrIP = $IP.ToString()        $StrIP = $StrIP.Replace('			  		<label class="control-label"><strong>TSM Server IP:</strong> ','')        $StrIP = $StrIP.Replace('</label>','')        # Get Cached Content        $Content = $strSession | Select-String -SimpleMatch '<span class="loadingError label label-important status" style="display: none;">'        $ContentCount = $Content.Count - 1         $ContentNumber = 0        while ($ContentNumber -le $ContentCount)        {            $ContentValue = $strSession | Select -index ($Content[$ContentNumber].LineNumber - 3)            $ContentValue = $ContentValue.Replace('										<td>','')            $ContentValue = $ContentValue.Replace('<br>','')            New-Variable -name "ContentName$ContentNumber" -Value $ContentValue -Force                                    $ContentUpdate = $strSession | Select -index ($Content[$ContentNumber].LineNumber + 11)            $ContentUpdate = $ContentUpdate.Split(">")[1]            $ContentUpdate = $ContentUpdate.Replace('</span','')            if(($ContentUpdate -ne "Up to Date") -eq $True) {                $ContentUpdate = '<FONT COLOR="ff0000">Out of Date</FONT>'            }            Clear-Variable ContentTTS -ErrorAction SilentlyContinue            $ContentTTS = $Session.ToString() -split "[`r`n]" | Select -index ($Content[$ContentNumber].LineNumber + 33)            $ContentTTS = $ContentTTS.Replace('													<input type="hidden" name="_downloadTTSEnabled" /><input type="checkbox" name="downloadTTSEnabled" checked="','')            $ContentTTS = $ContentTTS.Split('"')[0]                            if (($ContentTTS -ne "checked") -eq $True) {                $ContentTTS = '<FONT COLOR="ff0000">disabled</FONT>'            }                                        $ContentAttributes = "<dl><dt>Content</dt><dd> $($ContentUpdate)</dd><dt>TTS</dt><dd>$($ContentTTS)</dd></dl>"                                   New-Variable -name "ContentAttributes$ContentNumber" -Value $ContentAttributes -Force            $ContentNumber++                    }                            $TSMInfo = New-Object -TypeName psobject -Property ([ordered] @{                    Name = "<a href='$($URI)'>$($StrName)</a>"                    Status = "Good"                    Version = $StrVersion                    # Added the Central Office Domain to the Table Header                    'Domain (.drc-centraloffice.com)' = $StrDomain                    IP = $StrIP                    $ContentName0 = $ContentAttributes0                    $ContentName1 = $ContentAttributes1                    $ContentName2 = $ContentAttributes2                    $ContentName3 = $ContentAttributes3                    $ContentName4 = $ContentAttributes4        })        $TSMsInfo += $TSMInfo    }}# Assemble the HTML Header and CSS for our Report
$Head = @"
<title>TSM Server Report</title>
<style type="text/css">
<!--
    body { font-family: Verdana, Geneva, Arial, Helvetica, sans-serif; }
 
    #report { width: 835px; }
 
    table {
        border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;
        font: 10pt Verdana, Geneva, Arial, Helvetica, sans-serif;
    }
 
    table td {
        border-width: 1px;padding: 3px;border-style: solid;border-color: black;
        font-size: 12px;
        text-align: left;
        white-space: nowrap;
    }
 
    table th {
        border-width: 1px;padding: 3px;border-style: solid;border-color: black;background-color: #6495ED;
        font-size: 12px;
        font-weight: bold;
        text-align: left;
    }
 
    h2 { clear: both; font-size: 130%; }
 
    h3 {
            clear: both;
            font-size: 115%;
            margin-left: 20px;
            margin-top: 30px;
    }
 
    p { margin-left: 20px; font-size: 12px; }

    dt { display: block; float: left; width: 125px; text-align: right; }
    dt:after {content:':';}
    dd { display:block;}
 
    table.list { float: left; }
 
        table.list td:nth-child(1) {
            font-weight: bold;
            border-right: 1px grey solid;
            text-align: right;
    }
 
    table.list td:nth-child(2) { padding-left: 7px; }
    table tr:nth-child(even) td:nth-child(even) { background: #CCCCCC; }
    table tr:nth-child(odd) td:nth-child(odd) { background: #F2F2F2; }
    table tr:nth-child(even) td:nth-child(odd) { background: #CCCCCC; }
    table tr:nth-child(odd) td:nth-child(even) { background: #F2F2F2; }
    table tr:Hover TD {Background-Color: #C1D5F8;}
    div.column { width: 320px; float: left; }
    div.first { padding-right: 20px; border-right: 1px  grey solid; }
    div.second { margin-left: 30px; }
    table { margin-left: 20px; }
-->
</style>
 
"@#Convert Object to HTML.Object - Use System.Web to add hyperlinks - output file and open it.$HTMLTSMInfo= $TSMsInfo | ConvertTo-Html -Head $head -PreContent "<h2>TSM Server Report</h2><br><br>" -PostContent “<br><h5>For questions or suggestions, contact Joseph@DeployHappiness.com</h5>”Add-Type -AssemblyName System.Web
[System.Web.HttpUtility]::HtmlDecode($HTMLTSMInfo) | Out-File .\TSMReport.htm -Force# Open newly created reportInvoke-Expression .\tsmreport.htm