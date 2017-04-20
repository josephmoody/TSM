# TSM Hosts. IP or DNS Name (without domain) or a combination
$tsmHosts = @("drc-ces-01", "drc-chs-01", "drc-cms-01") ###### CHANGE ME ######

# TSM IP addresses example
#$tsmHosts = @("10.2.5.119", "10.2.5.112")

# TSM OU Name Example
#$tsmHosts = Get-ADComputer -Filter * -SearchBase "OU=TSM,OU=Servers,DC=Test,DC=local" | Sort name | where Name -NE TSM-Access | select -ExpandProperty name

# TSM Domain. Change to your domain.
$tsmDomain = "polk.k12.ga.us" ###### CHANGE ME ######