# A script that removes disabled accounts in Local AD

$SboxDomains = Import-Csv "D:\required_files\SBOXDomains.csv"

$credential = Import-Csv -Path "D:\required_files\SBOX.csv"

$un = $credential.username

$pw = $credential.key | ConvertTo-SecureString

For ($i = 0; $i -lt $SboxDomains.Count; $i ++){

    $sboxDomain = $SboxDomains[$i].Domain

    $sboxCredUN ="$($sboxDomain)\$($un)"
    
    $connectionCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $sboxCredUN, $pw

    $aDSession = New-PSSession -ComputerName $SboxDomains[$i].ServerFQDN -Credential $connectionCredentials -Authentication Kerberos

    ####### Put in scriptblock

    Invoke-Command -Session $aDSession -ScriptBlock {
        
        $domain = (Get-ADDomain).DnsRoot

        $baseOU  =  (Get-ADDomain).DistinguishedName

        $daysInactive = 30

        if($domain -eq "interbrand.internal"){

            $searchb= "OU=Disabled_Users,OU=NorthAmerica,DC=interbrand,DC=internal"
            $searchbFre = "OU=Disabled_Users_Recurring_Freelancers,OU=NorthAmerica,DC=interbrand,DC=internal"

        }else {

            $searchb =  ("OU=Disabled_Users," + $baseOU)
            $searchbFre =  ("ou=Disabled_Users_Recurring_Freelancers," + $baseOU)

        }

        $daysInactiveFre = 180

        $time = (Get-Date).Adddays(-($DaysInactive))

        $timeFre = (Get-Date).Adddays(-($daysInactiveFre)) 

        $usersToDelete = @()

        Get-ADUser -Filter { LastLogonTimeStamp -lt $time -and enabled -eq $false } -SearchBase $searchb -Properties LastLogonTimeStamp,LastLogon,whenChanged | ForEach-Object {$usersToDelete += $_}
        
        if($domain -ne "ddb.mexico.com"){

            Get-ADUser -Filter { LastLogonTimeStamp -lt $timeFre -and enabled -eq $false } -SearchBase $searchbFre -Properties LastLogonTimeStamp,LastLogon,whenChanged | ForEach-Object {$usersToDelete += $_}
        }

        for ($j = 0; $j -lt $usersToDelete.Count; $j ++){
            
            $userDN = $usersToDelete[$j].DistinguishedName

            Write-Host "Deleting $userDN" -ForegroundColor Yellow

            Get-ADUser -Identity $userDN |  Remove-ADObject -Recursive -Confirm:$false
        }
    }
}

$_Domains = @()

Function OfficeLocations(){
    
    for ($i = 0; $i -lt $SboxDomains.Count; $i ++){

        $global:_Domains += "<li>$($SboxDomains[$i].Domain)</li>"
    } 
}

OfficeLocations
$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
#ADAdmins@ddb.com
$to= "ADAdmins@ddb.com"
$subject= "DDB NA Disabled Users Cleanup"
$body= "
    <h1>DDB NA Disabled Users Cleanup</h1>
    <p>Disabled users older than 30 days have been removed from AD in the following domains:</p>
    <br/>
    <ul>
        $_Domains
    </ul>
"

# Send the mail message

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High