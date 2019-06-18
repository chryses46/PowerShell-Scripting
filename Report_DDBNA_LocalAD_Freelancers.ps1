Import-Module ConvertCsvToXlsx

$SboxDomains = Import-Csv "D:\required_files\SBOXDomains.csv"

$credential= Import-Csv -Path "D:\required_files\SBOX.csv"

$un= $credential.username

$pw= $credential.key | ConvertTo-SecureString

$SBOXFreelancers = @()

$ExpiresDays = 12

$Today = (Get-date)

$date= get-date -Format MMddyy

Foreach($TargetDomain in $SboxDomains){

    $Domain = $TargetDomain.Domain

    $credun="$($Domain)\$($un)"
    
    $DomainCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credun, $pw
    
   
    # Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired

    Import-Module ActiveDirectory
    if ($Domain -eq "interbrand.internal"){
        $Freelancers = Get-ADUser -Filter {wwwhomepage -like "Freelancer"} -Server $TargetDomain.ServerIP -Credential $DomainCredentials -Properties * |`
     Where-Object {$_.AccountExpirationDate -ne $null -and $_.DistinguishedName -notlike "*Disabled_Users*" -and $_.DistinguishedName -notlike "*AD_Cleanup*"} |`
     Where-Object{$_.DistinguishedName -like "*OU=Temporary,OU=Users,OU=Cincinnati,DC=interbrand,DC=internal*" -or $_.DistinguishedName -like "*OU=Temp,OU=Users,OU=Dayton,DC=interbrand,DC=internal" -or $_.DistinguishedName -like "*OU=Temp,OU=Users,OU=IB Group,DC=interbrand,DC=internal" -or $_.DistinguishedName -like "*OU=Temp,OU=Users,OU=IB Health,DC=interbrand,DC=internal" -or $_.DistinguishedName -like "*OU=Temp,OU=Users,OU=Los Angeles,DC=interbrand,DC=internal" -or $_.DistinguishedName -like "*OU=IBNY Temp,OU=Users,OU=New York,DC=interbrand,DC=internal" -or $_.DistinguishedName -like "*OU=Temp,OU=Users,OU=San Francisco,DC=interbrand,DC=internal"} |`
     Select Name, Mail, AccountExpirationDate, City, Title, distinguishedName, company
    }
    else{
        $Freelancers = Get-ADUser -Filter {wwwhomepage -like "Freelancer"} -Server $TargetDomain.ServerIP -Credential $DomainCredentials -Properties * |`
     Where-Object {$_.AccountExpirationDate -ne $null -and $_.DistinguishedName -notlike "*Disabled_Users*" -and $_.DistinguishedName -notlike "*AD_Cleanup*"} |`
     Select Name, Mail, AccountExpirationDate, City, Title, distinguishedName, company
    }

    if($Freelancers -ne $null){

        ForEach($User in $Freelancers) {
    
         $AccountExpirationDate = ($user.AccountExpirationDate).Date
         $DaysToExpire = (New-TimeSpan -Start $Today -End $AccountExpirationDate).Days

            if($DaysToExpire -le $ExpiresDays){

                $UserValues = [Ordered]@{
                    Name= $user.Name;
                    "Account Expiration Date" = $user.AccountExpirationDate;
                    Email = $user.mail;
                    City = $user.City;
                    Company = $user.company;
                    Title = $user.Title;
                    Domain = $Domain;
                    DN = $user.distinguishedName
                }

                $UserEntry = [pscustomobject]$UserValues
                $SBOXFreelancers +=  $UserEntry
            }
        }
    }

    Get-PSSession | Remove-PSSession
}
$csvFilePath = "D:\reports\DDB_North_America_Freelancer_Report_$date.csv"

$SBoxFreelancers | Export-Csv -Path $csvFilePath -NoTypeInformation

$xlsxFilePath = Convert-CsvToXlsx($csvFilePath);

Start-Sleep -Seconds 15

$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
#ADAdmins@ddb.com
$to= "ADAdmins@ddb.com"
$subject= "DDB North America Freelancer Report"
$body= "The following attached xlsx document lists user accounts that will expire in 12 days or less. Please check the list."

# Send the mail message #ADAdmins@ddb.com

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
    -Attachments $xlsxFilePath