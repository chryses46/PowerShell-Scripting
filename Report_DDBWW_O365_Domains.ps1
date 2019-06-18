


    $list = Import-CSV -path "D:\required_files\DDBWW_MBAudit.csv"

# Iterate through the entire list. Each iteration is one connection to the specified MSOL/EOL environment.

$msoldomains = @()

$list | ForEach-Object { 

    $un = $_.Username;
    $pw = $_.Key | ConvertTo-SecureString
    $loc= $_.Office;

    Write-Host "Attempting to connect to $loc Office 365 Portal. Please wait..."

    $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $un, $pw
    Connect-MsolService -Credential $Creds -ErrorAction SilentlyContinue


    $EOLConn = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
    Import-PSSession $EOLConn

    $domain= (Get-MsolDomain | Where-Object {$_.isDefault -eq $true}).Name
    $exdomain= (Get-OrganizationConfig).Name

Get-MsolDomain | ForEach-Object {

$name = $_.name
$status = $_.status
$auth = $_.authentication
$tenant = $exdomain

             $values = [Ordered]@{
                "Domain"= $name
                "Tenant" = $tenant
                "Authentication" = $auth
                "Status"= $status
            
            }

            $domaininfo = [pscustomobject]$values

            $msoldomains +=  $domaininfo
            }

Get-PSSession | Remove-PSSession

}

$date= get-date -Format MMddyy

$csvFilePath = "D:\reports\DDBWW_O365_Domains_Report_$date.csv"
$msoldomains | Export-Csv -Path $csvFilePath -NoTypeInformation
$xlsxFilePath = Convert-CsvToXlsx($csvFilePath);

Start-Sleep -Seconds 15

$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
#wwpowershellreporting@ddb.com
$to= "wwpowershellreporting@ddb.com"
$subject= "DDB World Wide Office 365 Domain Report"
$body= "The following attached xlsx document lists domains and their properties for DDB World Wide offices."

# Send the mail message #ADAdmins@ddb.com

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
    -Attachments $xlsxFilePath