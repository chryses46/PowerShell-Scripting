<# This script gets O365 users that are licensed, then displays the license associated with the User Principal Name of the user
    This script runs against all of NA DDB Office 365 Tenants.
    This script runs every Monday
#>

$list = Import-CSV -path "D:\required_files\DDBNA_MBAudit.csv"

$UserInfo = @()

# Iterate through the entire list. Each iteration is one connection to the specified MSOL/EOL environment.
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


    #Create User Array
    $MSOLList= @()

    $MSOLList += Get-MsolUser -All | Where {$_.isLicensed -eq "True"} | Select UserPrincipalname,Licenses,UsageLocation,ProxyAddresses,Office

    #Iterate through user data and format requested information
    foreach($user in $MSOLList){
        $userSMTP = (Get-mailbox $user.userprincipalname).PrimarySMTPAddress

        $Props = [Ordered]@{

            UserPrincipalName = $user.UserPrincipalName;
            Email= $userSMTP;
            License = $user.Licenses.AccountSku.SkuPartNumber;
            UsageLocation = $user.UsageLocation;
            ProxyAddresses = [string]::join(";",$user.ProxyAddresses);
            Office = $user.Office;
            ObjectID = [string]($user.ObjectId);
            Office365Tenant = $domain
         }   
        
        $obj = [pscustomobject]$Props
        $UserInfo = $UserInfo + $obj
     }

    Get-PSSession | Remove-PSSession
}


$date= get-date -Format MMddyy

$csvFilePath = "D:\reports\DDBNA_License_Report_$date.csv"

#Export requested information
$UserInfo | Export-Csv -Path $csvFilePath -NoTypeInformation

#convert to XLSX
$xlsxFilePath = Convert-CsvToXlsx($csvFilePath);

Start-Sleep -Seconds 15

$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
$to= "DDBNA_O365Admins@ddb.com"
$subject= "DDB North America Office 365 License Report"
$body= "Please see attached report."

# Send the mail message

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
            -Attachments $xlsxFilePath 