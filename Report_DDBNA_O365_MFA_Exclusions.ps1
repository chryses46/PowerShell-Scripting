 
#GetMFA-Excluded Users for DDB NA Office 365 Tenants

$list = Import-CSV -path "D:\required_files\DDBNA_MBAudit.csv"

$ExcludeList = @()
 
 $list | ForEach-Object{

    $un = $_.Username;
    $pw = $_.Key | ConvertTo-SecureString
    $loc= $_.Office;

    
    $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $un, $pw
    Connect-MsolService -Credential $Creds

    $EOLConn = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
    Import-PSSession $EOLConn

    $exdomain= (Get-OrganizationConfig).Name

    $MFAExcludeID = (Get-MsolGroup -all | Where-Object {$_.displayname -eq "MFA Service Account Exclusion"}).ObjectID

    Get-MsolGroupMember -GroupObjectId $MFAExcludeID | ForEach-Object {

    $Props = [Ordered]@{
                    Email= $_.emailaddress
                    Tenant= $exdomain
                    }

                $obj = [pscustomobject]$Props
                $ExcludeList += $obj
            }

    Get-PSSession | Remove-PSSession

    }

$date= get-date -Format MMddyy

$csvFilePath = "D:\Reports\DDBNA_O365_MFA_Exclude_List_$date.csv"

$ExcludeList | Export-Csv -Path $csvFilePath -NoTypeInformation

#Convert to XLSX
$xlsxFilePath = Convert-CsvToXlsx($csvFilePath);

Start-Sleep -Seconds 15

$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
$to= "ddbna_o365Admins@ddb.com"
$subject= "DDBNA MFA Exclusion List"
$body= "Please see attached."

# Send the mail message

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High -Attachments $xlsxFilePath

