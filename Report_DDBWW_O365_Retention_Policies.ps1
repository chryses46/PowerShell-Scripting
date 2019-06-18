<# 

This script works with a securely-stored CSV file that lists locations' EOL Admin and secure password. 
The passwords are created previously by the PSService log in, to ensure proper functioning of the Scheduled Task "DDB World Wide Office 365 Mailbox Auditing, Retention and Licensing Report"

This script will enable auditing on mailboxes that are not already enabled in a target Office 365 tenant.

#>

Import-Module ConvertAppendCSVXLSX

# Import the list of user names and secured passwords

$list = Import-CSV -path "D:\required_files\DDBWW_MBAudit.csv"

$date= get-date -Format MMddyy


$_retentionPolicies = @()
$_retentionTags = @()

 # Iterate through the entire list. Each iteration is one connection to the specified MSOL/EOL environment.
$list | ForEach-Object {

    $un = $_.Username;
    $pw = $_.Key | ConvertTo-SecureString
    $loc= $_.Office;

    Write-Host "Attempting to connect to $loc Office 365 Portal. Please wait..." -ForegroundColor Yellow

    $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $un, $pw
    Connect-MsolService -Credential $Creds -ErrorAction SilentlyContinue


    $EOLConn = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
    Import-PSSession $EOLConn


    $domain= (Get-MsolDomain | Where-Object {$_.isDefault -eq $true}).Name

    $exdomain= (Get-OrganizationConfig).Name
   
    if($domain -ne $null -and $exdomain -ne $null){
        Write-Host "Connected. 365 AD Domain = $domain. Tenant Domain Name = $exdomain." -ForegroundColor Green
    }else{
        Write-Host "Connection to $loc Office 365 failed. Exiting." -ForegroundColor Red
        Get-PSSession | Remove-PSSession
        break;
    }

    # Get Retention Policies for the tenant.  
        Write-Host "Obtaining retention policies for $exdomain Exchange Online tenant." -ForegroundColor Yellow

        $retentionPoliciesArray = Get-RetentionPolicy | Select Name,RetentionPolicyTagLinks

        foreach($policy in $retentionPoliciesArray){
            
            $Props = [Ordered]@{
                "Name" = $policy.Name
                "Retention Policy Tag Links" = $policy.RetentionPolicyTagLinks;
                "Tenant" = $exdomain;
                "Tenant Office Name" = $loc;
            }

            $policyDetails = [pscustomobject]$Props
            $global:_retentionPolicies += $policyDetails

        }

    # Get Retention Tags for the tenant
        Write-Host "Obtaining retention tags for $exdomain Exchange Online tenant." -ForegroundColor Yellow

        $retentionTagsArray = Get-RetentionPolicyTag | Select Name,Type,AgeLimitForRetention,RetentionAction

        foreach($tag in $retentionTagsArray){
            
            $Props = [Ordered]@{
                "Name" = $tag.Name;
                "Type" = $tag.Type;
                "Age Limit For Retention" = $tag.AgeLimitForRetention;
                "Retention Action" = $tag.RetentionAction;
                "Tenant" = $exdomain;
                "Tenant Office Name" = $loc;
            }

            $tagDetails = [pscustomobject]$Props
            $global:_retentionTags += $tagDetails

        }

    Get-PSSession | Remove-PSSession

}

#Consolidate CSVs
$retPolPath = "D:\reports\DDBWW_Retention_Policies.csv"
$retTagPath = "D:\reports\DDBWW_Retention_Tags.csv"

$_retentionPolicies | Export-Csv -Path $retPolPath -NoTypeInformation
$_retentionTags | Export-Csv -Path $retTagPath -NoTypeInformation

$outfile = "D:\reports\DDBWW_Retention_Report_$date.xlsx"

ConvertAppendCSVXLSX -inputfile @($retPolPath, $retTagPath) -outfile $outfile

Start-Sleep -Seconds 20

# Prepare to email the report.
$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <noreply@PSService.local>"
$to= "wwpowershellreporting@ddb.com"
$subject= "DDB Worldwide Retention Report"
$body= "
    <p style='text-indent:1em;'>The Office 365 retention report for DDB Worldwide has been completed and is attached to this email.
    </p>"

# Send the mail message
Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
    -Attachments $outfile
