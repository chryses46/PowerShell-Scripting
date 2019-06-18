<# 

This script works with a securely-stored CSV file that lists locations' EOL Admin and secure password. 
The passwords are created previously by the PSService log in, to ensure proper functioning of the Scheduled Task "DDB World Wide Office 365 Mailbox Auditing, Retention and Licensing Report"

This script will enable auditing on mailboxes that are not already enabled in a target Office 365 tenant.

#>


# Import the list of user names and secured passwords
$list = Import-CSV -path "D:\required_files\DDBNA_MBAudit.csv"

$date= get-date -Format MM/dd/yy

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
   
    $Users = Get-Mailbox -Filter {AuditEnabled -eq $false} | select Distinguishedname

    if($Users -eq $null){
        Write-Host "All users in the $exdomain Exchange Online tenant are enabled for auditing."
    
        } else { write-host "Users have been found in the $exdomain Exchange Online teannt with auditing disabled. Enabling now."}

    #Set users that need it, for Mailbox Auditing.
    foreach($user in $users){
    
        Write-Host "Enabling mailbox auditing on $($user.Distinguishedname)."

        Set-Mailbox `
                $user.Distinguishedname `
                -AuditEnabled $true `
                -AuditLogAgeLimit 365 `
                -AuditOwner Create,HardDelete,MailboxLogin,MoveToDeletedItems,SoftDelete,Update
    }

Get-PSSession | Remove-PSSession

}

$_officeLocation = @()

Function OfficeLocations(){
    
    foreach($loc in $list){

        $global:_officeLocations += "<li>$($loc.Office)</li>"
    } 
}

OfficeLocations

# Prepare to email the reports.
$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <noreply@PSService.local>"
#DDBNA_O365Admins@ddb.com
$to= "DDBNA_O365Admins@ddb.com"
$subject= "DDBNA O365 Mailbox Auditing Complete ($($date))"
$body= "
    <h1>Mailbox Auditing Complete</h1>
    <p>Recently created mailboxes in DDB North America Office 365 tenants have been enabled for auditing. The following tenants were affected:</p>
    <br/>
    <ul>
        $_officeLocations
    </ul>
    "

# Send the mail message
Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High
