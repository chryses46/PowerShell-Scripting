# This script is used to test Office 365 connections using credentilas given for DDB WorldWide offices.

$list = Import-CSV -path "D:\required_files\DDBWW_MBAudit_testing.csv"

$ErrorEmail = @()

$list | ForEach-Object{

$un = $_.Username;
$pw = $_.Key | ConvertTo-SecureString
$loc= $_.Office;

    Write-Host "Attempting to connect to $loc Office 365 Portal. Please wait..."

#Convert password and log into MSOL.

    $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $un, $pw
    Connect-MsolService -Credential $Creds

#Create EXOPSSession

    $EOLConn = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
    Import-PSSession $EOLConn

#Get org info

    $domain= (Get-MsolDomain | Where-Object {$_.isDefault -eq $true}).Name
    $exdomain= (Get-OrganizationConfig).Name

#Check PW Expire Status

    $pwdExpireStatus = Get-MsolUser -UserPrincipalName $un | Select PasswordNeverExpires

    if(!$pwdExpireStatus){
        Set-MsolUser -UserPrincipalName $un -PasswordNeverExpires $true;
    }
    else
    {
        Write-Host "$un's password is already set to never expire."
    }

    if($domain -eq $NULL -OR $exdomain -eq $NULL){
        $ErrorEmail += $loc
        
    }else{Write-Host "Connection to the $loc Office 365 tenant successful."}

Get-PSSession | Remove-PSSession

}


if($ErrorEmail -ne $NULL){

    #Prepare email for error reports
    $smtpserver= "100.117.4.120"
    $from= "PowerShell Server Reporting <noreply@PSService.local>"
    $to= "daniel.frank@ddb.com"
    $subject= "DDBWW Office 365 Connection Issues Detected!"
    $body= "
        <p style='text-indent:1em;'>Errors have been detected during the testing of DDBWW Office 365 connections from the PowerShell scritping server at 100.116.4.118. Please look into these immediately.
        <br/>
        Failed Connections:
        <br/>
        $(foreach($line in $erroremail){
        $line
        '<br/>'
        })
        </p>"

    # Send the mail message

    Send-MailMessage -SmtpServer $smtpserver -From $from -To $to -Subject $subject -Body $body -BodyAsHtml -Priority High

    }