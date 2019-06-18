#Create a weekly scheduled task that runs at 10:15AM and execute this script to generate password expiration summaries sent out to HD Ticket Systems
Import-Module ActiveDirectory

#----- Configurable Settings ------------------------------------------------------------------------------------------------------------------------
# Interval of days until expire to send emails
$warningIntervals = 7,6,5,4,3,2,1,0

#----- End Configurable Settings -------------------------------------------------------------------------------------------------------------------


#Constant used to determine if password never expires flag is set
$ADS_UF_DONT_EXPIRE_PASSWD = 0x00010000

#Constant used to determine if user must change password at next login
$REQUIRED_PASSWORD_CHANGE_LASTSET = 0

$CurrentDate = [datetime]::Now.Date
# DEPRECATED
#$adminEmailContent = ""

$dateToday = get-date -Format MMddyy


$DefaultDomainPasswordPoliy = Get-ADDefaultDomainPasswordPolicy
 
$smtp = new-object Net.Mail.SmtpClient($smtpServer)

$SboxDomains = Import-Csv -Path "D:\required_files\SBOXDomains.csv"

$credential= Import-Csv -Path "D:\required_files\SBOX.csv"

$un= $credential.username

$pw= $credential.key | ConvertTo-SecureString

$passwordsExpiring = @()

function GetDaysToExpire([datetime] $expireDate)
{
    $date =  New-TimeSpan $CurrentDate $expireDate
    return $date.Days
}

function GetPasswordExpireDate($user)
{
    return $pswdExpireDate = [datetime]::FromFileTimeUTC($user.pwdLastSet+$DefaultDomainPasswordPoliy.MaxPasswordAge.Ticks)
}

function IsInWarningIntervals([int] $daysToExpire)
{
    
    foreach( $interval in $warningIntervals )
    {
       if ( $daysToExpire -eq $interval )
       {
        return $True
       }
    }
    
    return $False
}


Foreach($TargetDomain in $SboxDomains)
{

    $Domain = $TargetDomain.Domain

    $credun="$($Domain)\$($un)"
    
    $DomainCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credun, $pw
    
    #get all users
    $users = Get-AdUser -Filter * -Server $TargetDomain.ServerIP -Credential $DomainCredentials -Properties distinguishedName, userAccountControl, pwdLastSet, userprincipalname, mail, DisplayName, samAccountName, accountExpires, enabled, city, company, title | Where-Object {$_.DistinguishedName -notlike "*,OU=HPE,*"} 

    foreach( $user in $users ) 
    {  
        #if account is enabled and password never expire flag does not exist, then process user
        if ( ($user.enabled -eq $True) -and (($user.userAccountControl -band $ADS_UF_DONT_EXPIRE_PASSWD) -eq 0) ) 
        {        
            $pwdExpires = GetPasswordExpireDate $user 
            $daysToExpire = GetDaysToExpire $pwdExpires 
        
        
            #if day falls on warning interval
            if( IsInWarningIntervals $daysToExpire )
            {
                $UserValues = [Ordered]@{
                    Name= $user.Name;
                    "Password Expired" = "False";
                    "Password Expiration Date" = $pwdExpires;
                    Email = $user.mail;
                    City = $user.City;
                    Company = $user.company;
                    Title = $user.Title;
                    Domain = $Domain;
                    DN = $user.distinguishedName
            
                }

                $UserEntry = [pscustomobject]$UserValues

                $passwordsExpiring +=  $UserEntry
            
            }
        }
        
        #if days to expire is negative, password has expired.  add to admin email
        if( $daysToExpire -lt 0 )
        {
            $UserValues = [Ordered]@{
                Name= $user.Name;
                "Password Expired" = "True";
                "Password Expiration Date" = $pwdExpires;
                Email = $user.mail;
                City = $user.City;
                Company = $user.company;
                Title = $user.Title;
                Domain = $Domain;
                DN = $user.distinguishedName
            
            }

            $UserEntry = [pscustomobject]$UserValues

            $passwordsExpiring +=  $UserEntry
        }
    }
    
    Get-PSSession | Remove-PSSession
}

$csvFilePath = "D:\reports\DDB NA Password Exipation Summary_$dateToday.csv"

$passwordsExpiring | Export-Csv -Path $csvFilePath -NoTypeInformation

$xlsxFilePath = Convert-CsvToXlsx($csvFilePath);

Start-Sleep -Seconds 15

$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
#ADAdmins@ddb.com
$to= "ADAdmins@ddb.com"
$subject= "DDB NA Password Expiration Summary"
$body= "The DDBNA Password Expiration Summary report is complete."

# Send the mail message #ADAdmins@ddb.com

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
    -Attachments $xlsxFilePath

## REFUSE BIN
    
<# DISABLING EMAIL TO USERS AND SINGLE-LINE EMAIL TO ADMIN
function EmailUser($user)
{
    $pwdExpires = GetPasswordExpireDate $user
    $daysToExpire = GetDaysToExpire $pwdExpires 
    
    if( $daysToExpire -eq 0 )
    {
        $emailContent = $user.DisplayName + ", your password will expire today at " + $pwdExpires.ToShortTimeString() + " GMT"
    }
    else
    {
        $emailContent = $user.DisplayName + ", your password will expire in " + $daysToExpire + " days."
    }
}
function AppendAdminEmailMail($user)
{
$pwdExpires = GetPasswordExpireDate $user
$daysToExpire = GetDaysToExpire $pwdExpires 

return "The password for account """ + $user.samAccountName + """ will expire in " + $daysToExpire + " days at " + $pwdExpires + [System.Environment]::NewLine + [System.Environment]::NewLine
}
function EmailAdmin($content) 
{
    if( [string]::IsNullOrEmpty($content) -and $alwaysSendAdminSummary )
    {
        $content = "There are no user's with expired passwords or users that need to change their password."
    }
    if( [string]::IsNullOrEmpty($content) -ne $True )
    {
        $smtp.Send($fromEmail, $adminEmail, "Password Expiration Summary", $content)
    }
}

function AppendAdminEmailNoMail($user)
{
    $pwdExpires = GetPasswordExpireDate $user
    $daysToExpire = GetDaysToExpire $pwdExpires 
   
    return "The password for account """ + $user.samAccountName + """ will expire in " + $daysToExpire + " days at " + $pwdExpires + " and there is no associated email address to send a notification to" + [System.Environment]::NewLine + [System.Environment]::NewLine
}

function AppendAdminEmailExpiredAccount($user)
{
    $pwdExpires = GetPasswordExpireDate $user
    $daysToExpire = GetDaysToExpire $pwdExpires 
   
    if( $user.pwdLastSet -eq 0 )
    {
        return "The user account """ + $user.samAccountName + """ is set to require a password change at next logon and the user has not yet changed it" + [System.Environment]::NewLine + [System.Environment]::NewLine
    }
    else
    {
        return "The password for user account """ + $user.samAccountName + """ expired " + $daysToExpire + " days ago on " + $pwdExpires + [System.Environment]::NewLine + [System.Environment]::NewLine
    }
}
#>