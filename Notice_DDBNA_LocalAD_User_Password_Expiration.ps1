# This script discoveres when users' passwords will expire, and will send them a notice.
# Script to Automated Email Reminders when Users Passwords due to Expire.
#


Import-Module ActiveDirectory

# Please Configure the following variables....


$SboxDomains = Import-Csv "D:\required_files\SBOXDomains.csv"

$credential= Import-Csv -Path "D:\required_files\SBOX.csv"

$un= $credential.username

$pw= $credential.key | ConvertTo-SecureString

$smtpServer="100.117.4.120"

$expireindays = 15

$from = "Network Password Expiration Notification <Password_Expiration_Notice@ddb.com>"

$logging = "Enabled" # Set to Disabled to Disable Logging

$logFile = "C:\mylog.csv" # ie. c:\mylog.csv

$testing = "Disabled" # Set to Disabled to Email Users

$testRecipient = "daniel.frank@ddb.com"

$date = Get-Date -format ddMMyyyy

# Check Logging Settings

if (($logging) -eq "Enabled"){
   
    # Test Log File Path

    $logfilePath = (Test-Path $logFile)

    if (($logFilePath) -ne "True"){

        # Create CSV File and Headers

        New-Item $logfile -ItemType File

        Add-Content $logfile "Name,EmailAddress,City,Domain,DaystoExpire,ExpiresOn,ObjectSid"

    }

}


Foreach($TargetDomain in $SboxDomains){

    $Domain = $TargetDomain.Domain

    $credun="$($Domain)\$($un)"
    
    $DomainCredentials = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $credun, $pw
    
   
    # Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired

    Import-Module ActiveDirectory

    $users = get-aduser -filter * -Server $Targetdomain.ServerIP -Credential $DomainCredentials -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress, objectSid, City |where {$_.Enabled -eq "True"} | where { $_.PasswordNeverExpires -eq $false } | where { $_.passwordexpired -eq $false }

    $DefaultmaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge

    # Process Each User for Password Expiry

    foreach ($user in $users)

    {

        $Name = $user.givenName

	    $Surname = $user.Surname

        $emailaddress = $user.emailaddress

        $city = $user.City

        $passwordSetDate = $user.PasswordLastSet

        $userSid = $user.objectSid

        $PasswordPol = (Get-AduserResultantPasswordPolicy $user)

        # Check for Fine Grained Password

        if (($PasswordPol) -ne $null)
        {
            $maxPasswordAge = ($PasswordPol).MaxPasswordAge
        }
        else
        {
            # No FGP set to Domain Default
            $maxPasswordAge = $DefaultmaxPasswordAge
        }  

        $expireson = $passwordsetdate + $maxPasswordAge

        $today = (get-date)

        $daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days

        # Set Greeting based on Number of Days to Expiry.

        # Check Number of Days to Expiry

        $messageDays = $daystoexpire

        if (($messageDays) -gt "1")
        {
            $messageDays = "in " + "$daystoexpire" + " days"
        }
        elseif(($messageDays) -eq "1")
        {
            $messageDays = "in " + "$daystoexpire" + " day"
        }
        else
        {
            $messageDays = "TODAY"
        }

        # Email Subject Set Here

        $subject="Your NETWORK password will expire $messageDays ($($domain))."

        # Email Body Set Here, Note You can use HTML, including Images.

           $body ="
    <p>Dear $name $Surname,</p>
    <p>&nbsp;</p>
    <p>Your Network/Paige password will expire $messageDays.</p>
    <p>Please change your Network/Paige password to avoid being restricted from the network, your computer, and other services.</p>
    <p>**If you are offsite, you must VPN first**</p>
    <p>Please go to: <a href='http://omcpaige.com'>http://omcpaige.com</a></p>
    <ul>
      <li>Enter your work email address and then click Next</li>
    </ul>
    <ul>
      <li>Enter your NETWORK/PAIGE password and then click Sign In</li>
    </ul>
    <ul>
      <li>Select your name in the upper right-hand corner and click on Settings</li>
    </ul>
    <ul>
      <li>Go to Change Password and enter your current and new password</li>
    </ul>
    <p>&nbsp;</p>
    <p>*Password requirements*<br>
      <ul><li>at least 14 characters</li>
      <li>At least 3 of the following 4 types of characters:
      <ul><li>lowercase letter</li></ul>
      <ul><li>uppercase letter</li></ul>
      <ul><li>number</li></ul>
      <ul><li>symbol</li></ul>
      </li>
      <li>no parts of your username.</li>
      <li>At least 1 day must have elapsed since you last password change</li></ul></p>
    <p>&nbsp;</p>
    <p>Still need help, please contact Paige.<br>
      <strong>Paige Portal</strong>:&nbsp;&nbsp;<a href='http://www.omcpaige.com/'>www.omcpaige.com</a> <br>
      <strong>Paige Voice</strong>:&nbsp; 888-MyPaige or 888-697-2443<br>
      <strong>Paige Chat</strong>:&nbsp;&nbsp;<a href='http://ddb.link/paigechat'>ddb.link/paigechat</a> <br>
    "
    
        # If Testing Is Enabled - Email Administrator

        if (($testing) -eq "Enabled")
        {
            $emailaddress = $testRecipient
        }

        # If a user has no email address listed

        if (($emailaddress) -eq $null)
        {
            $emailaddress = $testRecipient
        }

        # Send Email Message

        if (($daystoexpire -ge "0") -and ($daystoexpire -lt $expireindays))
        {
             # If Logging is Enabled Log Details

            if (($logging) -eq "Enabled" -and $testing -eq "Enabled")
            {
                $currentLog = Import-Csv -Path "C:\mylog.csv"

                $unique = $true

                foreach ($line in $currentLog)
                {
                    if($userSid -eq $line.objectSid)
                    {
                        $unique = $false

                    }
                    else
                    {
                        Add-Content $logfile "$Name,$emailaddress,$city,$Domain,$daystoExpire,$expireson,$userSid"
                    }
                }
            }

            # Send Email Message

            Send-Mailmessage -smtpServer $smtpServer -from $from -to $emailaddress -subject $subject -body $body -bodyasHTML -priority High  

        }

    } # End User Processing

    Get-PSSession | Remove-PSSession
}