# This script works with the Helper_DDBNA_O365_MFA_Enablement script to notify users that they will be MFA'd the next day
# If the above mentioned script is running on a Tuesday morning, this script should be set for Monday morning.

# Get tenant credentials for 365 connection
$list = Import-CSV -path "D:\required_files\DDBNA_MBAudit.csv"

# Get the csv created in above script
$_enablingTomorrow = Import-CSV -Path "D:\reports\DDBNA_MFA_To_Enable.csv"

# Global Mail server settings
$emailSmtpServer = "100.117.4.120"
$emailToTest = "daniel.frank@ddb.com"
$emailFromAddress = "DDB IT <noreply@ddb.com>"
$emailccAddresses = "ITofficeleads@ddb.com","ADadmins@ddb.com"

# Connect to each tenant
$list | ForEach-Object{

    # Set Office's Log In Information
    $un = $_.Username;
    $pw = $_.Key | ConvertTo-SecureString
    $loc= $_.Office;
    
    Write-Host "Attempting to connect to $loc Office 365 Portal. Please wait..." -ForegroundColor Green

        # Convert password and log into MSOL.
        $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $un, $pw
        Connect-MsolService -Credential $Creds

        # Create EXOPSSession
        $EOLConn = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
        Import-PSSession $EOLConn

        # Get organizational information
        $domain= (Get-MsolDomain | Where-Object {$_.isDefault -eq $true}).Name
        $exdomain= (Get-OrganizationConfig).Name

    if($domain -ne $null -and $exdomain -ne $null){
        Write-Host "Connected. 365 AD Domain = $domain. Tenant Domain Name = $exdomain." -ForegroundColor Green
    }else{
        Write-Host "Connection to $loc Office 365 failed. Exiting." -ForegroundColor Red
        Get-PSSession | Remove-PSSession
        break;
    }

####################
####################
    Function Email-OneDayToMFA($userSMTP){
Write-Host "Sending One Day To MFA email to $userSMTP."
        $to= $userSMTP
        $subject= "MFA : Securing your email using Multi-factor Authentication (MFA)"
        $body= "
            <p>Hello,</p>

            <p>DDB IT wanted to remind you of upcoming changes to your email account to help combat the recent increase in phishing attempts and to fulfill our security commitments to our clients. Tomorrow we will implement Microsoft MFA (Multi-Factor Authentication) on your account.</p>
            <h3>To get started, please download the following apps on your phone:</h3>
            <ol>
                <li>Outlook app (Apple Mail is not supported)</li>
                <li>Microsoft Authenticator app</li>
            </ol>
            <p><h3>Configure your MFA settings:</h3><a href=' https://aka.ms/MFAsetup'>https://aka.ms/MFAsetup</a> (Please reference the attached documents.)</p>
            <p><strong>Note</strong>: You will need to authenticate for each application. Outlook, Outlook mobile, Skype for business, etc. As well, you will need to re-authenticate every 30 days.</p>
            <h3>Need help, call Paige.</h3> 
            <ul>
                <li>Paige Portal: <a href='www.omcpaige.com'>www.omcpaige.com</a>(use your network password)</li>
                <li>Paige Voice: 888-MyPaige or 888-697-2443</li>
                <li>Paige Chat: <a href='ddb.link/paigechat'>ddb.link/paigechat</a></li>
            </ul>
            ---- 

            <h3>What is Phishing?</h3>
                <p>Phishing is the fraudulent practice of sending emails in which an attacker masquerades as a reputable person or company in order to distribute malicious links or attachments attempting to induce individuals to reveal personal information, such as passwords and credit card numbers.</p> 

            <h3>MFA?</h3>
                <p>Microsoft MFA (Multi-Factor Authentication) helps safeguard access to data and applications while meeting user demand for a simple sign-in process. It provides additional security by requiring a second form of authentication and delivers strong authentication via a range of easy verification options. Many other global companies such as Google, Apple and large banks have been using Multi-Factor Authentication for years.</p>
            
            <h3>MFA Enrollment requires registering a cell phone number that will be readily accessible to you.</h3>
                <p>Once MFA is configured, a six-digit code will be sent via text message to your cell phone whenever you need to authenticate your account on any device or computer. You will need to enter the six-digit code to complete the authentication process.</p> 

            <p>Thank you,</p>
            <p>DDB IT</p>
            "

        #Send the message
        Send-MailMessage -To $to -From $emailFromAddress -Subject $subject -Body $body -BodyAsHtml -SmtpServer $emailSmtpServer -Priority High

    }
####################
####################
    
    # Itirate through the list and send emails where appropriate
    for($i = 0; $i -lt $_enablingTomorrow.Count; $i++){

        $user = $_enablingTomorrow[$i]

        $userTenant = $user.Tenant;

        if($userTenant -eq $exdomain){

            $userSMTP = $user.Email;

            Email-OneDayToMFA $userSMTP;
        }
    }       

    Get-PSSession | Remove-PSSession
}

 
