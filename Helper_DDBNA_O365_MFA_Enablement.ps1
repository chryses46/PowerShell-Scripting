# This script enables O365 accounts that have recently been created.
# Note, only licensed users can be MFA enabled.
# Accounts created more than 23 days ago (to script run date) will be MFA'd the following week (run time)
# Accounts less than the 23 day mark will be ignored and added to a list that is sent to admins.

  
# Get tenant credentials for 365 connection
$list = Import-CSV -path "D:\required_files\DDBNA_MBAudit.csv"

# Get today's date
$today = (Get-date)

# Array for data to be sent to admins
$_dataToSend = @()

# Array to contain users whom will be enabled on the next cycle.
$_toEnable = @()

# Get the list of users that are to be enabled this run.
$_enablingToday = Import-CSV -Path "D:\reports\DDBNA_MFA_To_Enable.csv"


# Global Mail server settings
$emailSmtpServer = "100.117.4.120"
$emailToTest = "daniel.frank@ddb.com"
$emailFromAddress = "DDB IT <noreply@ddb.com>"
$emailccAddresses = "ITofficeleads@ddb.com","ADadmins@ddb.com"

# Iterate through each office
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

    if($_enablingToday){
        ####################################################################################################
        # First Enable last week's users
        ####################################################################################################

        # Enable MFA for a licensed user
        Function Enable-MFA($userUPN){

Write-Host "Enabling MFA on $userUPN."
      
            $St = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement

            $St.RelyingParty = "*"

            $st.State = "Enabled"
    
            $Sta = @($st)

            Set-MsolUser -UserPrincipalName $userUPN -StrongAuthenticationRequirements $sta

        }

        # Iterate through users and enable MFA where appropriate
        for ($i = 0; $i -lt $_enablingToday.Count; $i++){

            $user = $_enablingToday[$i];
        
            $userTenant = $user.Tenant;

            if($userTenant -eq $exdomain){

                $userUPN = $user.UserPrincipalName;

                Enable-MFA $userUPN;
            }
        }
        ####################################################################################################
        # End enablement of last week's users
        ####################################################################################################
    }

# Array for to-be enabled next week users
    $_enablingNextWeek = @()

# Array for less than 23 day users
    $_lessThan23Days = @()

# Get the current list of MFA exclusions in the tenant
    $mfaExcludeList = @()

    $mfaExcludeID = (Get-MsolGroup -all | Where-Object {$_.displayname -eq "MFA Service Account Exclusion"}).ObjectID

    Get-MsolGroupMember -GroupObjectId $mfaExcludeID | ForEach-Object {

        $mfaGroupUserProps = [Ordered]@{
            
                Email= $_.emailaddress
                Tenant= $exdomain
            }

                $mfaExcludedUser = [pscustomobject]$mfaGroupUserProps
                $mfaExcludeList += $mfaExcludedUser
        }

####################
####################
# Check if a user is set to be excluded
    Function Check-MfaExclusionStatus($userUPN){
        
        $excluded = $null
        
        for($i = 0; $i -lt $mfaExcludeList.Count; $i++){
            
            $exclusionToMatch = $mfaExcludeList[$i].Email;
            
            if($userUPN -eq $exclusionToMatch){
                
                $excluded = $true;

            }
        }

        return $excluded;

    }
####################
####################

####################
####################
# Gets the user's primary SMTP based on proxy addresses
    Function Get-UserSMTP($userUPN){
        
        $userSMTP = $null;

        $userProxies = (Get-MsolUser -UserPrincipalName $userUPN).ProxyAddresses

        for ($i = 0; $i -lt $userProxies.Count; $i++){
            $proxy = $userProxies[$i];

            if($proxy -clike "SMTP:*"){
                $userSMTP = $proxy -creplace 'SMTP:',''
            }
        }

        return $userSMTP
    }
####################
####################

####################
####################
# Check the current MFA status of the user
    Function Check-CurrentMfaStatus($userUPN){
        
        $userMfaStatus = (Get-MsolUser -UserPrincipalName $userUPN | Select StrongAuthenticationRequirements).StrongAuthenticationRequirements.State;
        
        if($userMfaStatus -ne "Enabled" -and $userMfaStatus -ne "Enforced"){$userMfaStatus = "Disabled"}

        return $userMfaStatus;

    }
####################
####################

####################
####################
# Check the creation date. 
    Function Determine-Enablement($userUPN, $userSMTP){
        
        $userCreationDate = (Get-Msoluser -UserPrincipalName $userUPN).WhenCreated

        if($userCreationDate -le $today.AddDays(-23)){
                    
            Write-Host "$userUPN was created on $($userCreationDate.ToShortDateString()) and is more than 23 days old. MFA will enable next week." -ForegroundColor Yellow
            
            #Add to enablingToday array
 Write-Host "Adding $userUPN to EnablingNextWeek Array" -ForegroundColor Yellow
            $userProps = [Ordered]@{
                UPN = $userUPN;
                Email = $userSMTP;
                CreationDate = $userCreationDate;

            }
            
            $userData = [pscustomobject]$userProps

            $global:_enablingNextWeek += $userData

            #Email the user they will be enabled within the week
            Email-OneWeekToMFA $userSMTP

        }else{

            Write-Host "$userUPN was created on $($userCreationDate.ToShortDateString()) and is less than 23 days old. No MFA will occur."

            #Add to lessThan23days array
Write-Host "Adding $userUPN to lessThan23Days Array" -ForegroundColor Yellow
            $userProps = [Ordered]@{
                UPN = $userUPN;
                Email = $userSMTP;
                CreationDate = $userCreationDate;
            }

            $userData = [pscustomobject]$userProps

            $global:_lessThan23Days += $userData
        }

    }
####################
####################

####################
####################
# Send an email if MFA will be enabled within the week
    Function Email-OneWeekToMFA($userSMTP){
Write-Host "Sending One Week To MFA email to $userSMTP."
        $to= $userSMTP
        $subject= "MFA : Securing your email using Multi-factor Authentication (MFA)"
        $body= "
            <p>Hello,</p>

            <p>DDB IT wanted to make you aware of upcoming changes to your email account to help combat the recent increase in phishing attempts and to fulfill our security commitments to our clients. If we haven't already implemented Microsoft MFA (Multi-Factor Authentication) on your account, we will enable it within one week.</p>
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

####################
####################
# Construct admin CSV/XLSX file
    Function GatherMfaUserNotice(){
Write-Host "Gathering users for Admin notice..." -ForegroundColor Green   
        for($i = 0; $i -lt $global:_enablingNextWeek.Count; $i ++){
           
           $user = $global:_enablingNextWeek[$i];

Write-Host "Processing $user in EnablingNextWeek array..." -ForegroundColor Yellow
            
            $userProps = [Ordered]@{

                "UserPrincipalName" = $user.UPN;
                "Email" = $user.Email;
                "Tenant" = $exdomain;
                "Days To MFA" = (New-TimeSpan -Start $today.AddDays(-30) -End $user.CreationDate).Days;
                "Enabling Next Week" = "True";
            }

            $userDetails = [pscustomobject]$userProps

            $global:_dataToSend += $userDetails

            # Add to DDBNA_MFA_To_Enable.csv for next week's run
            $global:_toEnable += $userDetails

        }

        for($i = 0; $i -lt $global:_lessThan23Days.Count; $i ++){

            $user = $global:_lessThan23Days[$i]

Write-Host "Processing $($user.UPN) in lessThan23Days array..." -ForegroundColor Yellow

            $userProps = [Ordered]@{

                "UserPrincipalName" = $user.UPN;
                "Email" = $user.Email;
                "Tenant" = $exdomain;
                "Days To MFA" = (New-TimeSpan -Start $today.AddDays(-30) -End $user.CreationDate).Days;
                "Enabling Next Week" = "False"
             }

            $userDetails = [pscustomobject]$userProps

            $global:_dataToSend += $userDetails

        }


    }
####################
####################

# Get licensed users in the tenant
    $users = Get-MsolUser -All | Where {$_.isLicensed -eq "True"} | Select UserPrincipalName

####################
####################
# Main loop - iterating through users list
    for($i = 0; $i -lt $Users.Count; $i++){
        $user = $Users[$i];
Write-Host "----------------------------------------" -ForegroundColor Yellow
Write-Host "Working on $($user.UserPrincipalName)..." -ForegroundColor Green        
        # Set defaults
        $isExcluded = $null

        # Set current user fields
        $userUPN = $user.UserPrincipalName;

        # Get user primarySMTP
        $userSMTP = Get-UserSMTP $userUPN

        # Check if the user is MFA excluded
        $isExcluded = Check-MfaExclusionStatus $userUPN
        
        # If the user is not excluded
        if(!$isExcluded){
            
            #Get the current MFA status of the user
            $mfaStatus = Check-CurrentMfaStatus $userUPN

            if($mfaStatus -eq "Disabled"){
            
                #Determine creation date and if to enable or not, take action based on results
                
                Determine-Enablement $userUPN $userSMTP

            }
        }
    }
####################
####################

    GatherMfaUserNotice;

    #End the current 365 Powershell session
    Get-PSSession | Remove-PSSession

}

$date= get-date -Format MMddyy

if($_enablingToday){
# Clear/ delete the old CSV file for enabled users
del "D:\reports\DDBNA_MFA_To_Enable.csv"
}
$_toEnable | Export-CSV -Path "D:\reports\DDBNA_MFA_To_Enable.csv" -NoTypeInformation

$usersToMfaCsvPath = "D:\reports\DDBNA_Users_To_MFA_$date.csv"

$_dataToSend | Export-Csv -Path $usersToMfaCsvPath -NoTypeInformation

$usersToMfaReportPath = Convert-CsvToXlsx($usersToMfaCsvPath)

Start-Sleep -Seconds 5
        

#Send summery emails to ITOfficeLeads@ddb.com, etc
$subject= "DDB North America MFA Users To Be Enabled"
$body= "
    <h2>What is this?</h2>
    <p>The attached Excel file details users that will be enabled (or not) for MFA.</p>
    <strong>Note:</strong></p>The users on this particular email have not been modified. This email is what you will expect to see.</p>
    <h3>Considerations:</h3>
    <ol>
        <li>Users older than or equal to 23 days from creation will be enabled next week.</li>
        <li>Users who are less than 23 days old will show the remaining days until they are MFA'd</li>
    </ol>
    "

#Send the message
Send-MailMessage -To $emailccAddresses -From $emailFromAddress -Subject $subject -Body $body -BodyAsHtml -SmtpServer $emailSmtpServer -Priority High `
            -Attachments $usersToMfaReportPath