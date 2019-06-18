    #Get Licensed users and associated lincenses in the tenant

    Import-Module ConvertAppendCSVXLSX

    $list = Import-CSV -path "D:\required_files\DDBNA_MBAudit.csv"

    $LicenseUserInfo = @()
    $_globalAdmins = @()

    #Iterate through each office in the csv file
    $list | ForEach-Object{

    #Set Office's Log In Information
    $un = $_.Username;
    $pw = $_.Key | ConvertTo-SecureString
    $loc= $_.Office;
    
    Write-Host "Attempting to connect to $loc Office 365 Portal. Please wait..." -ForegroundColor Yellow

        #Convert password and log into MSOL.
        $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $un, $pw
        Connect-MsolService -Credential $Creds

        #Create EXOPSSession
        $EOLConn = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
        Import-PSSession $EOLConn

        #Get organizational information
        $domain= (Get-MsolDomain | Where-Object {$_.isDefault -eq $true}).Name
        $exdomain= (Get-OrganizationConfig).Name

    if($domain -ne $null -and $exdomain -ne $null){
        Write-Host "Connected. 365 AD Domain = $domain. Tenant Domain Name = $exdomain." -ForegroundColor Green
    }else{
        Write-Host "Connection to $loc Office 365 failed. Exiting." -ForegroundColor Red
        Get-PSSession | Remove-PSSession
        break;
    }
        
    #Get the current list of MFA exclusions in the tenant
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

    #Get all mailboxes in the tenant
    $Mailboxes = Get-Mailbox -Filter * -ResultSize Unlimited | Select `
                     PrimarySmtpAddress,
                     UserPrincipalName,
                     DeliverToMailboxAndForward,
                     ForwardingSmtpAddress,
                     ForwardingAddress,
                     RetentionPolicy,
                     RetentionHoldEnabled,
                     LitigationHoldEnabled,
                     LitigationHoldOwner,
                     RecipientType,
                     isDirSynced
        

    #Iterate through each MailBox and get MailBox data

        for($i = 0; $i -lt $Mailboxes.Count; $i++){
            
            $mb = $Mailboxes[$i]
            
            #Reset variables in loop for each user.
            $userLic=$null
            $userInfo=$null
            $mfaExcludeStatus=$null

            #Set Exchange values for current user.
            $mbSmtp = $mb.PrimarySmtpAddress;
            $mbDeliverAndForward = $mb.DeliverToMailboxAndForward;
            $mbFwdSmtp = $mb.ForwardingSmtpAddress;
            $mbFwd = $mb.ForwardingAddress;
            $mbRetentionPolicy = $mb.RetentionPolicy;
            $mbRetentionHold = $mb.RetentionHoldEnabled;
            $mbLitigationHold = $mb.LitigationHoldEnabled;
            $mbLitigationHoldOwner = $mb.LitigationHoldOwner;
            $mbRecipientType = $mb.RecipientType;
            $mbUpn = $mb.UserPrincipalName;
            if($mb.isDirSynced -eq $true){$mbCloud = "False"}else{$mbCloud= "True"}

            #Determine if the user is MFA Excluded
            foreach($mfaExcludedUser in $mfaExcludeList){$exclusionToMatch = $mfaExcludedUser.Email; if($mbSmtp -eq $exclusionToMatch){$mfaExcludeStatus = "Excluded"}}

            #Function to be called if $mbUpn is found
            function Get-365UserData([string] $mbUpn){
                
                $userInfo=(Get-MsolUser -UserPrincipalName $mbUpn | Select `
                    FirstName,
                    LastName,
                    DisplayName,
                    Office,
                    Department,
                    Licenses,
                    StrongAuthenticationRequirements,
                    BlockCredential,
                    ProxyAddresses,
                    UserPrincipalname,
                    WhenCreated,
                    PhoneNumber)
            
                #Information related to User's MSOL User Object.
                $userDisplayName = $userInfo.DisplayName;
                $userFn = $userInfo.FirstName;
                $userLn = $userInfo.LastName;
                $userDepartment = $userInfo.Department;
                $userOffice = $userInfo.Office;
                $userSignInBlocked = $userInfo.BlockCredential;
                $userMfaState = $userInfo.StrongAuthenticationRequirements.State;
                $userUPN = $userInfo.UserPrincipalName;
                $userCreated = $userInfo.WhenCreated;
                $userPhone = $userInfo.PhoneNumber;

                 #Determine if the user is MFA Excluded (Non-MailBox Accounts)
                foreach($mfaExcludedUser in $mfaExcludeList){$exclusionToMatch = $mfaExcludedUser.Email; if($userUPN -eq $exclusionToMatch){$userMfaExcluded = "Excluded"}}

                #Determine Licensing
                $userLic=[string]$userInfo.Licenses.AccountSku.SkuPartNumber

                if($userLic -eq $null){
                    $userLic = "None"
                }elseif($userLic -eq "ENTERPRISEPACK"){
                    $userLic= "E3"
                }elseif($userLic -eq "EXCHANGEENTERPRISE"){
                    $userLic= "P2"
                }elseif($userLic -eq "EXCHANGESTANDARD"){
                    $userLic= "P1"
                }

                $userInfoProps = [Ordered]@{
                    "Display Name" = $userDisplayName;
                    "First Name" = $userFn;
                    "Last Name" = $userLn;
                    "License Type" = $userLic;
                    Office = $userOffice;
                    Department = $userDepartment;
                    "User Disabled" = $userSignInBlocked;
                    "MFA Status" = $userMfaState;
                    "MFA Excluded" = $userMfaExcluded;
                    "When Created" = $userCreated;
                    "Phone Number" = $userPhone;
                }

                $userData = [pscustomobject]$userInfoProps

                return $userData
            }
            #End of Get-365UserData function

            if($mbUpn -ne $null){
                $userData = Get-365UserData($mbupn)
                }else{
                $userData = $null;
                }

            #Final determination of MFA Exclude
            if($mfaExcludeStatus -eq $null){if($userData.'MFA Excluded' -ne $null){$mfaExcludeStatus = $userData.'MFA Excluded'}else{$mfaExcludeStatus = $null}}
           
            $Props = [Ordered]@{
                "Display Name" = $userdata.'Display Name';
                "First Name" = $userdata.'First Name';
                "Last Name" = $userdata.'Last Name';
                "User Principal Name" = $mbUpn;
                "Email Address" = $mbSmtp;
                "License Type" = $userdata.'License Type';
                Office = $userdata.'Office';
                Phone = $userdata.'Phone Number';
                Department = $userdata.'Department';
                "User Disabled" = $userdata.'User Disabled';
                "MFA Status" = $userdata.'MFA Status';
                "When Created" = $userdata.'When Created';
                "MFA Excluded" = $mfaExcludeStatus;
                "User Forward SMTP"= $mbFwdSmtp;
                "User Internal Forward" = $mbFwd;
                "User SMTP Deliver and Forward" = $mbDeliverAndForward;
                "Retention Policy" = $mbRetentionPolicy;
                "Retention Hold Enabled" = $mbRetentionHold;
                "Litigation Hold Enabled" = $mbLitigationHold;
                "Litigation Hold Owner" = $mbLitigationHoldOwner;
                Tenant = $exdomain;
                "Tenant Office Name" = $loc;
                "Office 365 Primary Domain" = $domain;                
                "Cloud Only Account" = $mbCloud;
                "Mailbox Type" = $mbRecipientType;
            }

            $obj = [pscustomobject]$Props
            $LicenseUserInfo += $obj
        }

    # Collect Global Admin users

        $GArole = Get-MsolRole -RoleName "Company Administrator"
    
        $gaUsers = Get-MsolRoleMember -RoleObjectId $GArole.ObjectId

        foreach($ga in $gaUsers){

            $Props = [Ordered]@{
                "Display Name" = $ga.DisplayName;
                "Email" = $ga.EmailAddress;
                "Is Licensed" = $ga.isLicensed;
                "Tenant" = $exdomain;
                "Tenant Office Name" = $loc;
            }

            $gaDetails = [pscustomobject]$Props
            $global:_globalAdmins += $gaDetails

        }

    Get-PSSession | Remove-Pssession

}

#Export requested information

$date= get-date -Format MMddyy
    
$mbCompleteFilePath = "D:\reports\DDBNA_O365_Complete.csv"
$gaReportPath = "D:\reports\DDBNA_Global_Admins.csv"

$LicenseUserInfo | Export-Csv -Path $mbCompleteFilePath -NoTypeInformation
$_globalAdmins | Export-Csv -Path $gaReportPath -NoTypeInformation

$outfile = "D:\reports\DDBNA_O365_Complete_$($date).xlsx"

ConvertAppendCSVXLSX -inputfile @($gaReportPath, $mbCompleteFilePath) -outfile $outfile

Start-Sleep -Seconds 30

$_officeLocations = @()

Function OfficeLocations(){
    
    foreach($loc in $list){

        $global:_officeLocations += "<li>$($loc.Office)</li>"
    } 
}

OfficeLocations

$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
#DDBNA_O365Admins@ddb.com
$to= "DDBNA_O365Admins@ddb.com"
$subject= "DDBNA Complete Office 365 Report"
$body= "
    <h1>DDB North America Complete Office 365 Report</h1>
    <p>Attached you will find an Excel report that contains MFA status, licensing, mailbox forwarding rules and more for DDB North American Office 365 tenants including:</p>
    <br/>
    <ul>
        $_officeLocations
    </ul>
"
# Send the mail message

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
            -Attachments $outfile