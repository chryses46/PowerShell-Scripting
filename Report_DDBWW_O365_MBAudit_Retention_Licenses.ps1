<# 

This script works with a securely-stored CSV file that lists locations' EOL Admin and secure password. 
The passwords are created previously by the PSService log in, to ensure proper functioning of the Scheduled Task "DDB World Wide Office 365 Mailbox Auditing, Retention and Licensing Report"

This script will enable auditing on mailboxes that are not already enabled in a target Office 365 tenant.

#>


# Import the list of user names and secured passwords
$list = Import-CSV -path "D:\required_files\DDBWW_MBAudit.csv"

$date= get-date -Format MMddyy


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

    #Get Retention Policy report for the domain.  
        Write-Host "Obtaining Retention Policy Reports for $exdomain Exchange Online tenant."

        $RetentionPolicyCsvPath = "D:\reports\$($exdomain)RetentionPolicies_$date.csv"
        Get-RetentionPolicy | Select Name,RetentionPolicyTagLinks | Export-Csv -Path $RetentionPolicyCsvPath -NoTypeInformation
        #Convert it to XLSX
        $RetentionPolicyXlsxPath = Convert-CsvToXlsx($RetentionPolicyCsvPath);
        Start-Sleep -Seconds 15

        $RetentionTagReportCsvPath = "D:\reports\$($exdomain)RetentionTags_$date.csv"
        Get-RetentionPolicyTag | Select Name,Type,AgeLimitForRetention,RetentionAction | Export-Csv -Path $RetentionTagReportCsvPath -NoTypeInformation
        #Convert it to XLSX
        $RetentionTagReportXlsxPath = Convert-CsvToXlsx($RetentionTagReportCsvPath);
        Start-Sleep -Seconds 15

    # Creates a csv file of the users' Retention Policies assigned. 
        $ExchangeUsers = Get-Mailbox -Filter * | Select PrimarySmtpAddress,RetentionPolicy
        $UserRetentionReportCsvPath = "D:\reports\$($exdomain)UserRetentionPolicy_$date.csv"
        $ExchangeUsers | Export-Csv -Path $UserRetentionReportCsvPath -NoTypeInformation

        #Convert it to XLSX
        $UserRetentionReportXlsxPath = Convert-CsvToXlsx($UserRetentionReportCsvPath);
        Start-Sleep -Seconds 15


    #Get Licensed users and associated lincenses in the tenant
        Write-Host "Obtaining User License Report for the $exdomain Office 365 tenant."
    
        $LicenseUsers = Get-MsolUser -All | Where {$_.isLicensed -eq "True"} | Select UserPrincipalName,Licenses,StrongAuthenticationRequirements,Title,Company,Department,DisplayName,Office,Country,State,BlockCredential
        $LicenseUserInfo = @()

        #Iterate through users and add their MSOL/EOL data to the array
        foreach($user in $LicenseUsers){
            
            $exInfo = (Get-Mailbox -Identity $user.UserPrincipalName -ErrorAction SilentlyContinue)
            $boxstats = (Get-MailboxStatistics $exInfo.PrimarySMTPAddress)
            
            $lic=[string]$user.Licenses.AccountSku.SkuPartNumber

                if($exInfo.isDirSynced -eq $true){
                    $cloud = "Synced Account"
                    }else{$cloud= "Cloud-Only Account"}

                if($lic -eq "ENTERPRISEPACK"){
                    $lic= "E3"
                    }elseif($lic -eq "EXCHANGEENTERPRISE"){
                    $lic= "P2"
                    }elseif($lic -eq "EXCHANGESTANDARD"){
                    $lic= "P1"
                    }


            $Props = [Ordered]@{
                "Display Name" = $user.DisplayName;
                "User Principal Name" = $User.UserPrincipalName;
                "Email Name" = $exInfo.Alias;
                "Email Domain" = $exdomain;
                Office = $User.Office;
                Country = $user.Country;
                State = $user.state;
                "Recipient Type" = $boxstats.MailboxTypeDetail;
                Company = $user.company;
                Department = $user.Department;
                "Job Title" = $user.Title;
                "Tenant Domain" = $domain;
                "Account Disabled" = $user.blockcredential;
                "Last Logon Time" = $boxstats.LastLogonTime;
                "Mailbox Creation Date" = $exInfo.WhenMailboxCreated;
                License = $lic;
                MFA=$User.StrongAuthenticationRequirements.State;
	            DirSyncStatus=$cloud;
            }

            $obj = [pscustomobject]$Props
            $LicenseUserInfo = $LicenseUserInfo + $obj
        }

        #Export requested information
        $LicenseReportCsvPath = "D:\reports\$($exdomain)LicenseReport_$date.csv"
        $LicenseUserInfo | Export-Csv -Path $LicenseReportCsvPath -NoTypeInformation

        #Convert it to XLSX
        $LicenseReportXlsxPath = Convert-CsvToXlsx($LicenseReportCsvPath);
        Start-Sleep -Seconds 15


    # Collect Global Admin Report

    $GArole = Get-MsolRole -RoleName "Company Administrator"
    
    $GaReportCsvPath = "D:\reports\$($exdomain)GlobalAdminReport_$date.csv"
    Get-MsolRoleMember -RoleObjectId $GArole.ObjectId | Export-Csv -Path $GaReportCsvPath -NoTypeInformation

    #Convert it to XLSX
    $GaReportXlsxPath = Convert-CsvToXlsx($GaReportCsvPath);
    Start-Sleep -Seconds 15


# Prepare to email the reports.
$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <noreply@PSService.local>"
$to= "wwpowershellreporting@ddb.com"
$subject= "$($exdomain)'s Office 365 Audit, Retention, License and Global Admin Reports."
$body= "
    <p style='text-indent:1em;'>The Office 365 audit, retention, license and global admin reports have been completed and attached to this email.
    </p>"

# Send the mail message
Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
    -Attachments $RetentionPolicyXlsxPath,$RetentionTagReportXlsxPath,$UserRetentionReportXlsxPath,$LicenseReportXlsxPath,$GaReportXlsxPath


Get-PSSession | Remove-PSSession
}