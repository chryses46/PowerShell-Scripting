﻿    #Get Licensed users and associated lincenses in the tenant

    $list = Import-CSV -path "D:\required_files\DDBWW_MBAudit.csv"

    $LicenseUserInfo = @()

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

        Write-Host "Obtaining User License Report for the $domain Office 365 tenant."
    
        $LicenseUsers = Get-MsolUser -All | Where {$_.isLicensed -eq "True"} | Select UserPrincipalName,Licenses,StrongAuthenticationRequirements,BlockCredential,Office
        

        #Iterate through users and add their MSOL/EOL data to the array

        foreach($user in $LicenseUsers){
            
            #$exInfo= (Get-Mailbox -Identity $user.UserPrincipalName)
            
            $lic=[string]$user.Licenses.AccountSku.SkuPartNumber

                if($exInfo.isDirSynced -eq $true){
                    $cloud = "AD-Synced Account"
                    }else{$cloud= "Cloud-Only Account"}

                if($lic -eq "ENTERPRISEPACK"){
                    $lic= "E3"
                    }elseif($lic -eq "EXCHANGEENTERPRISE"){
                    $lic= "P2"
                    }elseif($lic -eq "EXCHANGESTANDARD"){
                    $lic= "P1"
                    }
                    
                #if($User.StrongAuthenticationRequirements.State -eq "Disabled"){
                 #    $mfa = $False
                  #  }else{$mfa = $True}

        #if($mfa -eq $False -and $lic -eq "E3"){
                $Props = [Ordered]@{
                    UPN = $User.UserPrincipalName;
                    License = $lic;
                    Tenant=$exdomain;
                    MFA=$User.StrongAuthenticationRequirements.State;
                    SignInBlocked=$user.BlockCredential;
                    Domain=$domain;
                    Office=$user.office
                    }

                $obj = [pscustomobject]$Props
                $LicenseUserInfo += $obj
            }
        #}

        get-pssession | Remove-Pssession

}

#Export requested information

$date= get-date -Format MMddyy
    
$csvFilePath = "D:\reports\DDBWW_MFA_And_License_Report_$($date).csv"
        
$LicenseUserInfo | Export-Csv -Path $csvFilePath -NoTypeInformation

#Convert it to XLSX
$xlsxFilePath = Convert-CsvToXlsx($csvFilePath);

Start-Sleep -Seconds 15

$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
#wwpowershellreporting@ddb.com
$to= "wwpowershellreporting@ddb.com"
$subject= "DDBWW MFA and License Report"
$body= "Please see attached."

# Send the mail message

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
            -Attachments $xlsxFilePath  