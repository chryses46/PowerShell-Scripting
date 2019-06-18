    #Get Licensed users and associated lincenses in the tenant

    $list = Import-CSV -path "D:\required_files\DDBNA_MBAudit.csv"

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


    #Create list of exluded MFA users


        $ExcludeList = @()

        $MFAExcludeID = (Get-MsolGroup -all | Where-Object {$_.displayname -eq "MFA Service Account Exclusion"}).ObjectID

        Get-MsolGroupMember -GroupObjectId $MFAExcludeID | ForEach-Object {

            $Props = [Ordered]@{
                UPN= $_.emailaddress
                }

            $obj = [pscustomobject]$Props
            $ExcludeList += $obj
         }


<#
        $st = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
        $st.RelyingParty = "*"
        $st.State = $null
        $sta = @($st)

        $LicenseUserInfo += Get-MsolUserByStrongAuthentication -Requirements $sta
#>
        Write-Host "Obtaining User License Report for the $domain Office 365 tenant."
    
        $LicenseUsers = Get-MsolUser -All | Where {$_.isLicensed -eq "True"} | Select UserPrincipalName,Licenses,StrongAuthenticationRequirements,BlockCredential,Office,DisplayName
        

        #Iterate through users and add their MSOL/EOL data to the array

        foreach($user in $LicenseUsers){

            $UserUPN = $user.userprincipalname

            #Determine if the user is excluded for MFA

            $Excluded = "-"

            foreach ($ExcludedUser in $ExcludeList){
            
                if($UserUPN -eq $ExcludedUser.UPN){
                    $Excluded = "Excluded"
                    }
            }


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


                $Props = [Ordered]@{
                    UPN = $User.UserPrincipalName;
                    "Display Name" = $User.DisplayName;
                    MFA=$User.StrongAuthenticationRequirements.State;
                    "MFA Excluded"=$Excluded;
                    Domain=$domain;
                    Office=$user.office;
                    SignInBlocked=$user.BlockCredential;
                    License = $lic;
                    }

                $obj = [pscustomobject]$Props
                $LicenseUserInfo += $obj
            }

        get-pssession | Remove-Pssession

}

#Export requested information

$date= get-date -Format MMddyy

$csvFilePath = "D:\reports\DDBNA_O365_MFA_Report_$($date).csv"

$LicenseUserInfo | Export-Csv -Path $csvFilePath -NoTypeInformation

$XlsxFilePath = Convert-CsvToXlsx($csvFilePath);

Start-Sleep -Seconds 15

$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting <PSService@ad.corp>"
$to= "DDBNA_O365Admins@ddb.com"
$subject= "DDB NA MFA Report"
$body= "Please see attached."

# Send the mail message #DDBNA_O365Admins@ddb.com

Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
            -Attachments $XlsxFilePath  