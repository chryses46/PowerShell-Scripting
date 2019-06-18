<#
This script is meant to test the MSOL and EOL logins of associated tenant Exchange admins.
You can either enter a single user to test or use a csv file with the fields 'Username','Key','Office'. With UPN, secure-string password and Office location, respectively.  
#>

$file = $null
do{
    $User = Read-Host "Would you like to test [1] user, or [All]?"
    if($User -ceq [string]"All"){
        $file = Read-Host "Please input full file path..." 
        $list = Import-CSV -path $file -ErrorAction Stop

        # Test the connections

        $list | ForEach-Object {
            $un = $_.Username
            $pw = $_.Key | ConvertTo-SecureString
            $loc= $_.Office
            $FoundErrors = @()
            Write-Host "Testing the connection for the $loc Office 365 tenant."
            $Creds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $un, $pw
            Connect-MsolService -Credential $Creds -ErrorAction SilentlyContinue
            $domain= (Get-MsolDomain -ErrorAction SilentlyContinue | Where-Object {$_.isDefault -eq $true}).Name

            if($domain -ne $null){
                Write-Host "Connection to MSOL Services Success! Tenant primary domain name: $domain."

                    Write-Host "Now attempting to connect to EOL Services. Please wait..."
                    $EOLConn = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
                    Import-PSSession $EOLConn -ErrorAction SilentlyContinue
                    $exdomain = (Get-OrganizationConfig -ErrorAction SilentlyContinue).Name

                        if($exdomain -ne $null){
                            Write-Host "Connection to EOL Successful! Exchange Online Configuration Name is $exdomain"
                        } else {
                            Write-Host "Connection to $loc's EOL Services failed. Please check credentials and try again."
                            $FoundErrors = $FoundErrors+ "Connection to $loc's EOL Services failed. Please check credentials and try again."
                            }

                } else {
                    if($Error -like "*Connect-MsolService*"){
                        Write-Host "Connection to $loc's MSOL Services failed. Please check credentials and try again." 
                        $FoundErrors = $FoundErrors + "Connection to $loc's MSOL Services failed. Please check credentials and try again."
                        }
                    }


            $ErrorInfo = [Ordered]@{
                Office= $loc
                ReportedErrors= [string]$FoundErrors 
                }
            $LocationErrorInfo = [pscustomobject]$ErrorInfo
            $ErrorBin= $ErrorBin + $LocationErrorInfo

            Write-Host "Testing complete. Closing connection."

            $un = $null
            $pw = $null
            $loc = $null
            $FoundErrors = $null
            $LocationErrorInfo = $null
            $Creds = $null
            $domain = $null
            $EOLConn = $null
            $exdomain = $null
            Get-PSSession | Remove-PSSession
        }

        $ErrorBin | Export-Csv "D:\reports\Failed_Logins.csv" -NoTypeInformation

        if($ErrorBin -ne $null){

            # Prepare to email the failed logins report, if needed.
            $smtpserver= "100.65.2.193"
            $from= "PowerShell Server Reporting <noreply@PSService>"
            $to= "daniel.frank@ddb.com"
            $subject= "Failed O365 Logins Discovered!"
            $body= "
                <p style='text-indent:1em;'>Log in failures have been found during testing. The offices experiencing difficulties are listed in the attached report.
                </p>"

            # Send the mail message
            Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High `
                -Attachments "D:\reports\Failed_Logins.csv"
        }
        break;
    }elseif($User -ceq [int]1){
        
        $Creds = Get-Credential

        Write-Host "Testing the connection to MSOL/EOL."
        
        Connect-MsolService -Credential $Creds -ErrorAction SilentlyContinue
        $domain= (Get-MsolDomain -ErrorAction SilentlyContinue | Where-Object {$_.isDefault -eq $true}).Name

        if($domain -ne $null){
            Write-Host "Connection to MSOL Services Success! Tenant primary domain name: $domain."

                Write-Host "Now attempting to connect to EOL Services. Please wait..."
                $EOLConn = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
                Import-PSSession $EOLConn -ErrorAction SilentlyContinue
                $exdomain = (Get-OrganizationConfig -ErrorAction SilentlyContinue).Name

                    if($exdomain -ne $null){
                        Write-Host "Connection to EOL Successful! Exchange Online Configuration Name is $exdomain"
                    } else {
                        Write-Host "Connection to EOL Services failed. Please check credentials and try again."
                        }
        
        $Creds = $null
        $domain = $null
        $EOLConn = $null
        $exdomain = $null
        Get-PSSession | Remove-PSSession

        } else {
            if($Error -like "*Connect-MsolService*"){
                Write-Host "Connection to MSOL Services failed. Please check credentials and try again.";
                Get-PSSession | Remove-PSSession 
            }
        }
    Write-Host "Testing complete. Closing connection.";
    break;
    }else{Write-Host "Try again. Please enter either exactly '1' or 'All'.";}
}while($User -cne [int]1 -or $User -cne [string]"All")



