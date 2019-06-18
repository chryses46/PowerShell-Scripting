# Import the list of user names and secured passwords

$list = Import-CSV -path "D:\required_files\DDBNA_MBAudit.csv"

# Iterate through the entire list. Each iteration is one connection to the specified MSOL/EOL environment.

#Array for all location Stats to be exported
$Stats=@()

$list | ForEach-Object { 

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

        $Users = Get-Mailbox -Filter * | Where-Object {$_.LitigationHoldEnabled -eq $true} | Select PrimarySmtpAddress

        ForEach($User in $Users){
            
            Write-Host "working on $($User.PrimarySmtpAddress)."
            $size = (Get-MailboxFolderStatistics -Identity $user.PrimarySmtpAddress -FolderScope RecoverableItems | Where-Object {$_.Name -eq "Recoverable Items"}).FolderAndSubfolderSize
            $size
            [regex]$regex = "\d+,?\d+,?\d+,?\d+,?\d+";
            $NewSize= ($regex.Matches($size)).Value
            $NewSize= $NewSize -replace '[,]',''
            $NewSize= [decimal]$NewSize / 1GB
            $NewSize= [Math]::Round($NewSize, 2)
            $Stats= `
                $Stats + `
                (Get-Mailboxfolderstatistics `
                     -Identity $User.PrimarySmtpAddress `
                     -FolderScope RecoverableItems | Where-Object {$_.Name -eq "Recoverable Items"} |`
                        Select @{Name="Email Address";Expression={"$($User.PrimarySmtpAddress)"}}, `
                            FolderPath,@{Name="Size in GB";Expression={$NewSize}},@{Name="Tenant Domain";Expression={$domain}})
        }



        Get-PSSession | Remove-PSSession
        }

$date= get-date -Format MMddyy

$csvFilePath = "D:\reports\DDBNA_O365_Mailbox_Retention_Size_Report_$date.csv"
$Stats |Sort-Object -Property "Size in GB" -Descending | Export-Csv -Path $csvFilePath -NoTypeInformation
$xlsxFilePath = Convert-CsvToXls($csvFilePath)

Start-Sleep -Seconds 15
        
$smtpserver= "100.117.4.120"
$from= "PowerShell Server Reporting Server <PSService@ad.corp>"
$to= "DDBNA_O365Admins@ddb.com"
$subject= "DDBNA Office 365 Retention Mailbox Size Report."
$body= "
    <p style='text-indent:1em;'> The </strong>DDBNA</strong> Office 365 Retention Mailbox Size Report has been attached to this email.
    </p>"


Send-MailMessage -To $to -From $from -Subject $subject -Body $body -BodyAsHtml -SmtpServer $smtpserver -Priority High -Attachments $xlsxFilePath