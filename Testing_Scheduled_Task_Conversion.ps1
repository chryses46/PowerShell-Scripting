#This is a test file. Testing converting a csv to xlsx file

#import the module
Import-Module ConvertCsvToXlsx

#create a csv file

$csv = "Test" > D:\reports\test.csv

$csvFilePath = "D:\reports\test.csv"

#convert the csv to xlsx
$xlsxFilePath = Convert-CsvToXlsx($csvFilePath)

Start-Sleep -Seconds 3

if($xlsxFilePath -ne $null){
    Write-Host "$csvFilePath converted to xlsx and is located at $xlsxFilePath." -ForegroundColor Yellow
    }else{Write-Host "Conversion failed" -ForegroundColor Red}
     
#delete the csv

del $csvFilePath