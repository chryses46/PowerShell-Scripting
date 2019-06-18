# Function to convert csv to xlsx files

function Convert-CsvToXls([string] $csvFilePath){

[regex]$regex="^(.*?)\.csv"
$string = ($regex.Matches($csvFilePath)).Groups[1].Value
$xlsxFilePath = $string + ".xlsx"

$excel = New-Object -ComObject Excel.Application 
$excel.Visible = $true
$excel.Workbooks.Open($csvFile).SaveAs($xlsxFilePath,51)
$excel.Quit()

return $xlsxFilePath
}