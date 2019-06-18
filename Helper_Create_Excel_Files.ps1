# Put CSVs into Excel

Function Release-Ref ($ref) { 
([System.Runtime.InteropServices.Marshal]::ReleaseComObject( [System.__ComObject]$ref) -gt 0) 
[System.GC]::Collect() 
[System.GC]::WaitForPendingFinalizers() 
} 

Function ConvertAppendCSVXLSX { 

    [CmdletBinding( 
        SupportsShouldProcess = $True, 
        ConfirmImpact = 'low', 
        DefaultParameterSetName = 'file' 
        )] 
    Param (
        [Parameter( ValueFromPipeline=$True, Position=0, Mandatory=$True, HelpMessage="Name of CSV/s to import")] 
        [ValidateNotNullOrEmpty()] 
        [array]$inputfile, 
        [Parameter( ValueFromPipeline=$False, Position=1, Mandatory=$True, HelpMessage="Name of excel file output")] 
        [ValidateNotNullOrEmpty()] 
        [string]$outfile
    ) 

    Begin {

        #Configure regular expression to match full path of each file 
        #[regex]$regex = "^\w\:" 

        #Find the number of CSVs being imported 
        $count = ($inputfile.count -1) 

        #Create Excel Com Object 
        $excel = new-object -com excel.application
     
        #Disable alerts 
        $excel.DisplayAlerts = $False
     
        #Show Excel application 
        $excel.Visible = $False
     
        #Add workbook 
        $workbook = $excel.workbooks.Add()

        #Define initial worksheet number 
        $i = 1
     
    } 
    Process { 
        ForEach ($file in $inputfile) { 

        #Remove reports directory before file name
        [regex]$regexPrefix = "\w\:\\reports\\"
        $prefix = ($regexPrefix.Matches($file)).Value
        $prefix = $prefix -Replace ('[\\]','\\')
        $fileBase = $file -replace ($prefix,'')

        #If more than one file, create another worksheet for each file 
        If ($i -gt 1) { $workbook.worksheets.Add() | Out-Null } 

        #Use the first worksheet in the workbook (also the newest created worksheet is always 1) 
        $worksheet = $workbook.worksheets.Item(1) 

        #Add name of CSV as worksheet name 
        $worksheet.name = "$((GCI $file).basename)" 

        #Open the CSV file in Excel, must be converted into complete path if not already done 
    
#        If ($regex.ismatch($input)) { 
            $tempcsv = $excel.Workbooks.Open($file)
             
#            } ElseIf ($regex.ismatch("$($input.fullname)")) { 
#            $tempcsv = $excel.Workbooks.Open("$($input.fullname)")
             
#            }Else {
#                $tempcsv = $excel.Workbooks.Open("$($pwd)\$input")

#            } 

        $tempsheet = $tempcsv.Worksheets.Item(1) 
    
        #Copy contents of the CSV file 
        $tempSheet.UsedRange.Copy() | Out-Null
         
        #Paste contents of CSV into existing workbook 
        $worksheet.Paste()
         
        #Close temp workbook 
        $tempcsv.close()
         
        #Select all used cells 
        $range = $worksheet.UsedRange
         
        #Autofit the columns 
        $range.EntireColumn.Autofit() | out-null
         
        $i++
         
        } 
    }
    End { 
    #Save spreadsheet 
    $workbook.saveas("$outfile")
     
    Write-Host -Fore Green "File saved to $outfile"
     
    #Close Excel 
    $excel.quit()

    #Release processes for Excel 
    $a = Release-Ref($range) } 

}

Export-ModuleMember -Function ConvertAppendCSVXLSX