# This report cleans up reports one month or oldder by moving them into D:\reports\Archived Reports. Those reports are then deleted if older than 90 days.

$reportDir = "D:\reports";

$reportArchiveDir = "D:\reports\Archived Reports";

$reports = Get-ChildItem -path $reportDir -Attributes !Directory+!System

$monthMark = (Get-Date).AddMonths(-1);

$reports | ForEach-Object{
    
    $fileName = ($_.Name);
    $creationDate = ($_.CreationTime);
    $filePath = "$reportDir\$fileName"

    if($creationDate -le $monthMark){

        Write-Host "$fileName was created on $creationDate and will be archived." -ForegroundColor Yellow;

        Move-Item -Path $filePath -Destination $reportArchiveDir;
    }
}


$archivedReports = Get-ChildItem -Path $reportArchiveDir

$threeMonthMark = (Get-Date).AddMonths(-3);

$archivedReports | ForEach-Object{

    $fileName = ($_.Name);
    $creationDate = ($_.CreationTime);
    $filePath = "$reportArchiveDir\$fileName";

    if($creationDate -lt $threeMonthMark){
        
        Write-Host "$fileName was created on $creationDate and will be deleted." -ForegroundColor Red;

        del $filePath;
    }  
}