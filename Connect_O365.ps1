function Connect-Office365{

$CreateEXOPSSession = (Get-ChildItem -Path $env:userprofile `
-Filter CreateExoPSSession.ps1 -Recurse -ErrorAction SilentlyContinue -Force | `
 Select -Last 1).DirectoryName
 . "$CreateEXOPSSession\CreateExoPSSession.ps1"

Connect-EXOPSSession

Connect-MsolService
}

Export-ModuleMember -Function Conntect-Office365