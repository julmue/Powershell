# rewire folder for powershell profile

$psdir = "~\Documents\WindowsPowerShell"

if(Test-Path -PathType Container -Path $psdir){
  Write-Host "direcotry $psdir already exists"
    }
else {
    New-Item
        -Path $psdir
        -ItemType SymbolicLink
        -Value Get-Location
}