# -----------------------------------------------------------------------------
# Module Imports

Import-Module -name posh-git
Import-Module -name get-ChildItemColor
Import-Module -name PSBookmark
# glb	-> Get-LocationBookmark
# rlb	-> Remove-LocationBookmark
# goto	-> Set-LocationBookmarkAsPWD
#	You don't have to type the alias name. Instead, you can just tab complete. This function uses dynamic parameters.
#		goto docs
# save	-> Save-LocationBookmark
#	This will save $PWD as scripts
#		save scripts 
#	This will save C:\Documents as docs
#		save docs C:\Documents

function Get-LocationBookmarkPath([Parameter(Mandatory=$false)][string]$Bookmark) {    
    if(!($Bookmark)) {
        # set default path
        ${env:USERPROFILE}
    } else {
        (Get-LocationBookmark)[$Bookmark]
    }
 }

set-alias bookmarks glb
set-alias path Get-LocationBookmarkPath

# -----------------------------------------------------------------------------
# Directories

$env:path += (";" + $Env:userprofile + "\Scripts")


# -----------------------------------------------------------------------------
# Functions

# This needs some more attentions and sensible defaults or error or whatnot
# function Get-PathFromBookmark ([string]$BookmarkName) {
# 	(Get-LocationBookmark).Get_Item($BookmarkName)
# }

function Kill-ProcessByName ([string]$PorcessName) {
	taskkill /F /IM ($PorcessName + ".exe")
}

function Close-ExplorerWindows {
	(New-Object -comObject Shell.Application).Windows() | foreach-object {$_.quit()} 
}

function Open-Universal ([Parameter(Mandatory=$false)][string]$Path) {
	if(!($Path)) {
		Set-Location ${env:USERPROFILE}
	} elseif((Get-LocationBookmark).ContainsKey($Path)) {
		Set-LocationBookmarkAsPWD($Path)
	} else {
		set-location($Path)
	}
}

function Open-Explorer ([Parameter(Mandatory=$false)][string]$Path) {
	$RealLocation = Get-Location
	Jump-Explorer($Path)
	Set-Location($RealLocation)
}

function Jump-Explorer ([Parameter(Mandatory=$false)][string]$Path) {
	if(!($Path)) {
		$Path = Get-Location
	}
	Open-Universal($Path)
	Start .
}

function Open-UniqueExplorer ([Parameter(Mandatory=$false)][string]$Path) {
	Close-ExplorerWindows
	Open-Explorer($Path)
}

function Jump-UniqueExplorer ([Parameter(Mandatory=$false)][string]$Path) {
	Close-ExplorerWindows
	Jump-Explorer($Path)
}


# hide dotfiles
function Hide-Dotfiles ([Parameter(Mandatory=$false)][string]$Path){
	if(!($Path)) {
		$Path = Get-Location
	}
    Get-ChildItem $Path -force | 
        Where-Object {$_.name -like ".*" -and $_.attributes -match 'Hidden' -eq $false} | 
        Set-ItemProperty -name Attributes -value ([System.IO.FileAttributes]::Hidden)
}



# Generates two functions to get to directory parents
# u4 -> up 4 levels
# uuuu -> up 4 levels
for($i = 1; $i -le 5; $i++){
  $u =  "".PadLeft($i,"u")
  $unum =  "u$i"
  $d =  $u.Replace("u","../")
  Invoke-Expression "function $u { push-location $d }"
  Invoke-Expression "function $unum { push-location $d }"
}



# sudo command in powershell
# https://www.elasticsky.de/2012/12/powershell-sudo/
function elevate-console {
    powershell -new_console:a
}

set-alias sudo elevate-console;



# create symbolic link
function make-link ($link, $target) {    New-Item -Path $link -ItemType SymbolicLink -Value $target}

set-alias rewire make-link


# invoke command on multiple git repositories
function gitmulti (
    [Parameter(Mandatory=$false)][string]$Path,
    [Parameter(Mandatory=$false)][int]$Depth,
    [Parameter(Mandatory=$false)][string]$Cmd
    ) {

    $gitFolderName = ".git"

    # The root directory to perform the command in
	if(!($Path)) {
		$Path = Get-Location
	}

    # How deep down you want to look for .git folders
	if(!($Depth)) {
		$Depth = 2
	}

    # The command you want to perform
 	if(!($Cmd)) {
		$Cmd = "status -s"
	}


    # Finds all .git folders by given path, the -match "h" parameter is for hidden folders 
    $gitFolders = Get-ChildItem -Path $Path -Depth $Depth -Recurse -Force | 
        Where-Object { $_.Mode -match "h" -and $_.FullName -like "*\$gitFolderName" }

    ForEach ($gitFolder in $gitFolders) {

        # Remove the ".git" folder from the path 
        $folder = $gitFolder.FullName -replace $gitFolderName, ""

        Write-Host "Performing git $Cmd in folder: '$folder'..." -foregroundColor "green"

        # Go into the folder
        Push-Location $folder 

        # Perform the command within the folder
        # & "git $Cmd"
        Invoke-Expression "git $cmd"

        # Go back to the original folder
        Pop-Location
    }
}

# -----------------------------------------------------------------------------
# Backups 
# TODO: Refactor better logic ... very rudimentary

function Backup {
    
    Write-Host "Backing up ..." -foregroundColor "green"
    Backup-Portables
    Backup-Dotfiles
    Backup-Kb
    Write-Host "Backing up: Done" -foregroundColor "green"
}


function Backup-Portables {
    Write-Host "Backing up portables" -foregroundColor "green"
    Push-Repos -Path "C:\Users\jmueller\bin\bin_portable"
}

function Backup-Dotfiles {
    Write-Host "Backing up dotfiles" -foregroundColor "green"
    Push-Repos -Path "C:\Users\jmueller\dotfiles"
}


# save everything in kb
function Backup-Kb {


    Write-Host "Backing up knowledge base" -foregroundColor "green"
    
    # backup keepass database (aka gateway)
    Write-Host "Backing up knowledge base: Keepass Database" -foregroundColor "green"
    Push-Location "C:\Users\jmueller\kb\kb_gateway"
    Push-Repo
    Pop-Location
    
    # backup anki
    Write-Host "Backing up knowledge base: Anki Database" -foregroundColor "green"
    Push-Location "C:\Users\jmueller\kb\kb_ram_anki"
    Push-Repo
    Pop-Location

    # backup mindmaps
    Write-Host "Backing up knowledge base: Mindmaps" -foregroundColor "green"
    Push-Repos -Path "C:\Users\jmueller\kb\kb_mm"
    
    # backup notebooks
    Write-Host "Backing up knowledge base: Notebooks" -foregroundColor "green"
    Push-Repos -Path "C:\Users\jmueller\kb\kb_nbs"

    # backup snp specific notebooks
    Write-Host "Backing up knowledge base: SNP Notebooks" -foregroundColor "green"
    Push-Repos -Path "C:\Users\jmueller\kb\kb_nbs_snp"
}

# dispatcher of push-repo by one level
function Push-Repos ([Parameter(Mandatory=$true)][string]$Path){
    $Depth = 0
    $Cmd = "PUSH REPOSITORY TO ORIGIN"

    $OldPath = Get-Location

    set-location($Path)

    $folders = Get-ChildItem -Path $Path -Depth $Depth -Recurse -Force

    ForEach ($folder in $folders) {
    
        Write-Host "Performing git $Cmd in folder: '$folder'..." -foregroundColor "green"

        # Go into the folder
        Push-Location $folder 

        # Perform the command within the folder
        # & "git $Cmd"
        # Invoke-Expression 
        Push-Repo

        # Go back to the original folder
        Pop-Location
    }

    set-location($OldPath)

}

function Push-Repo {
    git add *
    git commit -m Get-Date
    # should check: branch equals master
    git push -u origin master
}


# -----------------------------------------------------------------------------
# Prompt Functions

function Test-Administrator {
    $user = [Security.Principal.WindowsIdentity]::GetCurrent();
    (New-Object Security.Principal.WindowsPrincipal $user).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}

# To change the prompt override the following function
# function prompt {"My Prompt "}

function Wrap-VcsStatus {
	if(WriVcsStatus) {
		
	}
}

function prompt {

	# Save Exit Code
	$realLASTEXITCODE = $LASTEXITCODE
	
	# Prompt
	Write-Host "$ENV:USERNAME@" -NoNewLine -ForegroundColor DarkYellow
	Write-Host "$ENV:COMPUTERNAME" -NoNewline -ForegroundColor Magenta
	Write-Host $($(Get-Location) -replace ($env:USERPROFILE).Replace('\','\\'), "~") -NoNewline -ForegroundColor Blue
	Write-Host " : " -NoNewline -ForegroundColor DarkGray
	Write-Host (Get-Date -Format G) -NoNewline -ForegroundColor DarkMagenta
	Write-Host " : (" -NoNewline -ForegroundColor DarkGray
	Write-VcsStatus
	Write-Host " )" -ForegroundColor DarkGray

	# Reset Exit Code	
	$global:LASTEXITCODE = $realLASTEXITCODE

	return "> "
}

# -----------------------------------------------------------------------------
# Alias
Set-Alias ls get-childitemcolor -option AllScope -Force
Set-Alias dir get-childitemcolor -option AllScope -Force
Set-Alias l get-childitemcolor -option AllScope -Force

# Explorer Interop
Set-Alias e Open-Explorer 
Set-Alias se Open-UniqueExplorer
Set-Alias j Jump-Explorer
Set-Alias sj Jump-UniqueExplorer
Set-Alias ke Close-ExplorerWindows
Set-Alias p Get-PathFromBookmark
Set-Alias c Open-Universal

Set-Alias stop Kill-ProcessByName

Set-Alias google es

# --
# incubator alias
Set-Alias cex Create-NamedExcelFile
# Set-Alias crush Minimize-AllWindows
Set-Alias pop Unminimize-AllWindows

# -----------------------------------------------------------------------------
# Incubator

# function that creates a (named) empty excel file in the current working directory
function Create-NamedExcelFile ([Parameter(Mandatory=$false)][string]$Name) {
	if ($Name) {
		$OutputFile = Join-Path $PWD $Name
		}
	else {
		$OutputFile = Join-Path $PWD  "new"
	}
	
	$excel = New-Object -ComObject excel.application
	$excel.visible = $True
    $workbook = $excel.Workbooks.Add()
	$workbook.SaveAs($OutputFile)
 	$excel.Quit()

}


# function that brings the windows of an application into focus
# * function that minimizes all windows but those of one process
# * function that maximizes all windows of a specific process
function Show-Process($Process, [Switch]$Maximize)
{
# does not work in all cases
(New-Object -ComObject WScript.Shell).AppActivate((get-process $Process).MainWindowTitle)
}

function Minimize-AllWindows() {
$shell = New-Object -ComObject "Shell.Application"
$shell.minimizeall()
}

function Unminimize-AllWindows() {
$shell = New-Object -ComObject "Shell.Application"
$shell.undominimizeall()
}

# todo
function zoom([Parameter(Mandatory=$false)][string]$Process) {
    test $Process    
    }


function test([Parameter(Mandatory=$false)][string]$Process) {
    [void] [System.Reflection.Assembly]::LoadWithPartialName("'Microsoft.VisualBasic")
    Get-Process | Where-Object {$_.Name -like $Process}
    # $a = Get-Process | Where-Object {$_.Name -like $Process}
    [Microsoft.VisualBasic.Interaction]::AppActivate($a.ID)
}
