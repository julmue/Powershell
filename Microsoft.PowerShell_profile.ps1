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

# -----------------------------------------------------------------------------
# Directories

$env:path += (";" + $Env:userprofile + "\Scripts")


# -----------------------------------------------------------------------------
# Functions

# This needs some more attentions and sensible defaults or error or whatnot
function Get-PathFromBookmark ([string]$BookmarkName) {
	(Get-LocationBookmark).Get_Item($BookmarkName)
}

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

# -----------------------------------------------------------------------------
# Scaffolding and Build Functions

function Create-LetterDe ([Parameter(Mandatory=$false)][string]$LetterName) {

	$Source = ($ENV:userprofile + "\Templates\Brief\")

	if($LetterName) {
		$Target = (".\" + $LetterName)
	} else {
		$Target = ".\Brief"
	}
	Copy-Item -Recurse $Source $Target

	if($LetterName) {
		Rename-Item -Path ($Target + "\Brief.pandoc") -NewName ($LetterName + ".pandoc")
	}
}

function Build-LetterDe ([Parameter(Mandatory=$true)][string]$LetterName) {
	$Filename = (Get-Item $LetterName).Basename
	$Accentcolor = "B93C5A"
	$Textcolor = "444444"

	pandoc (".\" + $Filename + ".pandoc") `
		--template=letter.latex `
		-o (".\" + $Filename + ".pdf") `
		-V lang=german `
		-V fontfamily=mathpazo `
		-V textcolor=$Textcolor `
		-V accentcolor=$Accentcolor `
		-V letteroption=DIN `
		-M classoption='fromalign=right' `
		-M classoption='foldmarks=false' `
		-M classoption='fromrule=aftername'
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

Set-Alias kip Kill-ProcessByName

# Scaffolding 
Set-Alias brief Create-LetterDe

# -----------------------------------------------------------------------------
# Basic Setup
set-location C:\Users\jmueller