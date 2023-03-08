param (
    [Parameter(position = 0)][string]$ArchiveName="",
    [string[]]$SubDir,
    [switch]$InstallFromRoot = $false,
    [switch]$All = $false,
    [switch]$Help = $false
)
function Show-Help {
    $helpText = @"

Install fonts from the specified archive file into the Windows Fonts folder.
Usage: FontInsta.ps1 -ArchiveName <name> [-SubDir <sub_dir_1>,<sub_dir_2>,...] [-InstallFromRoot] [-All] [-Help]
Parameters:
-ArchiveName        The name of the font archive to install
                    (required). You can also specify the path
                    with archive file.
-SubDir             An array of subdirectories to install
                    fonts from (optional). -Subdir
                    should be separated by commas and enclosed
                    within double quotes.
-InstallFromRoot    A switch indicating whether to install
                    fonts from the root directory of the
                    archive (optional, default is off). Comes
                    in handy while using -SubDir option.
-All                A switch indicating whether to install
                    files from subdirectories.
-Help               Display the help message.
"@
    Write-Host $helpText
}

function installFonts {
    foreach ($font in $allFonts) {
        if (Test-Path "C:\Windows\Fonts\$($File.Name)") {
            Write-Host "Font $($font.Name) already installed"
        }
        else {
            Write-Host "Installing $($font.Name)"
            $CopyFlag = [String]::Format("{0:x}", $CopyOptions);
            $objFolder.CopyHere($font.fullname, $CopyFlag)
        }
    }
}

if ($ArchiveName -eq "") {
    Write-Host "Error: Archive name is required." 
    Show-Help
    return
}

$path = Split-Path -Path $ArchiveName -Parent
$filename = Split-Path -Path $ArchiveName -Leaf

$FontsFolder = "Font$([System.IO.Path]::GetFileNameWithoutExtension($filename))"

Expand-Archive -Path $ArchiveName -DestinationPath $FontsFolder -Force

$FONTS = 0x14
$CopyOptions = 4 + 16;
$objShell = New-Object -ComObject Shell.Application
$objFolder = $objShell.Namespace($FONTS)

$allFonts = ""

if ($InstallFromRoot) {     
    $allFonts = Get-ChildItem -Path "$FontsFolder" -File | Where-Object { $_.Extension -in ".ttf", ".otf", ".ttc", ".fon", ".fnt", ".pfb", ".pfm", ".cff", ".afm", ".postscript" }
    if ($allFonts) {
        installFonts
    }
    else {
        Write-Host "Sorry no fonts found to install"
    }
}

if ($SubDir) {
    foreach ($sub_dirs in $SubDir) {
        if ($All) {
            $allFonts = Get-ChildItem -Path "$FontsFolder\$sub_dirs" -Recurse | Where-Object { $_.Extension -in ".ttf", ".otf", ".ttc", ".fon", ".fnt", ".pfb", ".pfm", ".cff", ".afm", ".postscript" }
        }
        else {  
            $allFonts = Get-ChildItem -Path "$FontsFolder\$sub_dirs"-File | Where-Object { $_.Extension -in ".ttf", ".otf", ".ttc", ".fon", ".fnt", ".pfb", ".pfm", ".cff", ".afm", ".postscript" }
        }
        if ($allFonts) {
            installFonts
        }
        else {
            Write-Host "Sorry no fonts found to install"
        }
    }
}
elseif (!($InstallFromRoot)) {
    if ($All) {
        $allFonts = Get-ChildItem -Path $FontsFolder -Recurse | Where-Object {$_.Extension -in ".ttf", ".otf", ".ttc", ".fon", ".fnt", ".pfb", ".pfm", ".cff", ".afm", ".postscript"}
    }
    else {  
        $allFonts = Get-ChildItem -Path "$FontsFolder" -File | Where-Object { $_.Extension -in ".ttf", ".otf", ".ttc", ".fon", ".fnt", ".pfb", ".pfm", ".cff", ".afm", ".postscript" }
    }
    if ($allFonts) {
        installFonts
    }
    else {
        Write-Host "Sorry no fonts found to install"
    }
}
Remove-Item -Path $FontsFolder -Recurse -Force