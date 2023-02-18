param (
    [Parameter(position = 0)][string]$ArchiveName="",
    [Parameter(ValueFromRemainingArguments = $true)][string[]]$Subdirectories,
    [switch]$InstallFromRoot = $false,
    [switch]$All = $false,
    [switch]$Help = $false
)
function Show-Help {
    $helpText = @"

Install fonts from the specified archive file into the Windows Fonts folder.
Usage: FontInsta.ps1 -ArchiveName <name> [-SubDirectories <sub dir 1>,<sub dir 2>,...] [-InstallFromRoot] [-All] [-Help]
Parameters:
-ArchiveName        The name of the font archive to install.
                    (required) 
-SubDirectories     An array of subdirectories to install
                    fonts from (optional). -Subdirectories
                    should be separated by commas and enclosed
                    within double quotes.
-InstallFromRoot    A switch indicating whether to install
                    fonts from the root directory of the
                    archive (optional, default is off).
-All                A switch indicating whether to install
                    files from subdirectories.
-Help               Display the help message.
"@
    Write-Host $helpText
}

if ($ArchiveName -eq "") {
    Write-Host "Error: Archive name is required." 
    Show-Help
    return
}

$FontsFolder = "Font$([System.IO.Path]::GetFileNameWithoutExtension($ArchiveName))"

Expand-Archive -Path $ArchiveName -DestinationPath $FontsFolder -Force

$FONTS = 0x14
$CopyOptions = 4 + 16;
$objShell = New-Object -ComObject Shell.Application
$objFolder = $objShell.Namespace($FONTS)

$allFonts = ""

if ($InstallFromRoot) {     
    $allFonts = Get-ChildItem -Path "$FontsFolder" -File | Where-Object { $_.Extension -in ".ttf", ".otf", ".ttc", ".fon", ".fnt", ".pfb", ".pfm", ".cff", ".afm", ".postscript" }
    foreach ($font in $allFonts) {
        $dest = "C:\Windows\Fonts\$($font.Name)"
        if (Test-Path -Path $dest) {
            Write-Host "Font $($font.Name) already installed"
        } else {
            Write-Host "Installing $($font.Name)"
            $CopyFlag = [String]::Format("{0:x}", $CopyOptions);
            $objFolder.CopyHere($font.fullname,$CopyFlag)
        }
    }
}
if ($Subdirectories) {
    foreach ($sub_dirs in $Subdirectories) {
        if ($All) {
            $allFonts = Get-ChildItem -Path "$FontsFolder\$sub_dirs" -Recurse | Where-Object { $_.Extension -in ".ttf", ".otf", ".ttc", ".fon", ".fnt", ".pfb", ".pfm", ".cff", ".afm", ".postscript" }
        }
        else {  
            $allFonts = Get-ChildItem -Path "$FontsFolder\$sub_dirs"-File | Where-Object { $_.Extension -in ".ttf", ".otf", ".ttc", ".fon", ".fnt", ".pfb", ".pfm", ".cff", ".afm", ".postscript" }
        }
        foreach ($font in $allFonts) {
            $dest = "C:\Windows\Fonts\$($font.Name)"
            if (Test-Path -Path $dest) {
                Write-Host "Font $($font.Name) already installed"
            } else {
                Write-Host "Installing $($font.Name)"
                $CopyFlag = [String]::Format("{0:x}", $CopyOptions);
                $objFolder.CopyHere($font.fullname,$CopyFlag)
            }
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
    foreach ($font in $allFonts) {
        $dest = "C:\Windows\Fonts\$($font.Name)"
        if (Test-Path -Path $dest) {
            Write-Host "Font $($font.Name) already installed"
        } else {
            Write-Host "Installing $($font.Name)"
            $CopyFlag = [String]::Format("{0:x}", $CopyOptions);
            $objFolder.CopyHere($font.fullname,$CopyFlag)
        }
    }
}
Remove-Item -Path $FontsFolder -Recurse -Force