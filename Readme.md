To use this PowerShell script, follow the instructions below:

1. Download the script from [here]()
2. Copy the script file (FontInsta.ps1) to same folder where the archive file is placed.
3. `Shift+Right-Click` on the empty area of the folder and select ***Open PowerShell window here*** from the context menu.
4. Run the script with the command: `.\FontInsta.ps1 -ArchiveName <name of font archive>`

The script will extract the fonts from the archive file and install them into the Windows Fonts folder.

By default, the script will only install fonts from the **root directory** of the archive. If you want to install fonts from **subdirectories**, use the `-SubDirectories` parameter to specify which subdirectories to install from.

Use the `-All` parameter to include fonts from subdirectories within the specified directory or subdirectories.

# USAGE AND OPTIONS
## USAGE:
    .\FontInsta.ps1 -ArchiveName <name> [-SubDirectories <sub dir 1>,<sub dir 2>,...] [-InstallFromRoot] [-All] [-Help]
## OPTIONS:
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

# Examples
This will only install the fonts from the root directory of "Roboto.zip"

    .\FontInsta.ps1 -ArchiveName "Roboto.zip"

This will only install the fonts from the "Roboto" and "Poppins" subdirectories.

    .\FontInsta.ps1 -ArchiveName myFonts.zip -Subdirectories "Roboto","Poppins"

This will install the fonts from the "Roboto" and "Poppins" subdirectories and also the files from the root directory of "myFonts.zip"

    .\FontInsta.ps1 -ArchiveName myFonts.zip -Subdirectories "Roboto","Poppins" -InstallFromRoot

This will install the fonts files from the root directory and subdirectories and also thier subdirectories
    
    .\FontInsta.ps1 -ArchiveName myFonts.zip -All