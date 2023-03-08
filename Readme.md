To use this PowerShell script, follow the instructions below:

1. Download the script from [here](https://github.com/a-sajjad72/fontInsta/archive/refs/heads/master.zip) and extract it where you want.
2. From the extracted archive, **copy** the script file (FontInsta.ps1) to same folder where the archive file is placed. If don't want to copy the script, then you have to provide the full path with archive file (done in step 4).
3. `Shift+Right-Click` on the empty area of the folder and select **Open PowerShell window here** from the context menu.
4. Run the script with the command: `.\FontInsta.ps1 -ArchiveName <name_of_font_archive>`

The script will extract the fonts from the archive file and install them into the Windows Fonts folder.

By default, the script will only install fonts from the **root directory** of the archive. If you want to install fonts from **subdirectories**, use the `-SubDir` parameter to specify which subdirectories to install from.

Use the `-All` parameter to include fonts from subdirectories within the specified directory or subdirectories.

# USAGE AND OPTIONS
## USAGE:
    .\FontInsta.ps1 -ArchiveName <[path]\name> [-SubDir <sub_dir_1>,<sub_dir_2>,...] [-InstallFromRoot] [-All] [-Help]
## OPTIONS:
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

# Examples
This will only install the fonts from the root directory of "Roboto.zip"

    .\FontInsta.ps1 -ArchiveName "Roboto.zip"

This will only install the fonts from the "Roboto" and "Poppins" subdirectories.

    .\FontInsta.ps1 -ArchiveName myFonts.zip -Subdir "Roboto","Poppins"

This will install the fonts from the "Roboto" and "Poppins" subdirectories and also the files from the root directory of "myFonts.zip".

NOTE: when using `-SubDir` the fonts in root directory are ignored.

    .\FontInsta.ps1 -ArchiveName myFonts.zip -Subdir "Roboto","Poppins" -InstallFromRoot

This will install the fonts files from the root directory and subdirectories and also thier subdir
    
    .\FontInsta.ps1 -ArchiveName myFonts.zip -All