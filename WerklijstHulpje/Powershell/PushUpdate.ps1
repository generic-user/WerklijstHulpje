# INFO
# By Cattoor Bjorn (bjorn.cattoor@infrabel.be)
# GITHUB https://github.com/generic-user/WerklijstHulpje
# Download releases: https://github.com/generic-user/WerklijstHulpje/releases
# Doel: 
# 1. Maak een lokale copy van de server (MIRROR)
# 2. voedt de gevonden bestanden aan onze werklijsthulpje.exe
# 3. maak van de gedownloade badge bestanden een backup in een archiefbestand
# 4. verwijder alle tijdelijke bestanden
# 5. upload de gewijzigde bestanden terug naar de REMOTE_LOCATION (mirror)
# Funtie-compleet 18 november 2021. 

# functions

function Remove-Folders {
    [CmdletBinding()]
    param([parameter(Mandatory = $true)][System.IO.DirectoryInfo[]] $Folders) 
    if ( $Folders -eq "" -or $Folders.Count -eq 0 ) {
        Write-Verbose "No folders to delete."
    }
    else {
        foreach ($Folder in $Folders) {
            Remove-Folder $Folder
        } 
    }
}
function Remove-Folder {
    [CmdletBinding()]
    param ([parameter(Mandatory = $true)][System.IO.DirectoryInfo] $FolderPath ) 
    if ( $FolderPath.Exists ) {
        #http://stackoverflow.com/questions/7909167/how-to-quietly-remove-a-directory-with-content-in-powershell/9012108#9012108
        Get-ChildItem -Path  $FolderPath.FullName -Force -Recurse | Remove-Item -Force -Recurse
        Remove-Item $FolderPath.FullName -Force
        Write-Verbose "Deleted folder: $FolderPath.FullName"
    }
    else { Write-Warning "Failed to delete folder-path: $FolderPath.FullName, folder was not found." }
}

function Remove-Files {
    [CmdletBinding()]
    param ([parameter(Mandatory = $true)][System.IO.FileInfo[]] $FilesToDelete ) 
    if ( $FilesToDelete -eq "" -or $FilesToDelete.Count -eq 0 ) {
        Write-Verbose "No files to delete."
    }
    else {
        foreach ($File in $FilesToDelete) {
            if ( $File.Exists ) {
                Remove-Item $File.FullName -Force -ErrorAction Stop
                Write-Verbose "Deleted file:  $File.FullName"
            }
            else { Write-Warning "Failed to delete file-path: $File.FullName, file was not found." }
        } 
    }
}

function Invoke-Utility {
    <#
.SYNOPSIS
Invokes an external utility, ensuring successful execution.
https://stackoverflow.com/a/48877892

.DESCRIPTION
Invokes an external utility (program) and, if the utility indicates failure by 
way of a nonzero exit code, throws a script-terminating error.

* Pass the command the way you would execute the command directly.
* Do NOT use & as the first argument if the executable name is not a literal.

.EXAMPLE
Invoke-Utility git push

Executes `git push` and throws a script-terminating error if the exit code
is nonzero.
#>
    $exe, $argsForExe = $Args
    $ErrorActionPreference = 'Continue' # to prevent 2> redirections from triggering a terminating error.
    # The ampersand (&) here tells PowerShell to execute that command, instead of treating it as a cmdlet or a string. 
    try { & $exe $argsForExe } catch { Throw } # catch is triggered ONLY if $exe can't be found, never for errors reported by $exe itself
    if ($LASTEXITCODE) { Throw "$exe indicated failure (exit code $LASTEXITCODE; full command: $Args)." }
}

function New-TemporaryDirectory {
    $parent = [System.IO.Path]::GetTempPath()
    $name = [System.IO.Path]::GetRandomFileName()
    New-Item -ItemType Directory -Path (Join-Path $parent $name)
}

function Start-Executable {
    param(
        [String] $FilePath,
        [String[]] $ArgumentList
    )
    $OFS = " "
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo.FileName = $FilePath
    $process.StartInfo.Arguments = $ArgumentList
    $process.StartInfo.UseShellExecute = $false
    $process.StartInfo.RedirectStandardOutput = $true
    if ( $process.Start() ) {
        $output = $process.StandardOutput.ReadToEnd() `
            -replace "\r\n$", ""
        if ( $output ) {
            if ( $output.Contains("`r`n") ) {
                $output -split "`r`n"
            }
            elseif ( $output.Contains("`n") ) {
                $output -split "`n"
            }
            else {
                $output
            }
        }
        $process.WaitForExit()
        & "$Env:SystemRoot\system32\cmd.exe" `
            /c exit $process.ExitCode
    }
}

################################################################################
###############CODE####STARTS####HERE###########################################
################################################################################

# INPUT

# This is the path to the server containing the user files
$REMOTE_LOCATION = "X:\I-AM.A3\8.DRAWING_OFFICE\8.1.Common\Werktijdregistratie" 
# This is the path to the server containing the user files (For DEBUGGING)
# $REMOTE_LOCATION = "C:\TEMP\fake_server"
# Where do you want to mirror the server files 2 --> Uses robocopy to make a mirror.
$MIRROR_SERVER_PATH = "C:\TEMP\mirror_server" 

# Path to the template you want to use
$REMOTE_TEMPLATE_FILE = "C:\Users\cwn8400\OneDrive - INFRABEL\administrative\excel\werklijsten en P30bis\werkmap aanpassingen 2021\_template.Badge2021_01.xlsm"

# Path to where the executable of the C# application is
$APP = "C:\Users\cwn8400\Documents\GitHub\WerklijstHulpje\WerklijstHulpje\bin\x64\Release\WerklijstHulpje.exe"

# Where do you want to write your backups to?
$backups = "c:\backups\werklijsthulpje\"

# RoboCOPY arguments quick reference: extra arguments: /v (verbose) /l (debug only)
# INFO: https://docs.microsoft.com/en-us/windows-server/administration/windows-commands/robocopy
$ROBOCOPY_ARGUMENTS = @('/tee', '/log:C:\temp\copy-log.txt', '/mt', '/z', '/xd', '"00 I.NW.06"', '"00 Pensioen"', '"bck"', '"bak"', '/xf', '"*~$*"', '"bck"', '"bak"', '/lev:2', '/s', '/v', '/mir'; '/dcopy:T', '/XA:H', '/W:5')

# Explanation pattern [voornaam].[achternaam].Badge[YYYY].xlsm
$year = '2021'
$FILE_FILTER_STRING = '*.*.Badge' + $year + '.xlsm'

################################################################################
################################################################################
################################################################################

### NO CHANGE AFTER THIS ###

#Requires -Version 7.0
'Requires version 7.0'
"Running PowerShell $($PSVersionTable.PSVersion)."

$title = 'Welcome to werklijsthulpje'
$question = 'Selecteer een actie die je wenst uit te voeren'
$choices = 'Badge bestanden &downloaden' , 'Badge bestanden uploaden naar de &server' , 'Upgraden volgen &template' , '&Quit'

$decision = $Host.UI.PromptForChoice($title, $question, $choices, 3)

switch ($decision) {
    0 { 
        if ( (Test-Path $REMOTE_LOCATION)) {
            Write-Warning "Copying files FROM remote: $REMOTE_LOCATION TO local mirror folder: $MIRROR_SERVER_PATH"
            Start-Executable robocopy.exe $REMOTE_LOCATION, $MIRROR_SERVER_PATH, $FILE_FILTER_STRING, $ROBOCOPY_ARGUMENTS
        }
        else {
            Throw ("Path error, please check your provided paths: $REMOTE_LOCATION")
        }
    }
    1 { 
        if ((Test-Path $MIRROR_SERVER_PATH) -and (Test-Path $REMOTE_LOCATION)) {
            Write-Warning "Bellow is a preview ... please review it before continuing." -WarningAction Continue
            Start-Executable robocopy.exe $MIRROR_SERVER_PATH , $REMOTE_LOCATION , $FILE_FILTER_STRING, $ROBOCOPY_ARGUMENTS, '/l'
            Read-Host "Above is a preview ... please review it before continuing, press enter when ready ..."
            Write-Warning "Do you REALLY want to continue with the opperation?" -WarningAction Inquire
            Start-Executable robocopy.exe $MIRROR_SERVER_PATH , $REMOTE_LOCATION , $FILE_FILTER_STRING, $ROBOCOPY_ARGUMENTS
        }
        else {
            Throw ("Path error, please check your provided paths: $REMOTE_LOCATION local mirror: $MIRROR_SERVER_PATH")
        }
    }
    2 { 
        $APPName = (Get-Item $APP).BaseName
        if ((Test-Path -Path $MIRROR_SERVER_PATH) -eq $false)
        { Throw ("$MIRROR_SERVER_PATH is not valid") }

        if ((Test-Path -Path $REMOTE_TEMPLATE_FILE) -eq $false)
        { Throw ("$REMOTE_TEMPLATE_FILE is not valid") }

        # Get a list of files we want to process $_.Name -match ".*(\.Badge2021\.xlsm)" -and
        $Remote_Originals = Get-ChildItem $MIRROR_SERVER_PATH -Directory | Where-Object { $_.Name -cnotlike "00 I.NW.06" -and $_.Name -cnotlike "bck" } | Get-ChildItem -File | Where-Object { $_.Name -match ".*(\.Badge2021\.xlsm)$" } 

        Write-Host "Server adress is: $MIRROR_SERVER_PATH" -InformationAction Continue 

        # Show file's found
        $OriginalsCount = $Remote_Originals.Count
        Write-Host "We found $OriginalsCount files to handle." -InformationAction Continue 
        if ($OriginalsCount -eq 0)
        { Throw "We did not find any filles to handle." }

        # Get a temporaly directory to work in
        $Temp_WorkingFolder = New-TemporaryDirectory
        Write-Host "Temporary working-directory => $Temp_WorkingFolder"

        # INFO: Reason for try:  Make sure to clean up afterwards
        try {
            # Copy template to our workingdirectory
            $Local_TemplateFilePath = $Temp_WorkingFolder.FullName + "\" + (Get-Item $REMOTE_TEMPLATE_FILE).Name  
    
            Copy-Item -Path $REMOTE_TEMPLATE_FILE -Destination $Local_TemplateFilePath
            Write-Host "Template file copied to => $Local_TemplateFilePath"

            # we need a string to pass the files as arguments to our c# app
            # in this grammar:  WerklijstHulpje.exe -t templatefile.txt -o file1.txt file2.txt
            $argument_t = " -t ""$Local_TemplateFilePath"""
            $argument_o = " -o "

            # Copy Files in from remote destination
            $toDest = $Temp_WorkingFolder.FullName
            Write-Warning "Copying files from remote TO this location --> ($toDest)"

            ForEach-Object -InputObject $Remote_Originals -Parallel {
                Copy-Item -Path $_.FullName -Destination $($using:toDest) -Force -ErrorAction Stop
            }

            Write-Host ( "  ok => " + $Temp_WorkingFolder.FullName)

            # Build arguments
            foreach ($file in $Remote_Originals) {
                $local_FilePath = $Temp_WorkingFolder.FullName + "\" + $file.Name
                $local_FilePathResult = $Temp_WorkingFolder.FullName + "\" + $file.BaseName + ".new" + $file.Extension
                $argument_o += """" + $local_FilePath + """ "
            }

            Write-Warning ("Starting application => " + $APPName)
            $ensemble = ($APP + $argument_t + $argument_o)
            # start the application
            $succes = Invoke-Expression $ensemble
            if ($succes -contains "Failed:")
            { Throw $succes }
            else { $succes }

            Write-Warning "Copying files TO remote ..." 

            # Move result to remote server
            foreach ($file in $Remote_Originals) {
                $local_FilePathResult = $Temp_WorkingFolder.FullName + "\" + $file.BaseName + ".new" + $file.Extension
                Move-Item -Path $local_FilePathResult -Destination $file.FullName -Force -ErrorAction Stop
            }

            Write-Host (" $OriginalsCount new files copied to destination => " + $MIRROR_SERVER_PATH)

            # Zip original files, move into backup-folder
            if ((Test-Path $backups -PathType Container) -eq $false )
            { mkdir $backups -ErrorAction Continue }
    
            # Create the actual zip-file
            Write-Host ("Zipping files to backup ...")
            $zipName = $backups + $Temp_WorkingFolder.BaseName + ".zip"
            Compress-Archive -Path ($Temp_WorkingFolder.FullName + "*")  -DestinationPath $zipName -Force -ErrorAction Stop -CompressionLevel Optimal
            Write-Host ("  ok => " + $zipName)

        }
        catch { Throw }
        finally {
            # Clean up after ourself
            Remove-Folder $Temp_WorkingFolder 
            Write-Host "Cleaning up the mess I made here: $Temp_WorkingFolder => Removed"
        }

        Write-Host "We are done here, thank you for using with Werklijsthulpje!"
        exit
    }
    Default {
        Write-Host "Canceled by user"
        exit
    }
}





