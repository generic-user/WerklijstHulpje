# INFO
# By Cattoor Bjorn
# Doel: 
# 1. zoek alle werklijst bestanden
# 2. maak hiervan een backup in een archiefbestand
# 3. voedt de gevonden bestanden aan onze werklijsthulpje
# 4. verwijder de bestanden op de REMOTE_LOCATION
# 5. kopieer de nieuwe bestanden
# Funties compleet 6januarie 2021. 

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

workflow copyfilesFromRemote {
    param($files,$destination)
    foreach -parallel -ThrottleLimit 4 ($file in $files) {
        Copy-Item -Path $file.FullName -Destination $destination -Force   
    }
}

################################################################################
###############CODE####STARTS####HERE###########################################
################################################################################


# INPUT
 
$REMOTE_LOCATION = "C:\IAM-A3-8-DO\8.1.Common\Werktijdregistratie" 
$REMOTE_TEMPLATE_FILE = "C:\IAM-A3-8-DO\8.1.Common\Werktijdregistratie\00 I.NW.06\Beheer\00 Startbestanden\2021\_template.Badge2021.xlsm"
$APP = "F:\Source\Repos\concepts\WerklijstHulpje\WerklijstHulpje\bin\Debug\WerklijstHulpje.exe"
$backups = "c:\backups\werklijsthulpje\"

################################################################################
################################################################################
################################################################################

### NO CHANGE AFTER THIS ###
$APPName = (gi $APP).BaseName
if ((Test-Path -Path $REMOTE_LOCATION) -eq $false)
   { Throw "$REMOTE_LOCATION is not valid" }

if ((Test-Path -Path $REMOTE_TEMPLATE_FILE) -eq $false)
    { Throw "$REMOTE_LOCATION is not valid" }

# Get a list of files we want to process $_.Name -match ".*(\.Badge2021\.xlsm)" -and
$Remote_Originals = Get-ChildItem $REMOTE_LOCATION -Directory | Where-Object { $_.Name -cnotlike "00 I.NW.06" } | Get-ChildItem -File -Recurse | Where-Object { $_.Name -match ".*(\.Badge2021\.xlsm)$"  } 

# Show file's found
$OriginalsCount = $Remote_Originals.Count
Write-Host "We found $OriginalsCount files to handle." -InformationAction Continue 
if ($OriginalsCount -eq 0)
    { Throw "We did not find any filles to handle." }

# Get a temporaly directory to work in
$WorkDir = New-TemporaryDirectory
Write-Host "Temporary working-directory => $WorkDir"

# Make sure to clean up afterwards
try {
    # Copy template to our workingdirectory
    $Local_TemplateFilePath = $WorkDir.FullName + "\" + (Get-Item $REMOTE_TEMPLATE_FILE).Name  
    
    Copy-Item -Path $REMOTE_TEMPLATE_FILE -Destination $Local_TemplateFilePath
    Write-Host "Template file copied to => $Local_TemplateFilePath"

    # we need a string to pass the files as arguments to our c# app
    # in this grammar:  WerklijstHulpje.exe -t templatefile.txt -o file1.txt file2.txt
    $argument_t = " -t ""$Local_TemplateFilePath"""
    $argument_o = " -o "

    # Copy Files in parralel from remote destination
    Write-Host "Copying files from remote ..."
    copyfilesFromRemote -files $Remote_Originals -destination $WorkDir.FullName -ErrorAction Stop
    Write-Host ( "  ok => " + $WorkDir.FullName)

    # Build arguments
    foreach ($file in $Remote_Originals){
        $local_FilePath = $WorkDir.FullName + "\" + $file.Name
        $local_FilePathResult = $WorkDir.FullName + "\" + $file.BaseName  + ".new" +  $file.Extension
        $argument_o +=  """" + $local_FilePath + """ "
    }

    Write-Host ("Starting application => " + $APPName)
    $ensemble = ($APP + $argument_t + $argument_o)
    # start the application
    $succes = Invoke-Expression $ensemble
    if ($succes -contains "Failed:")
        {Throw $succes}
    else { $succes }

    Write-Host ("Copying files to remote ...")
    # Move result to remote server
    foreach ($file in $Remote_Originals){
        $local_FilePath = $WorkDir.FullName + "\" + $file.Name
        $local_FilePathResult = $WorkDir.FullName + "\" + $file.BaseName  + ".new" +  $file.Extension
        Move-Item -Path $local_FilePathResult -Destination $file.FullName -Force -ErrorAction Stop
    }

    Write-Host ("  ok => " + $REMOTE_LOCATION)

    # Zip original files, move into backup-folder
    if ((Test-Path $backups -PathType Container) -eq $false )
        {mkdir $backups -ErrorAction Continue}
    
    # Create the actual zip-file
    Write-Host ("Zipping files to backup ...")
    $zipName = $backups + $WorkDir.BaseName + ".zip"
    Compress-Archive -Path ($WorkDir.FullName + "*")  -DestinationPath $zipName -Force -ErrorAction Stop -CompressionLevel Optimal
    Write-Host ("  ok => " + $zipName)

}
catch {Throw}
finally {
   Remove-Folder $WorkDir 
   Write-Host "Temporary working-directory => Removed"
}

 Write-Host "We are done here, thank you for flying with Werklijsthulpje!"



