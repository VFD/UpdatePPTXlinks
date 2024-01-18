<#
.SYNOPSIS
    Mass changing links in PPTX file (Office).

.DESCRIPTION
    Check the Data File and update it as needed (Update-PPTXLinks.psd1).
    This is a PowerShell Data format, read the comments to update.
    
    Just start the script, better in the PS ISE.

    About Working and Destination folders:
    If they exists, all files inside are deleted to prevent crash.
    After running, copy your new PPTX files before restarting this script.

    It's recommanded to use the local temp folder.
    C:\temp
    Because sometime it is not subject to encryption.

.INPUTS
    None.

.OUTPUTS
    None.

.NOTES
    Author : Vincent Dubois
#>


# ----------------------------------------------------------------------
# TO DO:
#     Add change for slide mask

# Repertoire c:\temp
# 

# ----------------------------------------------------------------------

# ----------------------------------------------------------------------
# Load Config File and Initialise globals
# ----------------------------------------------------------------------
# Import the Data (Same Name as this script), if failed, stop the script
$PARAMETERS = Import-PowerShellDataFile $PSScriptRoot\$((Get-Item $MyInvocation.InvocationName).BaseName).psd1 -ErrorAction Stop

# Working and Destination folders
$WorkingFolder = $PARAMETERS.Folders.WorkingFolder
$DestinationFolder = $PARAMETERS.Folders.DestinationFolder

# Location of the slides files with the URL
$SlidesLocation = $PARAMETERS.Folders.SlidesLocation
$SlidesRelsLocation = $PARAMETERS.Folders.SlidesRelsLocation

# PPTX Folder TO CHANGE
#$FolderOfPPTX = -join( ([environment]::GetFolderPath('MyDocuments')), "\", $PARAMETERS.Folders.ProjectFolder, "\", $PARAMETERS.Folders.ProjectChangeLinks)
$FolderOfPPTX = -join($PARAMETERS.Folders.ProjectFolder, "\", $PARAMETERS.Folders.ProjectChangeLinks)

# Add Tab
$addTAB = "`t"

#Clear-Variable -Name ("PARAMETERS", "WorkingFolder", "DestinationFolder")

# ----------------------------------------------------------------------
# ----------------------------------------------------------------------
function New-WorkingDirectories {
    <#
    .SYNOPSIS
        Check for the working folders.
    .DESCRIPTION
        If exists empty them, else create them.
    #>

    [CmdletBinding()]
    param ()

    if ( ([IO.Directory]::Exists((Join-Path(Get-Location) $DestinationFolder)) )) {
        # Remove all inside if exist
        Remove-Item "$DestinationFolder\*" -Force -Recurse
        Write-Warning "$DestinationFolder  exist, now empty"
    }
    else {
        # Create
        Write-Host ($addTAB, "Create: ", $DestinationFolder)
        New-Item -ItemType 'Directory' -Name $DestinationFolder
    }
    #
    if ( ([IO.Directory]::Exists((Join-Path(Get-Location) $WorkingFolder)) )) {
        # Remove all inside if exist
        Remove-Item "$WorkingFolder\*" -Force -Recurse
        Write-Warning "$WorkingFolder exist, now empty"
    }
    else {
        # Create
        Write-Host ($addTAB, "Create: ", $WorkingFolder)
        New-Item -ItemType 'Directory' -Name $WorkingFolder
    }
}

# ----------------------------------------------------------------------
# ----------------------------------------------------------------------
function Expand-ZipFilePPTX {
    <#
    .SYNOPSIS
        UNZIP the renamed PPTX file in the Working folder.
    .DESCRIPTION
        UNZIP the renamed PPTX file in the Working folder.
        Delete the '.zip' after.
    #>

    [CmdletBinding()]
    param (
        # Name of the ZIP file to UNZIP.
        [Parameter(Mandatory)]
        [string]$File
    )
    Write-Host ($addTAB, $addTAB, "Expand Zip: ", $File)
    Expand-Archive -LiteralPath "$WorkingFolder\$File" -DestinationPath "$WorkingFolder" -Force
    Write-Host ($addTAB, $addTAB, "Remove Zip")
    Remove-Item "$WorkingFolder\*.zip"
}

# ----------------------------------------------------------------------
# ----------------------------------------------------------------------
function Compress-ZipFilePPTX {
    <#
    .SYNOPSIS
        ZIP all the files to recreate the PPTX File in the Destination Folder.
    .DESCRIPTION
        ZIP all the files to recreate the PPTX File in the Destination Folder.
    #>

    [CmdletBinding()]
    param (
        # Name of the ZIP file to ZIP.
        [Parameter(Mandatory)]
        [string]$File
    )
    Write-Host ($addTAB, $addTAB, "Compress to Zip: ", $File)
    # Zip and Move-Item
    Compress-Archive -Path "$WorkingFolder\*" -DestinationPath "$DestinationFolder\$File"
}

# ----------------------------------------------------------------------
# ----------------------------------------------------------------------
function Update-LinkURLinPPTX {
    <#
    .SYNOPSIS
        For each files, and for each URL replace URL.
        Need to be improved
    #>

    [CmdletBinding()]
    param ()

    # ----------------------------------------------------------------------
    # Go to location of 'Slides' (change text)
    # TO DO
    Set-Location $WorkingFolder\$SlidesLocation
    $files = Get-ChildItem -filter *.xml -name -file
    #$files
    # Work on files same code need to be a function
    foreach ($file in $files) {
	    Write-Host ($addTAB, $addTAB, "Working on file: ", $file)
	    $newFile = -join($file, ".new")
        # foreach URL to do
        foreach ($Link in $PARAMETERS.AllLinks) {
	        # replace
            (Get-Content $file).replace($Link.OldURL, $Link.NewURL) | Set-Content $newFile
            Remove-Item $file
            Rename-Item $newFile -NewName $file
        }
    }
    Set-Location $FolderOfPPTX
    #
    # ----------------------------------------------------------------------
    # Go to location of 'rels' (change URI in object)
    Set-Location $WorkingFolder\$SlidesRelsLocation
    # Get all .rels files
    $files = Get-ChildItem -filter *.rels -name -file
    # Work on files
    foreach ($file in $files) {
	    Write-Host ($addTAB, $addTAB, "Working on file: ", $file)
	    $newFile = -join($file, ".new")
        # foreach URL to do
        foreach ($Link in $PARAMETERS.AllLinks) {
	        # replace
            (Get-Content $file).replace($Link.OldURL, $Link.NewURL) | Set-Content $newFile
            Remove-Item $file
            Rename-Item $newFile -NewName $file
        }
        #
    }
    # return to 
    Set-Location $FolderOfPPTX
    # ----------------------------------------------------------------------
}

# ----------------------------------------------------------------------
# ----------------------------------------------------------------------
function Start-ChangeLinkInPPTX {
    <#
    .SYNOPSIS
        MAIN Function.
    #>

    [CmdletBinding()]
    param ()

    # Counting
    [int]$FileProcessed = 0
    
    # Current Location
    $currentLocation = Get-Location

    Set-Location $FolderOfPPTX

    Write-Host ("`nPPTX URL Changer")
    Write-Host ("================")
    # List all pptx in the folder
    $pptxFiles = Get-ChildItem -filter *.pptx -name -file

    # If there is PPTX file so we can (the array is not empty)
    if ( $null -ne $pptxFiles ) {
        # Create / Empty working and destination folders
        New-WorkingDirectories
        #
        Write-Host ($addTAB, "PPTX found, start processing.")
        # Main action
        foreach ($pptxFile in $pptxFiles) {
            Write-Host ($addTAB, "->", $pptxFile)
            # Copy the file in the working folder
            Write-Host ($addTAB, $addTAB, "Copy file in: ", $WorkingFolder)
            Copy-Item $pptxfile $WorkingFolder -force
            # Rename
            Write-Host ($addTAB, $addTAB, "Rename the file, add '.zip' at the end.")
            $newName = -join($pptxFile, '.zip')
            Rename-Item -Path "$WorkingFolder\$pptxFile" -NewName $newName
            # Unzip
            Expand-ZipFilePPTX -File $newName
            # Do It
            Update-LinkURLinPPTX
            # Zip
            Compress-ZipFilePPTX -File $newName
            Write-Host($addTAB, $addTAB, "Remove extention '.zip' from file.")
            # Rename
            Rename-Item "$DestinationFolder\$newName" -NewName $pptxFile
            # Empty Work
            Remove-Item "$WorkingFolder\*" -Force -Recurse
            $FileProcessed++
        }
    }
    else {
        Write-Warning "No PPTX found in $FolderOfPPTX. Nothing to do."
    }
    #
    if ( $FileProcessed ) {
        Write-Host ($addTAB, "End of work, $FileProcessed files changed, no more files found.")
        Write-Host ($addTAB, "Check the result in the destination folder: ", $FolderOfPPTX, "\", $DestinationFolder)
    }
    # Return to
    Set-Location $currentLocation
}

# ----------------------------------------------------------------------
# ----------------------------------------------------------------------
Start-ChangeLinkInPPTX
 
# -----[EOS]------------------------------------------------------------