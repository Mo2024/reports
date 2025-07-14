param(
    [Parameter(Mandatory = $true)]
    [string]$pptFile,

    [Parameter(Mandatory = $true)]
    [string]$newFilePath
)

# Now you can use $pptFile and $newFilePath as variables in your script
Write-Host "PowerPoint file: $pptFile"
Write-Host "New Excel file path: $newFilePath"


# $pptFile = "C:\Users\mohamed.hasan\OneDrive - Bank ABC\Desktop\temp\ila Daily Stability Report - 13 July 2025.pptx"
# $newFilePath = "C:\Users\mohamed.hasan\OneDrive - Bank ABC\Desktop\temp\13 July 2025 - ila Daily GIT Stability Report Data.xlsx"
$msoTrue = -1

# Check if presentation exists
if (-not (Test-Path $pptFile)) {
    Write-Host "ERROR: PowerPoint file not found at path: $pptFile"
    exit
}

# Start PowerPoint application
$powerpoint = New-Object -ComObject PowerPoint.Application
$powerpoint.Visible = $msoTrue

# Open the presentation
$presentation = $powerpoint.Presentations.Open($pptFile)

# Helper function to replace only file path part
function Replace-LinkedFilePath($fullLink, $newPath) {
    # Regex to capture the file path ending with .xlsx (case-insensitive)
    # and everything after it (internal reference)
    if ($fullLink -match "^(.*?\.xlsx)(.*)$") {
        $oldPath = $matches[1]
        $internalRef = $matches[2]
        
        # Make sure new path uses backslashes and is absolute
        $newPathFixed = $newPath -replace '/', '\'
        
        # Return new full link
        return $newPathFixed + $internalRef
    } else {
        # If pattern not matched, just return original link
        return $fullLink
    }
}

# Loop through all slides and shapes
foreach ($slide in $presentation.Slides) {
    foreach ($shape in $slide.Shapes) {
        # Write-Host $shape.Type
        # Write-Host $shape.LinkFormat.SourceFullName
        if ($shape.Type -eq 10 -or $shape.Type -eq 3) {
            $oldLink = $shape.LinkFormat.SourceFullName
            Write-Host "Old link: $oldLink"

            $fileName = [System.IO.Path]::GetFileName($oldLink)

            # Skip if the linked file is Book1.xlsx
            if ($fileName -ieq "Book1.xlsx") {
                Write-Host "Skipping link to as it is related to EFTS graph: $fileName"
                continue
            }

            $newLink = Replace-LinkedFilePath $oldLink $newFilePath

            if ($newLink -ne $oldLink) {
                Write-Host "Updating link to: $newLink"
                $shape.LinkFormat.SourceFullName = $newLink
                $shape.LinkFormat.Update()
            }
        }

    }
}



# Save and close
$presentation.UpdateLinks()
$presentation.Save()
$presentation.Close()

# Quit PowerPoint
$powerpoint.Quit()

# Release COM objects
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($presentation) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($powerpoint) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

Write-Host "Links updated successfully."
