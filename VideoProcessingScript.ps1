# Define parameters
param(
    [Parameter(Mandatory=$false)]
    [string]$InputPath,
    
    [Parameter(Mandatory=$false)]
    [string]$OutputPath,
    
    [Parameter(Mandatory=$false)]
    [string]$FileType = ".mp4",

    [Parameter(Mandatory=$false)]
    [string]$HandbrakePath = ".\HandBrakeCLI.exe",

    [Parameter(Mandatory=$false)]
    [int]$Height,

    [Parameter(Mandatory=$false)]
    [string]$Ratio,

    [Parameter(Mandatory=$false)]
    [string]$MediaToolkitPath = "C:\Users\janko\.nuget\packages\mediatoolkit\1.1.0.1\lib\net20\MediaToolkit.dll",

    [Parameter(Mandatory=$false)]
    [switch]$DebugMode,

    [Parameter(Mandatory=$false)]
    [switch]$NoVideoProcessing,

    [Parameter(Mandatory=$false)]
    [switch]$Help
)

# Function to display help
function Show-Help {
    Write-Host "Usage: .\script.ps1 -InputPath <Path> -OutputPath <Path> [Options]"
    Write-Host ""
    Write-Host "Parameters:"
    Write-Host "  -InputPath <Path>       Required. Path to the source directory."
    Write-Host "  -OutputPath <Path>      Required. Path to the target directory."
    Write-Host "  -FileType <Extension>   Optional. File type to process (default: .mp4)."
    Write-Host "  -HandbrakePath <Path>   Optional. Path to HandBrakeCLI.exe (default: .\HandBrakeCLI.exe)."
    Write-Host "  -Height <Number>        Optional. Target height of the video in pixels."
    Write-Host "  -Ratio <String>         Optional. Target aspect ratio in the form 'Width:Height' (e.g., '21:9', '4:3')."
    Write-Host "  -MediaToolkitPath <Path> Optional. Path to MediaToolkit.dll."
    Write-Host "  -DebugMode              Optional. Enables debug mode."
    Write-Host "  -NoVideoProcessing      Optional. Collects and outputs processing data without converting videos."
    Write-Host "  -Help                   Displays this help message."
    Write-Host ""
    Write-Host "Example:"
    Write-Host "  .\script.ps1 -InputPath 'c:\video\in folder' -OutputPath 'c:\video\out' -FileType '.mp4' -Height 1080 -Ratio '21:9' -DebugMode"
    exit
}


# Check if script was called without parameters
if ($PSCmdlet.MyInvocation.BoundParameters.Count -eq 0) {
    Show-Help
}

# Display help if -Help is used
if ($Help) {
    Show-Help
}

# Check if required parameters are present
if (-not $InputPath -or -not $OutputPath) {
    Write-Host "Error: InputPath and OutputPath are required parameters."
    Show-Help
}

# Load the MediaToolkit library
try {
    Add-Type -Path $MediaToolkitPath
}
catch {
    Write-Host "Error loading MediaToolkit library: $($_.Exception.Message)"
    Write-Host "Please check the path to MediaToolkit.dll and provide it using the -MediaToolkitPath parameter."
    exit
}

# Function to extract video metadata
function Get-VideoMetadata {
    param (
        [Parameter(Mandatory=$true)]
        [string]$VideoFilePath
    )

    try {
        # Open file with MediaToolkit
        $inputFile = New-Object MediaToolkit.Model.MediaFile -ArgumentList $VideoFilePath
        $engine = New-Object MediaToolkit.Engine

        # Retrieve metadata
        $engine.GetMetadata($inputFile)

        # Extract video data
        $videoData = $inputFile.Metadata.VideoData

        # Extract resolution and aspect ratio
        $width = $videoData.FrameSize.Split('x')[0]
        $height = $videoData.FrameSize.Split('x')[1]
        $aspectRatio = [math]::Round([double]$width / [double]$height, 2)

        return [PSCustomObject]@{
            Resolution   = "$width x $height"
            AspectRatio  = "$aspectRatio:1"
            OriginalWidth = [int]$width
            OriginalHeight = [int]$height
        }
    } catch {
        Write-Error "Error retrieving metadata: $($_.Exception.Message)"
        return $null
    }
}

# Function to calculate crop values
function Get-VideoCropValue {
    param (
        [int]$OriginalWidth,
        [int]$OriginalHeight,
        [int]$TargetHeight,
        [string]$TargetRatio
    )

    # Initial values
    $CropWidth = 0
    $CropHeight = 0
    $TargetWidth = $OriginalWidth

    if ($TargetRatio) {
        # Parse the TargetRatio
        $ratioParts = $TargetRatio -split ':'
        $TargetRatioWidth = [int]$ratioParts[0]
        $TargetRatioHeight = [int]$ratioParts[1]

        # Calculate multipliers
        $TargetMultiplier = $TargetRatioWidth / $TargetRatioHeight
        $SourceMultiplier = $OriginalWidth / $OriginalHeight

        if ($SourceMultiplier -gt $TargetMultiplier) {
            # Calculate TargetWidth
            $TargetWidth = ($OriginalHeight / $TargetRatioHeight) * $TargetRatioWidth
            # Calculate CropWidth
            $CropWidth = $OriginalWidth - $TargetWidth
            $CropHeight = 0
        }
    }

    return [PSCustomObject]@{
        CropWidth   = $CropWidth
        CropHeight  = $CropHeight
    }
}

# Function to retrieve target video resolution
function Get-TargetVideoDimension {
    param (
        [int]$OriginalWidth,
        [int]$OriginalHeight,
        [int]$TargetHeight,
        [string]$TargetRatio
    )

    # Initial values
    $TargetWidth = $OriginalWidth

    if ($TargetRatio) {
        # Parse the TargetRatio
        $ratioParts = $TargetRatio -split ':'
        $TargetRatioWidth = [int]$ratioParts[0]
        $TargetRatioHeight = [int]$ratioParts[1]

        # Calculate multipliers
        $TargetMultiplier = $TargetRatioWidth / $TargetRatioHeight
        $SourceMultiplier = $OriginalWidth / $OriginalHeight

        # check height
        if ($Height -eq 0) {
            $TargetHeight = $OriginalHeight
        }

        if ($SourceMultiplier -gt $TargetMultiplier) {
            # Calculate TargetWidth
            $TargetWidth = ($TargetHeight / $TargetRatioHeight) * $TargetRatioWidth
        }
    } elseif (-not $TargetRatio -and -not $Height) {
        $TargetWidth = $OriginalWidth
        $TargetHeight = $OriginalHeight
    } else {
        # Calculate TargetWidth based on TargetHeight
        $TargetWidth = ($OriginalWidth / $OriginalHeight) * $TargetHeight
    }

    return [PSCustomObject]@{
        TargetWidth  = [int]$TargetWidth
        TargetHeight = $TargetHeight
    }
}

# Function to retrieve file attributes
function Get-FileAttributes {
    param (
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)]
        [string]$FilePath
    )
    try {
        $item = Get-Item -LiteralPath $FilePath -ErrorAction Stop
        $attributes = $item | Select-Object -Property *
        $output = [PSCustomObject]@{
            FullName = $item.FullName
            Attributes = $attributes
        }
        Write-Output $output
    }
    catch {
        Write-Warning "Error retrieving attributes for '$FilePath': $($_.Exception.Message)"
    }
}

# Create file type filter
$fileFilter = "*" + $FileType

# Parameters for the HandBrake call
$params = "-e nvenc_h264 -b 17500 -B 128 -r 60 --auto-anamorphic --lapsharp=medium --encoder-profile high --encoder-level 5.2 --encoder-preset slow"

# List of all files in the target directory
$filesInTarget = (Get-ChildItem -File -Path $OutputPath | Where-Object {$_.Extension -eq $FileType}).Basename

# Number of files to process
$totalFilesToProcess = (Get-ChildItem -Path $InputPath -Filter $filefilter).Count

# Count the new files created
$processedFilesCount = 0

# Initialize an array to hold the data for the table
$tableData = @()

# Iterate over all files in the InputPath
Get-ChildItem -Path $InputPath -Filter $fileFilter | ForEach-Object {
    $currentFileBasename = $_.BaseName
    # Check if the basename of the current file is not in $FilesInTarget
    if (-not ($filesInTarget -contains $currentFileBasename)) {
        if ($DebugMode) {
            Write-Host "Processing file $currentFileBasename..."
        }
        $currentFile = $_.FullName
        $videoMetadata = Get-VideoMetadata -VideoFilePath "$currentFile"
        if (-not $videoMetadata) {
            Write-Error "Error retrieving metadata for $currentFile."
            continue
        }
        $cropValues = Get-VideoCropValue -OriginalWidth $videoMetadata.OriginalWidth -OriginalHeight $videoMetadata.OriginalHeight -TargetHeight $Height -TargetRatio $Ratio
        $targetResolution = Get-TargetVideoDimension -OriginalWidth $videoMetadata.OriginalWidth -OriginalHeight $videoMetadata.OriginalHeight -TargetHeight $Height -TargetRatio $Ratio

        if ($DebugMode) {            
            # Collect data for the table
            $tableData += [PSCustomObject]@{
                FileName          = $_.Name
                Resolution        = $videoMetadata.Resolution
                AspectRatio       = $videoMetadata.AspectRatio
                OriginalWidth     = $videoMetadata.OriginalWidth
                OriginalHeight    = $videoMetadata.OriginalHeight
                TargetWidth       = $targetResolution.TargetWidth
                TargetHeight      = $targetResolution.TargetHeight
                CropWidth         = $cropValues.CropWidth
                CropHeight        = $cropValues.CropHeight
            }
        }

        # Save file attributes
        $fileCreationTime = $_.CreationTime
        $fileLastWriteTime = $_.LastWriteTime
        $fileLastAccessTime = $_.LastAccessTime

        # Create the output file path
        $outputFile = Join-Path -Path $OutputPath -ChildPath $_.Name

        # Calculate crop values for HandBrakeCLI
        $cropTop = [int]($cropValues.CropHeight / 2)
        $cropBottom = $cropTop
        $cropLeft = [int]($cropValues.CropWidth / 2)
        $cropRight = $cropLeft

        # Replace placeholders in the command
        $command = "$HandbrakePath -i `"$($_.FullName)`" -o `"$outputFile`" $params --width $($targetResolution.TargetWidth) --height $($targetResolution.TargetHeight) --crop ${cropTop}:${cropBottom}:${cropLeft}:${cropRight}"

        # Output the value of the $command variable if debug mode is enabled
        if ($DebugMode) {
            Write-Host "Executing command: $command"
        }
        
        if (-not $NoVideoProcessing){
            # Execute the command
            Invoke-Expression $command

            # Increment the counter for processed files
            $processedFilesCount++

            # Copy file attributes
            try {
                $outputItem = Get-Item $outputFile
                $outputItem.CreationTime = $fileCreationTime
                $outputItem.LastWriteTime = $fileLastWriteTime
                $outputItem.LastAccessTime = $fileLastAccessTime
                $outputItem | Out-Null # Discard the result to suppress output
            } catch {
                Write-Warning "Error setting file attributes for '$outputFile': $($_.Exception.Message)"
            }
        }        
    }
}

# Output the number of processed files if debug mode is enabled
if ($DebugMode) {
    $tableData | Format-Table -AutoSize
}
