# Documentation for Script Usage

## Purpose

The main purpose of this script is to convert 25-second NVidia gameplay video clips into a format that can be shared on WhatsApp without exceeding file size limitations.

## Overview

This script processes and converts video files. It scans a specified input directory for videos of a certain type, extracts metadata, calculates target resolutions and aspect ratios, and uses HandBrakeCLI to convert the videos. Additionally, debug information can be output, and file attributes can be preserved.

## Prerequisites

To run this script, the following prerequisites must be met:

1. **PowerShell**: The script is written in PowerShell and requires a corresponding environment to execute.
2. **HandBrakeCLI**: The script uses HandBrakeCLI for video conversion. Ensure HandBrakeCLI is installed and the path to the executable is correctly specified.
3. **MediaToolkit**: The MediaToolkit library is used to extract video metadata. Ensure the DLL file is available and the path is correctly specified.

## Parameters

The script supports several parameters that specify the input and output paths, as well as various video conversion options.

### Mandatory Parameters

- `-InputPath <Path>`: Path to the source directory containing the video files to be processed.
- `-OutputPath <Path>`: Path to the target directory where the converted video files will be saved.

### Optional Parameters

- `-FileType <Extension>`: The file type to process (default: `.mp4`).
- `-HandbrakePath <Path>`: Path to HandBrakeCLI.exe (default: `.\HandBrakeCLI.exe`).
- `-Height <Number>`: Target height of the video in pixels.
- `-Ratio <String>`: Target aspect ratio in the form 'Width:Height' (e.g., '21:9', '4:3').
- `-MediaToolkitPath <Path>`: Path to MediaToolkit.dll.
- `-DebugMode`: Enables debug mode, which outputs additional information.
- `-NoVideoProcessing`: Collects and outputs processing data without converting videos.
- `-Help`: Displays a help message.

## Usage

### Display Help

To display help, run the script without parameters or use the `-Help` parameter:
powershell
.\script.ps1 -Help

### Example Call

A typical call to the script might look like this:

powershell
.\script.ps1 -InputPath 'C:\video\in' -OutputPath 'C:\video\out' -FileType '.mp4' -Height 1080 -Ratio '21:9' -DebugMode

This command processes all `.mp4` files in the directory `C:\video\in`, converts them to a height of 1080 pixels and an aspect ratio of 21:9, and saves the converted files in the directory `C:\video\out`. Additionally, debug information is output.

## Fixed HandBrake Parameters

The script uses the following fixed HandBrake parameters for video conversion:

- `-e nvenc_h264`: Uses NVENC H.264 encoder.
- `-b 17500`: Sets the video bitrate to 17500 kbps.
- `-B 128`: Sets the audio bitrate to 128 kbps.
- `-r 60`: Sets the frame rate to 60 fps.
- `--auto-anamorphic`: Enables automatic anamorphic video.
- `--lapsharp=medium`: Sets the lapsharp filter to medium.
- `--encoder-profile high`: Sets the encoder profile to high.
- `--encoder-level 5.2`: Sets the encoder level to 5.2.
- `--encoder-preset slow`: Sets the encoder preset to slow.

## Functionality

### Parameter Validation

The script first checks if the required parameters `InputPath` and `OutputPath` are provided. If not, an error message is displayed and the help message is shown.

### Loading the MediaToolkit Library

The MediaToolkit library is loaded to extract video metadata. If loading fails, an error message is displayed.

### Processing Video Files

The script scans the input directory for video files of the specified type. For each file not already present in the output directory, the following steps are performed:

1. **Extract Metadata**: Video metadata is extracted using MediaToolkit.
2. **Calculate Target Resolution and Aspect Ratio**: The target resolution and aspect ratio are calculated based on the specified parameters.
3. **Create HandBrakeCLI Command**: A HandBrakeCLI command is created to convert the video.
4. **Preserve File Attributes**: The original file attributes (creation time, modification time, access time) are transferred to the converted file.

### Debug Mode

When debug mode is enabled, additional information is output, including video metadata, calculated resolutions and aspect ratios, and executed HandBrakeCLI commands.

## Conclusion

This script provides a flexible way to batch convert video files with detailed control over resolution and aspect ratio. By using HandBrakeCLI and MediaToolkit, powerful conversion and metadata extraction functionalities are utilized.



