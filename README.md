# PowerShell File Processing Automation Tool

## Overview

`MVRS_To_Upload.ps1` is a PowerShell automation utility designed to
streamline a recurring file processing workflow.

The script automates the process of locating compressed files,
extracting contents, validating required data files, transferring output
files to a destination directory, performing cleanup operations, and
recording execution details through logging.

This project demonstrates practical systems administration automation
concepts including file workflow automation, validation checks, error
handling, and operational logging.

## Features

### Automated ZIP Processing

-   Searches a configured input directory for the most recent ZIP
    archive.
-   Moves the archive into a processing directory before extraction.

### File Extraction and Validation

-   Extracts ZIP contents using PowerShell archive tools.
-   Searches extracted contents for expected `.dat` files.
-   Validates that required files exist before continuing.

### Automated File Transfer

-   Renames validated `.dat` files to the expected output filename.
-   Moves processed files to the configured destination directory.

### Directory Management

-   Creates required directories if they do not already exist.
-   Maintains a predictable workflow structure.

### Cleanup Operations

-   Removes temporary extracted files and directories after processing.
-   Removes invalid archive files when expected output data is not
    found.

### Execution Logging

Records script activity including:

-   Directory creation
-   File processing steps
-   Successful operations
-   Errors and failures

## Technical Details

Built with:

-   PowerShell
-   Windows file system automation
-   Archive extraction workflows
-   File validation
-   Error handling
-   Execution logging

## Workflow

The script follows this process:

``` text
Input Directory
      |
      v
Locate Latest ZIP File
      |
      v
Move Archive to Processing Folder
      |
      v
Extract Archive Contents
      |
      v
Validate Required .dat File
      |
      v
Rename and Transfer File
      |
      v
Cleanup Temporary Files
      |
      v
Write Execution Logs
```

## Usage

### Requirements

-   Windows PowerShell 5.0+
-   Appropriate permissions to access configured directories

### Running the Script

Clone or download the repository.

Navigate to the script location:

``` powershell
cd path\to\repository
```

Execute:

``` powershell
.\MVRS_To_Upload.ps1 -Verbose
```

The `-Verbose` parameter displays additional runtime information during
execution.

## Configuration

The script uses variables at the beginning of the file to define
workflow locations:

``` powershell
$backupFolder
$extractFolder
$uploadFolder
$logFolder
$downloadsFolder
```

These values can be modified based on the environment where the script
is deployed.

## Engineering Considerations

This project was designed around several operational goals:

-   Reduce repetitive manual file handling
-   Create a repeatable workflow for recurring tasks
-   Improve troubleshooting through execution logs
-   Validate expected files before transfer
-   Reduce opportunities for user error

## Current Limitations

-   Workflow paths are configured directly within the script.
-   Designed for Windows PowerShell environments.
-   Assumes expected file types and workflow structure.
-   Logging is file-based rather than centralized.

## Future Improvements

Potential enhancements:

-   External configuration file support
-   Improved reporting output
-   Additional validation rules
-   More robust recovery workflows
-   Scheduled execution examples

## License

MIT License

Copyright (c) 2024 Kevin Cozart

Permission is hereby granted, free of charge, to any person obtaining a
copy of this software and associated documentation files (the
"Software"), to deal in the Software without restriction, including
without limitation the rights to use, copy, modify, merge, publish,
distribute, sublicense, and/or sell copies of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND.
