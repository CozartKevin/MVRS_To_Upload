# MVRS_To_Upload.ps1

## Overview

**MVRS_To_Upload.ps1** is a PowerShell script designed to automate the process of transferring `.dat` files from Temetra to MOM Software. This script identifies the latest ZIP file in a user's Downloads folder, extracts its contents, checks for the presence of a `.dat` file, and moves the `.dat` file to a specified upload folder. The script also performs cleanup tasks to maintain a clean working environment.

## Features

- **Automated ZIP Handling:** Identifies and processes the most recent ZIP file in the Downloads folder.
- **Directory Management:** Ensures all necessary directories exist or creates them if missing.
- **File Extraction and Verification:** Extracts ZIP contents and verifies the presence of a `.dat` file.
- **File Transfer:** Renames and moves the `.dat` file to the designated upload folder.
- **Comprehensive Cleanup:** Removes unnecessary files and directories post-processing.
- **Robust Logging:** Logs all actions and errors for easy tracking and debugging.

## Usage

### Prerequisites

- PowerShell 5.0 or higher
- Appropriate permissions to access and modify specified directories and files.

### Script Execution

1. **Download the Script:**
   - Clone or download the script from the GitHub repository.

2. **Set Up Environment:**
   - Ensure the following directories are correctly set up:
     - `C:\Scripts\TransferredZips`
     - `C:\MVRS\XFER\Upload`
     - `C:\Scripts\Logs`

3. **Run the Script:**
   - Open PowerShell with administrative privileges.
   - Navigate to the directory containing `MVRS_To_Upload.ps1`.
   - Execute the script using the following command, including the `-Verbose` parameter for detailed logging:
     ```powershell
     .\MVRS_To_Upload.ps1 -Verbose
     ```

### Parameters

- **$backupFolder**: Path to the backup folder where ZIP files are moved.
- **$extractFolder**: Path to the folder where ZIP files are extracted.
- **$uploadFolder**: Path to the upload folder where `.dat` files are moved.
- **$uploadFile**: Name of the upload file (`UPLOAD.dat`).
- **$logFolder**: Path to the folder where logs are stored.
- **$downloadsFolder**: Path to the Downloads folder of the current user.

### Script Details

The script performs the following steps:

1. **Define Paths**: Sets up paths for backup, extraction, upload, and logging.
2. **Ensure Directories**: Checks if necessary directories exist and creates them if not.
3. **Identify ZIP File**: Finds the most recent ZIP file in the Downloads folder.
4. **Move ZIP File**: Moves the identified ZIP file to the backup folder.
5. **Extract Contents**: Extracts the contents of the ZIP file to the extraction folder.
6. **Check for `.dat` File**: Searches for a `.dat` file in the extracted contents.
7. **Move `.dat` File**: Renames and moves the `.dat` file to the upload folder.
8. **Cleanup**: Removes all non-.zip, non-.dat, and non-.bat files and directories.
9. **Specific ZIP Removal**: If no `.dat` file is found, the script removes the ZIP file from the backup folder.
10. **Logging**: Logs all actions and errors throughout the script execution.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your improvements or bug fixes.

## License

MIT License

Copyright (c) [2024] Kevin Cozart.

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

### Note:
This software must include the reference to the original author, Kevin Cozart at Core Utilities Inc., when used, distributed, or modified.

---
[LinkedIn](https://www.linkedin.com/in/Cozartkevin)  
[GitHub](https://www.github.com/CozartKevin)
