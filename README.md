Excel Data Refresh Automation

This is a PowerShell automation script designed to bulk refresh data connections in Excel workbooks. This script automatically processes multiple Excel files, refreshing their data connections while implementing robust error handling and logging mechanisms.
Features

Automated refresh of data connections in Excel workbooks
Comprehensive logging system
Robust error handling with retry logic
File locking detection
Single instance enforcement
Automatic cleanup of Excel processes
Resource management for COM objects

Prerequisites

Windows PowerShell 5.1 or later
Microsoft Excel installed
OpenXML SDK V2.5
Appropriate permissions to:

Access and modify target Excel files
Create and write to log files
Start and stop Excel processes



Installation

Install OpenXML SDK:

Download from Microsoft's official website
Install to your preferred location
Note the installation path for configuration


Configure the script:

Open ExcelFileProcessor.ps1 in a text editor
Replace the following placeholders:
powershellCopy[LOG_FILE_PATH]      # Path where log files will be stored
[EXCEL_FILES_PATH]   # Directory containing Excel files to process
[OPENXML_SDK_PATH]   # Path to OpenXML SDK installation




Usage
Basic Usage

Open PowerShell with administrator privileges
Navigate to the script directory
Run the script:
powershellCopy.\ExcelFileProcessor.ps1


Monitoring

Check the log file at [LOG_FILE_PATH] for:

Processing status
Error messages
Execution times
Completion status



Configuration
Adjustable Timeouts
The script includes several configurable constants:
powershellCopy$WAIT_FOR_EXCEL_TIMEOUT_SECONDS = 40  # Max time to wait for Excel to be ready
$REFRESH_DELAY_SECONDS = 30           # Delay before refreshing data
$REFRESH_WAIT_SECONDS = 40            # Wait time after initiating refresh
$SAVE_CLOSE_DELAY_SECONDS = 30        # Delay between save and close operations
$ITERATION_DELAY_SECONDS = 15         # Delay between processing different files
$MAX_RETRY_ATTEMPTS = 3               # Maximum number of retry attempts
$RETRY_DELAY_SECONDS = 10             # Delay between retry attempts
Modify these values based on your system's performance and requirements.
Troubleshooting
Common Issues and Solutions

Script fails to start Excel

Verify Excel is properly installed
Check if Excel is running in the background
Ensure proper permissions


"File in use" errors

Check if files are open in Excel
Verify no other processes are accessing the files
Check network connectivity for networked files


OpenXML SDK errors

Verify SDK installation
Check path configuration
Ensure proper version (V2.5) is installed


Permission errors

Run PowerShell as administrator
Check file and folder permissions
Verify user account permissions



Best Practices

Before Running

Close all Excel instances
Ensure sufficient disk space for logs
Backup important Excel files


Monitoring

Check logs regularly
Monitor system resource usage
Verify file updates


Maintenance

Regularly clear old log files
Update timeout values as needed
Test after system changes



Contributing

Fork the repository
Create a feature branch
Commit your changes
Push to the branch
Create a Pull Request

Author
David Kapindula
Version History

1.0: Initial Release

Base functionality
Error handling
Logging system



Acknowledgments

Microsoft OpenXML SDK documentation
PowerShell documentation
