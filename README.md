Excel Data Refresh Automation

This is a PowerShell automation script designed to bulk refresh data connections in Excel workbooks.This script runs on a company server and is scheduled to execute automatically using Windows Task Scheduler, ensuring regular data updates without manual intervention.
Features

Automated refresh of data connections in Excel workbooks
Comprehensive logging system
Robust error handling with retry logic
File locking detection
Single instance enforcement
Automatic cleanup of Excel processes
Resource management for COM objects
Scheduled execution via Task Scheduler

Prerequisites

Windows Server (2016 or later)
Windows PowerShell 5.1 or later
Microsoft Excel installed on the server
OpenXML SDK V2.5
Windows Task Scheduler access
Appropriate permissions to:

Access and modify target Excel files
Create and write to log files
Start and stop Excel processes
Create and modify scheduled tasks



Installation

Install OpenXML SDK:

Download from Microsoft's official website
Install to your preferred location
Note the installation path for configuration


Configure the script:

Open ExcelFileProcessor.ps1 in a text editor
Replace the following placeholders:

[LOG_FILE_PATH]      # Path where log files will be stored
[EXCEL_FILES_PATH]   # Directory containing Excel files to process
[OPENXML_SDK_PATH]   # Path to OpenXML SDK installation




Task Scheduler Setup
Detailed Configuration Steps

Open Task Scheduler on the server

Press Win + R
Type "taskschd.msc" and press Enter


Create a new task:

In the right pane, click "Create Task"
General tab:

Name: "Excel Data Refresh"
Description: "Automated Excel data connection refresh"
Select "Run whether user is logged on or not"
Check "Run with highest privileges"
Configure for: Windows Server




Set up triggers:

Click the Triggers tab
Click "New"
Settings:

Begin the task: On a schedule
Daily
Start time: [Set your preferred time, typically during off-hours]
Recur every: 1 day
Check "Enabled"


Advanced settings:

Check "Stop task if it runs longer than: [set appropriate timeout]"
Check "Repeat task every: [your interval]" if needed




Configure actions:

Click the Actions tab
Click "New"
Action: Start a program
Settings:
CopyProgram/script: powershell.exe
Arguments: -ExecutionPolicy Bypass -File "[full_path_to_script]\ExcelFileProcessor.ps1"
Start in: [script_directory]



Set conditions:

Click the Conditions tab
Power:

Uncheck "Start the task only if the computer is on AC power"
Uncheck "Stop if the computer switches to battery power"


Network:

Check "Start only if the following network connection is available"
Select "Any connection"




Configure settings:

Click the Settings tab
Check "Allow task to be run on demand"
Check "Run task as soon as possible after a scheduled start is missed"
Check "If the task fails, restart every:"

Set to 5 minutes
Attempt to restart up to: 3 times


Check "Stop the task if it runs longer than: [set appropriate timeout]"
Check "If the running task does not end when requested, force it to stop"



Usage
Manual Execution (if needed)

Open PowerShell with administrator privileges
Navigate to the script directory
Run the script:
powershellCopy.\ExcelFileProcessor.ps1


Scheduled Execution

The script runs automatically according to the Task Scheduler configuration
Check Task Scheduler history for execution status and results

Monitoring

Check the log file at [LOG_FILE_PATH] for:

Processing status
Error messages
Execution times
Completion status


Monitor Task Scheduler history:

Open Task Scheduler
Select your task
Check the History tab for execution results


Server-Specific Considerations
Performance

Schedule runs during off-peak hours
Monitor server resource usage
Adjust timeout values based on server performance

Security

Use service account with minimum required permissions
Secure log file locations
Implement network share access controls
Regular security audit of service account permissions

Maintenance

Regular review of Task Scheduler history
Monitor disk space for logs
Verify service account permissions
Test after server updates or patches

Version History

1.1: Added Task Scheduler implementation

Detailed scheduling configuration
Server deployment documentation


1.0: Initial Release

Base functionality
Error handling
Logging system
