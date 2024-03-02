# Post-Offboarding Automation Script

## Overview

This PowerShell script automates the post-offboarding process for user accounts in a Microsoft 365 (Office 365) environment. The script focuses on handling tasks related to disabling accounts, archiving mailboxes, and notifying relevant stakeholders. It utilizes AzureAD and Exchange Online PowerShell modules for managing user accounts and Exchange Online features.

## Configuration Settings

The script begins with a settings section, where various global variables are defined. These settings include paths for storing audit files, master user data, backups, and secure key locations. Additionally, email configurations such as SMTP server details, recipient addresses, and email body templates are specified. Adjust these settings according to your organization's environment.

## Functions

### 1. Load-Module

This function checks if a PowerShell module is already imported. If not, it attempts to import the module, install it from an online gallery, or exits if the module is unavailable.

### 2. Transcribe

Initiates a transcript logging session, capturing the script's output for auditing purposes. The transcript file includes the username, computer name, and timestamp.

### 3. Check

Queries Azure AD for disabled user accounts and Exchange Online for shared mailboxes, identifying accounts to be processed. Generates CSV files containing relevant information for offboarding.

### 4. Load-Changes

Loads user data from the master file and checks its size. If the file size is outside the specified range, an email notification is sent to alert the team to check the master file.

### 5. Date-Check

Compares the specified date for action with the current date to determine if the action should be taken.

### 6. iniNotify

Sends an initial notification to a user's supervisor about upcoming changes to the mailbox settings, providing an overview of the changes and instructions for requesting access extensions.

### 7. Notify

Notifies a user's supervisor about upcoming changes to the mailbox settings, indicating actions like removing access, archiving the mailbox, and bouncing back emails.

### 8. Extend

Extends the forward for another 30 days, changing the action to "iniNOTIFY."

### 9. Archive

Disables access to the mailbox, removes permissions, disables email forwarding, appends "ARCHIVE" to the mailbox's last name, and renames the mailbox. Finally, sets the action to "COMPLETED."

### 10. Save-Changes

Saves changes to the master file, exports a backup file, and deletes backups older than 90 days.

### 11. Main Loop

Loads required modules, connects to Azure AD and Exchange Online, loads user data, and iterates through each user to determine the action based on the "ActionToTake" attribute.

## Usage

1. Adjust the configuration settings in the script according to your organization's environment.
2. Run the script in a PowerShell environment that has the necessary permissions to access Azure AD and Exchange Online.

## Notes

- Ensure that required modules (AzureAD, PSExcel, ActiveDirectory) are installed.
- The script uses secure key storage for authentication.
- Transcript logging provides an audit trail for script execution.
- Notifications are sent to supervisors about mailbox changes.
- Archiving includes disabling access, removing permissions, and renaming the mailbox.

**Disclaimer: Use this script responsibly and thoroughly test in a controlled environment before applying it in a production setting.**
