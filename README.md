# CalendarReviewerSync

## Overview
**CalendarReviewerSync** is a PowerShell script designed to automatically grant a designated **Active Directory group** **Reviewer permissions** on the calendars of all its members. This ensures that group members have consistent access to each other's calendars without manual intervention.

## Features
- 🚀 **Automates calendar permissions** for all members of a specified group.
- 🔄 Supports both **on-premises Exchange** and **Exchange Online (Office 365)**.
- 🌍 Dynamically detects the user's default calendar folder name, even if localized.
- ✅ Handles both granting and updating of **Reviewer** permissions.
- ⚠️ Includes error handling to detect missing mailboxes or permission conflicts.

## Prerequisites

### 1. Required PowerShell Modules
Ensure the following PowerShell modules are installed:

- **Active Directory Module** (For retrieving group members)
- **Exchange PowerShell Module** (For managing mailbox permissions)

For **Exchange Online**, install the required module using:

```Install-Module ExchangeOnlineManagement```

### 2. Permissions Required

The user running the script must have:

  - Exchange Administrator or Mailbox Permissions Admin rights.
  - Appropriate privileges to read group memberships in Active Directory.

### 3. Connectivity

  For on-premises Exchange, the script should run from a system with access to Exchange Management Shell.
  For Exchange Online, connect using:

    Connect-ExchangeOnline

## Installation

Clone the repository:

    git clone https://github.com/yourusername/CalendarReviewerSync.git

Navigate to the project directory:

    cd CalendarReviewerSync

Open the script (.ps1) in PowerShell or VS Code.

## Usage

  Set the group name in the script:

    $groupName = "IT-Support"  # Change to your desired group name

Run the script in PowerShell:

    .\CalendarReviewerSync.ps1

If using Exchange Online, ensure you've connected first:

    Connect-ExchangeOnline

## How It Works

- Retrieves all members of the specified Active Directory group.
- Identifies if each user has an on-premises or cloud mailbox.
- Detects the user's calendar folder name dynamically.
- Assigns Reviewer permissions to the specified group for each user's calendar.
- Handles errors and updates permissions if they already exist.

Example Output

    Processing user1@example.com...
    Granted permissions for user1@example.com
    Processing user2@example.com...
    Updated permissions for user2@example.com
    Processing user3@example.com...
    Warning: Mailbox not found for user3@example.com

## Troubleshooting

If you encounter permission errors, ensure the account running the script has Exchange admin rights.
If mailboxes are missing, check whether they exist using 

    Get-Mailbox -Identity user@example.com
If the script fails to find the calendar folder, verify the mailbox statistics manually:

    Get-MailboxFolderStatistics -Identity user@example.com -FolderScope Calendar
