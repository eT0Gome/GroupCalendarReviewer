# Before running this script, if you have cloud mailboxes, please run:
# Connect-ExchangeOnline -Prefix EXO

# Import the Active Directory module to retrieve group members
Import-Module ActiveDirectory

# Define the group name (replace with your group, e.g., "IT-Support")
$groupName = "IT-Support"

# Get the members of the group and their email addresses
$members = Get-ADGroupMember -Identity $groupName | Get-ADUser -Properties mail | Select-Object -ExpandProperty mail

# Function to dynamically retrieve the calendar folder name in the user’s language (for on-premises)
function Get-CalendarFolderName {
    param (
        [string]$mailbox,
        [string]$environment  # "onprem" or "cloud"
    )
    try {
        if ($environment -eq "onprem") {
            $stats = Get-MailboxFolderStatistics -Identity $mailbox -FolderScope Calendar
        } elseif ($environment -eq "cloud") {
            $stats = Get-EXOMailboxFolderStatistics -Identity $mailbox -FolderScope Calendar
        }
        $calendarFolder = $stats | Where-Object { $_.FolderType -eq "Calendar" } | Select-Object -First 1
        return $calendarFolder.Name
    } catch {
        Write-Warning "Could not retrieve calendar folder for $mailbox $($_.Exception.Message)"
        return $null
    }
}

# Process each member’s mailbox
foreach ($member in $members) {
    Write-Host "Processing $member..."

    # Determine if the mailbox is on-premises or in the cloud
    $localMailbox = Get-Mailbox -Identity $member -ErrorAction SilentlyContinue
    if ($localMailbox) {
        $environment = "onprem"
    } else {
        $cloudMailbox = Get-EXOMailbox -Identity $member -ErrorAction SilentlyContinue
        if ($cloudMailbox) {
            $environment = "cloud"
        } else {
            Write-Warning "Mailbox not found for $member"
            continue
        }
    }

    # For cloud mailboxes, use the invariant "Calendar" folder name.
    if ($environment -eq "cloud") {
        $calendarName = "Calendar"
    } else {
        $calendarName = Get-CalendarFolderName -mailbox $member -environment $environment
    }

    if ($calendarName) {
        # Construct the calendar identity (e.g., "user@domain.com:\Calendar")
        $calendarIdentity = "${member}:\$calendarName"

        # Attempt to grant or update permissions
        try {
            if ($environment -eq "onprem") {
                Add-MailboxFolderPermission -Identity $calendarIdentity -User $groupName -AccessRights Reviewer -ErrorAction Stop
            } else {
                Add-EXOMailboxFolderPermission -Identity $calendarIdentity -User $groupName -AccessRights Reviewer -ErrorAction Stop
            }
            Write-Host "Granted permissions for $member"
        } catch {
            if ($_.Exception.Message -like "*already exists*") {
                # If permission exists, update it
                if ($environment -eq "onprem") {
                    Set-MailboxFolderPermission -Identity $calendarIdentity -User $groupName -AccessRights Reviewer
                } else {
                    Set-EXOMailboxFolderPermission -Identity $calendarIdentity -User $groupName -AccessRights Reviewer
                }
                Write-Host "Updated permissions for $member"
            } else {
                Write-Warning "Error for $member $($_.Exception.Message)"
            }
        }
    } else {
        Write-Warning "Could not find calendar folder for $member"
    }
}
