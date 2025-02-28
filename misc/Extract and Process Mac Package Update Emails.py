import re
import os
import subprocess
import logging
from datetime import datetime
import argparse
from pathlib import Path

# Ensure the folder for log file exists
log_folder = os.path.expanduser('~/Desktop/Mac Package Updates/')
Path(log_folder).mkdir(parents=True, exist_ok=True)

# Set up logging to track script progress and errors
log_file = os.path.join(log_folder, 'script_log.txt')
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

logging.info("Started processing emails.")

# Parse input arguments to allow user customization
parser = argparse.ArgumentParser(description='Process raw email content for package updates.')
parser.add_argument('--folder', type=str, default="Mac Package Updates", help='The folder name to store results.')
args = parser.parse_args()

# Define file paths based on user input
desktop_path = os.path.expanduser(f"~/Desktop/{args.folder}/")
output_file_path = os.path.join(desktop_path, "Mac Package Updates.rtf")

# Ensure the folder exists
Path(desktop_path).mkdir(parents=True, exist_ok=True)

# Get current date in Day/Mon/Year format (e.g., 14/Feb/2025)
current_date = datetime.now().strftime("%d%b%Y")

# List of titles to exclude
excluded_titles = {
    "Adobe Acrobat Reader DC Continuous",
    "Adobe Acrobat DC Continuous",
    "Adobe Creative Cloud",
    "Apple macOS",
    "Apple macOS Sonoma",
    "Apple macOS Ventura",
    "Apple Safari",
    "Apple Xcode",
    "Blackmagic DaVinci Resolve",
    "Cisco Jabber",
    "DisplayLink Manager",
    "Docker Desktop",
    "Eclipse IDE",
    "JetBrains IntelliJ IDEA Ultimate",
    "JetBrains IntelliJ IDEA Community",
    "iLok License Manager",
    "JetBrains PyCharm Community",
    "JetBrains PyCharm Professional",
    "Microsoft Intune Company Portal",
    "Node.js",
    "Nudge",
    "Oracle SQLDeveloper",
    "Oracle VirtualBox",
    "QSR NVivo",
    "QSR NVivo 11",
    "QSR NVivo 12"
}

# AppleScript to search only in "Mac Package Updates" folder and return raw email content
apple_script = f'''
tell application "Mail"
    set senderEmail to "oit-macmgmt-sa@ohio.edu"
    set rawEmailContent to ""
    set targetMailbox to missing value
    set desktopPath to POSIX path of (path to desktop)
    set outputFolder to desktopPath & "{args.folder}/"

    -- Ensure the output directory exists
    do shell script "mkdir -p " & quoted form of outputFolder

    -- Loop through all accounts and mailboxes
    repeat with acc in every account
        repeat with mbx in (mailboxes of acc)
            if name of mbx is "{args.folder}" then
                set targetMailbox to mbx
                exit repeat
            end if
        end repeat

        if targetMailbox is not missing value then
            exit repeat
        end if
    end repeat

    -- Exit if mailbox is not found
    if targetMailbox is missing value then
        display dialog "Error: Mailbox '{args.folder}' not found in any account."
        return
    end if

    -- Extract messages from the mailbox
    set msgList to (messages of targetMailbox whose sender contains senderEmail)

    if (count of msgList) = 0 then
        display dialog "Error: No emails found from " & senderEmail & " in '{args.folder}'."
        return
    end if

    repeat with msg in msgList
        set emailSource to source of msg
        if emailSource is not missing value and emailSource is not "" then
            set rawEmailContent to rawEmailContent & emailSource & "\\n\\n"
            set read status of msg to true -- Mark as read
        end if
    end repeat

    return rawEmailContent
end tell
'''

# Run the AppleScript and capture the output
result = subprocess.run(["osascript", "-e", apple_script], capture_output=True, text=True)
if result.returncode != 0:
    logging.error(f"Error executing AppleScript: {result.stderr}")
    exit(1)

raw_email_content = result.stdout
logging.info("Successfully extracted emails.")

# Function to extract title and version
def extract_title_and_version(email_content, output_file, date):
    """Extracts title and version information from the raw email content and writes it to an RTF file."""
    # Regex pattern to match Title and Version
    title_version_pattern = r"Title:\s*(.*?)<br>Version:\s*(.*?)<br>"
    matches = re.findall(title_version_pattern, email_content)

    title_version_dict = {}

    # Compare versions and keep the newest version for each title
    for title, version in matches:
        if title in excluded_titles:
            continue  # Skip excluded titles

        if title in title_version_dict:
            existing_version = title_version_dict[title]
            if version > existing_version:  # Update to the newer version
                title_version_dict[title] = version
        else:
            title_version_dict[title] = version

    # RTF header and formatting
    rtf_content = r"{\rtf1\ansi\deff0 {\fonttbl {\f0\fswiss Helvetica;}}"
    rtf_content += r"\ The following applications have been added to Patch Management. \par"

    for title, version in title_version_dict.items():
        rtf_content += fr"\ {title} \ {version} \par"

    rtf_content += fr"\par\ Downloaded, repackaged, and signed on {date} by Mike \par"
    rtf_content += fr"\ Downloaded on {date} by Mike \par"
    rtf_content += "}"  # End of RTF document

    # Write to an RTF file
    with open(output_file, 'w', encoding='utf-8') as out_file:
        out_file.write(rtf_content)

# Extract data from raw email content
extract_title_and_version(raw_email_content, output_file_path, current_date)

logging.info("Finished processing emails. Titles and versions saved.")
