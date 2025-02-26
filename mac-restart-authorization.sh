#!/bin/bash

# This script checks and updates the system.restart authorization rule on macOS systems.
# It ensures that users can restart the system without needing administrator credentials.
# Author - Mike Croucher
# Date - February 18, 2025

# Function to check the current status of the system.restart authorization rule
check_restart_permission() {
    /usr/bin/security authorizationdb read system.restart 2>/dev/null | grep -c "allow"
}

# Function to update the system.restart authorization rule
update_restart_permission() {
    if /usr/bin/security authorizationdb write system.restart allow; then
        echo "Successfully updated the system.restart authorization rule to allow."
    else
        echo "Failed to update the system.restart authorization rule." >&2
        exit 1
    fi
}

# Check the current status of the system.restart authorization rule
RESTART_PERMISSION=$(check_restart_permission)

if [ "$RESTART_PERMISSION" -ne 1 ]; then
    # Allow all users to restart the system without requiring admin credentials
    update_restart_permission
else
    echo "The system.restart authorization rule is already set to allow."
fi
