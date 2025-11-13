#!/bin/bash

# Bash script to remove Excel developer add-in temp files
# Run this script from Git Bash

echo "Cleaning up Excel developer add-ins..."
echo ""

# Colors for output (if supported)
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
CYAN='\033[0;36m'
GRAY='\033[0;37m'
NC='\033[0m' # No Color

# Function to safely remove directory
remove_directory() {
    local path="$1"
    if [ -d "$path" ]; then
        if rm -rf "$path" 2>/dev/null; then
            echo -e "  ${GREEN}[OK]${NC} Removed directory: $path"
        else
            echo -e "  ${YELLOW}[WARNING]${NC} Could not remove directory: $path"
            echo -e "    ${YELLOW}You may need to close Excel first${NC}"
        fi
    else
        echo -e "  ${GRAY}[SKIP]${NC} Directory not found: $path"
    fi
}

echo -e "${CYAN}=== Cleaning Excel Add-in Directories ===${NC}"
echo ""

# Get user profile paths
LOCALAPPDATA="${LOCALAPPDATA:-$USERPROFILE/AppData/Local}"
if [ -z "$LOCALAPPDATA" ]; then
    LOCALAPPDATA="/c/Users/jt120832/AppData/Local"
fi

# Convert Windows paths to Git Bash format if needed
if [[ "$LOCALAPPDATA" =~ ^[A-Z]: ]]; then
    LOCALAPPDATA=$(cygpath -u "$LOCALAPPDATA" 2>/dev/null || echo "$LOCALAPPDATA" | sed 's|\\|/|g' | sed 's|^C:|/c|')
fi

# Clean Office 2016/365 (16.0) WEF directory
echo "Cleaning Office 2016/365 (16.0) WEF directories..."
WEF_DIR="$LOCALAPPDATA/Microsoft/Office/16.0/Wef"
remove_directory "$WEF_DIR"

# Clean Office 2013 (15.0) WEF directory
WEF_DIR_15="$LOCALAPPDATA/Microsoft/Office/15.0/Wef"
remove_directory "$WEF_DIR_15"

# Get temp directory path - Excel stores add-in files in AppData\Local\Temp\3
TEMP_3_DIR=""

# Method 1: Use TEMP environment variable
if [ -n "$TEMP" ]; then
    TEMP_3_DIR="$TEMP/3"
fi

# Method 2: Use TMP environment variable
if [ -z "$TEMP_3_DIR" ] && [ -n "$TMP" ]; then
    TEMP_3_DIR="$TMP/3"
fi

# Method 3: Construct from LOCALAPPDATA
if [ -z "$TEMP_3_DIR" ]; then
    TEMP_3_DIR="$LOCALAPPDATA/Temp/3"
fi

# Convert Windows paths to Git Bash format if needed
if [[ "$TEMP_3_DIR" =~ ^[A-Z]: ]]; then
    TEMP_3_DIR=$(cygpath -u "$TEMP_3_DIR" 2>/dev/null || echo "$TEMP_3_DIR" | sed 's|\\|/|g' | sed 's|^C:|/c|')
fi

# Clean the temp/3 directory (where Excel stores add-in files)
echo ""
echo "Cleaning temp directory..."
echo "Cleaning: $TEMP_3_DIR"
remove_directory "$TEMP_3_DIR"

echo ""
echo -e "${CYAN}=== Cleanup Complete ===${NC}"
echo -e "${YELLOW}Please close and restart Excel for changes to take effect.${NC}"
echo ""
