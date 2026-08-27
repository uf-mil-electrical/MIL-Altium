#!/usr/bin/env bash
set -e

REPO_URL="https://github.com/uf-mil-electrical/MIL-Altium.git"
DEST="C:/MIL-Altium"

# Check git is installed
if ! command -v git &> /dev/null; then
    echo "Git is not installed. Please install it from https://gitforwindows.org/ and re-run this script."
    exit 1
fi

# Abort if destination already exists
if [ -e "$DEST" ]; then
    echo "$DEST already exists. Aborting to avoid overwriting or duplicating files."
    exit 1
fi

echo "Cloning into $DEST..."
git clone "$REPO_URL" "$DEST"

echo "Done."
