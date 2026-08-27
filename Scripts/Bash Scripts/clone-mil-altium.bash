#!/usr/bin/env bash
set -e

REPO_URL="https://github.com/uf-mil-electrical/MIL-Altium.git"
DEST="C:/MIL-Altium"

# Check git is installed
if ! command -v git &> /dev/null; then
    echo "Git is not installed. Please install it from https://git-scm.com/download/win and re-run this script."
    exit 1
fi

if [ -d "$DEST" ]; then
    if [ -d "$DEST/.git" ]; then
        echo "Repo already exists at $DEST. Pulling latest changes..."
        git -C "$DEST" pull
    else
        echo "$DEST exists but is not a git repo. Aborting to avoid overwriting files."
        exit 1
    fi
else
    echo "Cloning into $DEST..."
    git clone "$REPO_URL" "$DEST"
fi

echo "Done."