#!/bin/bash

# Script to trigger GitHub Actions workflows via GitHub CLI

REPO="hoanganhduc/library"  # Replace with your GitHub repo
BRANCH="main"                   # Replace with your branch if needed

echo "Select an option:"
echo "1) Generate HTML (build.yml)"
echo "2) Send Email (sendbook.yml)"
read -p "Enter choice [1 or 2]: " choice

case "$choice" in
    1)
        gh workflow run build.yml --repo "$REPO" --ref "$BRANCH"
        ;;
    2)
        gh workflow run sendbook.yml --repo "$REPO" --ref "$BRANCH"
        ;;
    *)
        echo "Invalid choice."
        exit 1
        ;;
esac