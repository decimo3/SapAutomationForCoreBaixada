#!/bin/env bash

set -e

if [[ -z "$VIRTUAL_ENV" ]]; then
    source venv/Scripts/activate
fi

# DONE - Fetch tags and get the latest one
git fetch --tags
version=$(git describe --tags $(git rev-list --tags --max-count=1))
version_number="${version#v}"  # Remove 'v' prefix if the tag has it
echo "Version: $version_number"

# Extract major, minor, and patch versions
IFS='.' read -r MAJOR_VERSION MINOR_VERSION PATCH_VERSION <<< "$version_number"
export MAJOR_VERSION MINOR_VERSION PATCH_VERSION
echo "Major version: $MAJOR_VERSION"
echo "Minor version: $MINOR_VERSION"
echo "Patch version: $PATCH_VERSION"

# Write version file
envsubst < version_file.txt > version_file.tmp
mv version_file.tmp version_file.txt
cat version_file.txt

# Install dependencies
pip install -r requirements.txt

# Lint with pylint
#pylint src/main.py

# Test with pytest
#pytest src/test_main.py

# Build executable with pyinstaller
pyinstaller -n sap --icon appicon.ico --version-file version_file.txt --onefile src/main.py

# Create or update sqlite database
sqlite3 dist/sap.db < src/sap.sql

# Compress executable and related files
zip -j dist/sap.zip dist/sap.exe src/sap.conf src/sap.path dist/sap.db src/fileDialog.vbs src/erroDialog.vbs

# Generate SHA256 checksum
sha256sum dist/sap.zip | awk '{print $1}' > dist/sap.zip.sha256sum
SHA_WIN64_ZIP=$(cat dist/sap.zip.sha256sum)
echo $SHA_WIN64_ZIP

# Write release notes
envsubst < release_notes.md > release_notes.tmp
mv release_notes.tmp release_notes.md
cat release_notes.md

# Create a release on GitHub
gh release create $version --verify-tag --notes-file release_notes.md --title "SAP_BOT release" dist/sap.zip#sap.zip dist/sap.zip.sha256sum#sap.zip.sha256sum

# Reverting placeholder files
git restore release_notes.md
git restore version_file.txt