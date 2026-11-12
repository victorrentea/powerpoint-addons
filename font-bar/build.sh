#!/bin/sh
# Builds the 🔤 font bar that server.js starts (font-bar.js); restart the server after.
# Info.plist goes inside the binary and the local signing identity seals it, so
# the Automation permission (PowerPoint) stays granted across rebuilds.
set -e
cd "$(dirname "$0")"
mkdir -p .build
swiftc -O -o .build/font-bar FontBar.swift \
  -Xlinker -sectcreate -Xlinker __TEXT -Xlinker __info_plist -Xlinker Info.plist
IDENTITY="Victor Addons Local Code Signing"
if security find-identity -v -p codesigning | grep -Fq "$IDENTITY"; then
  codesign --force --sign "$IDENTITY" --identifier ro.victorrentea.powerpoint-font-bar .build/font-bar
else
  echo "⚠️  '$IDENTITY' not in the keychain: ad-hoc signature, PowerPoint permission asked again after each build"
  codesign --force --sign - --identifier ro.victorrentea.powerpoint-font-bar .build/font-bar
fi
echo "built $(pwd)/.build/font-bar"
