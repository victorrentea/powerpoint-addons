#!/bin/sh
# Builds the 🔤 font bar that server.js starts (font-bar.js); restart the server after.
set -e
cd "$(dirname "$0")"
mkdir -p .build
swiftc -O -o .build/font-bar FontBar.swift
echo "built $(pwd)/.build/font-bar"
