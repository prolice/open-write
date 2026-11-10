#!/usr/bin/env bash
# Appelle l'Azure CLI de Windows depuis WSL : tools/az.sh <arguments az>
AZ='C:\Program Files\Microsoft SDKs\Azure\CLI2\wbin\az.cmd'
args=""
for a in "$@"; do args="$args '${a//\'/\'\'}'"; done
powershell.exe -NoProfile -Command "\$OutputEncoding=[Console]::OutputEncoding=[Text.Encoding]::UTF8; & '$AZ' $args" | tr -d '\r'
