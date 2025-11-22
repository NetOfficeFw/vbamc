#!/usr/bin/env bash

dotnet publish -c Release -r osx-arm64 -f net10.0 --self-contained true
