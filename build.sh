#!/bin/bash

dotnet publish -c Release -r osx.10.15-x64 -f net6.0 --self-contained true -o dist
