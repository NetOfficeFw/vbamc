#!/usr/bin/env bash
###
### Codex environment setup script. 
###
set -euo pipefail

# Download and install .NET 10 SDK
curl -sSL https://dot.net/v1/dotnet-install.sh -o dotnet-install.sh
bash ./dotnet-install.sh --channel 10.0 --install-dir /usr/local/dotnet
ln -sf /usr/local/dotnet/dotnet /usr/local/bin/dotnet

# configure system with dotnet
cat >/etc/profile.d/dotnet.sh <<'EOF'
export DOTNET_ROOT=/usr/local/dotnet
export PATH=/usr/local/dotnet:$PATH
export PATH=$HOME/.dotnet/tools:$PATH
EOF

# clean up
rm -rf ./dotnet-install.sh
