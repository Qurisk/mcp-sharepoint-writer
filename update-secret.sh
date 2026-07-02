#!/bin/bash
# Usage: ./update-secret.sh <new_azure_client_secret>

set -e

if [ -z "$1" ]; then
  echo "Usage: $0 <new_azure_client_secret>"
  exit 1
fi

NEW_SECRET="$1"
ENV_FILE="/root/mcp-sharepoint/.env"

ssh root@49.12.227.147 "
  sed -i 's|^AZURE_CLIENT_SECRET=.*|AZURE_CLIENT_SECRET=$NEW_SECRET|' $ENV_FILE
  echo 'Secret updated.'
  pm2 restart mcp-sharepoint --update-env
  pm2 show mcp-sharepoint | grep status
"

echo "Done."
