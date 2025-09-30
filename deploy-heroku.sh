#!/bin/bash

# Exit on error
set -e

# Check for Heroku CLI
if ! command -v heroku &> /dev/null
then
    echo "Heroku CLI not found. Please install it first."
    exit 1
fi

# Login to Heroku (if not already logged in)
heroku whoami &> /dev/null || heroku login

# Create Procfile if it doesn't exist
if [ ! -f Procfile ]; then
    echo "Procfile not found."
    exit 1
fi

# Set Heroku remote to existing app 'microsoft-teams-mcp'
heroku git:remote -a microsoft-teams-mcp

# Ensure Node.js buildpack is set
heroku buildpacks:set heroku/nodejs -a microsoft-teams-mcp || true

# Commit Procfile if needed
if [ -n "$(git status --porcelain Procfile)" ]; then
    git add Procfile
    git commit -m "Update Procfile for Heroku deployment"
fi

# Commit package.json if needed (for tsx dependency)
if [ -n "$(git status --porcelain package.json)" ]; then
    git add package.json
    git commit -m "Add tsx dependency for TypeScript execution"
fi

# Set the branch to deploy
BRANCH_TO_DEPLOY="new-server-instance-per-request"

# Push to Heroku (force push to replace existing code from another repo)
git push heroku $BRANCH_TO_DEPLOY:main --force

echo "Deployment to Heroku app 'microsoft-teams-mcp' initiated."