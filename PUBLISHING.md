# Publishing Guide

## Prerequisites

1. **NPM Account**: Make sure you have an NPM account and are logged in:
   ```bash
   npm login
   ```

2. **Scoped Package Access**: Since this is a scoped package (`@floriscornel/teams-mcp`), ensure you have:
   - Permission to publish under the `@floriscornel` scope
   - Set `publishConfig.access` to `"public"` in `package.json` (already done)

## Publishing Steps

### 1. Version Update
Update the version in `package.json`:
```bash
# For patch releases (bug fixes)
npm version patch

# For minor releases (new features)
npm version minor

# For major releases (breaking changes)
npm version major
```

### 2. Build and Test
```bash
# Build the project
bun run build

# Test the CLI locally
node dist/index.js --help
node dist/index.js check
```

### 3. Publish to NPM
```bash
# Dry run to see what will be published
npm publish --dry-run

# Actually publish
npm publish
```

## Verification

After publishing, you can test the package:

```bash
# Test authentication command
npx @floriscornel/teams-mcp@latest authenticate

# Test help
npx @floriscornel/teams-mcp@latest --help

# Test in Cursor with this configuration:
```

```json
{
  "mcpServers": {
    "teams-mcp": {
      "command": "npx",
      "args": ["-y", "@floriscornel/teams-mcp@latest"]
    }
  }
}
```

## Troubleshooting

- **403 Forbidden**: Check if you have permission to publish under the `@floriscornel` scope
- **Package already exists**: Make sure you've incremented the version number
- **Missing files**: Check that all necessary files are included (see `package.json` files array)

## Automated Publishing (Optional)

Consider setting up GitHub Actions for automated publishing:

1. Create `.github/workflows/publish.yml`
2. Add NPM_TOKEN as a repository secret
3. Configure automatic publishing on tag creation 