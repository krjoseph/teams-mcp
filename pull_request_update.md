This PR implements [issue #6](https://github.com/floriscornel/teams-mcp/issues/6) with a complete and secure solution:

## ✨ Features

- **Secure Markdown Support**: Adds `format` parameter (`text`, `markdown`) to `send_channel_message`, `send_chat_message`, and `reply_to_channel_message` tools
- **HTML Conversion**: Markdown is converted to sanitized HTML using `marked` and `DOMPurify` libraries
- **Security-First**: All content is sanitized to prevent XSS attacks and malicious content

## 🔧 Implementation Details

### Libraries Added
- `marked` - Fast, reliable Markdown parser with GitHub Flavored Markdown support
- `dompurify` - HTML sanitizer to prevent XSS attacks
- `jsdom` - DOM implementation for Node.js environment

### Format Options
- `text` (default): Plain text messages
- `markdown`: Markdown content converted to sanitized HTML

### Security Features
- **HTML Sanitization**: Removes dangerous elements (scripts, event handlers, etc.)
- **Allowed Tags**: Only safe HTML tags permitted (p, strong, em, a, ul, ol, li, h1-h6, code, pre, etc.)
- **Safe Attributes**: Only safe attributes allowed (href, target, src, alt, title, width, height)
- **XSS Prevention**: Comprehensive protection against cross-site scripting

### Supported Markdown Features
- Text formatting (bold, italic, strikethrough)
- Links and images
- Lists (bulleted and numbered)
- Code blocks and inline code
- Headings (H1-H6)
- Line breaks and blockquotes
- Tables (GitHub-flavored)

## 🧪 Testing

- **100% test coverage** for new markdown utility
- **Comprehensive integration tests** for all messaging tools
- **Security tests** to verify XSS protection
- **Error handling tests** for edge cases

## 📚 Documentation

- Updated README with detailed usage examples
- Security features documentation
- Supported markdown syntax reference
- Migration guide for existing users

## 🔄 Backward Compatibility

- Existing plain text messages work unchanged
- Default behavior remains `text` format
- No breaking changes to existing API

## 📊 Coverage Report

```
All files     |   96.42 |    71.71 |   93.75 |   96.42 |
utils         |     100 |      100 |     100 |     100 |
markdown.ts   |     100 |      100 |     100 |     100 |
```

Closes #6.