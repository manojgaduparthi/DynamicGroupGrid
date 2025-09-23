# Contributing to Dynamic Group Grid PCF Control

Thank you for considering contributing to the Dynamic Group Grid PCF Control! This document provides guidelines and information for contributors.

## 🤝 How to Contribute

We welcome contributions in many forms:
- 🐛 Bug reports and fixes
- ✨ New features and enhancements
- 📖 Documentation improvements
- 🧪 Test cases and quality improvements
- 💡 Ideas and suggestions

## 📋 Before You Start

1. **Check existing issues** - Look for existing issues or discussions about your idea
2. **Create an issue** - For major changes, create an issue first to discuss
3. **Fork the repository** - Create your own fork to work on
4. **Set up development environment** - Follow the setup instructions below

## 🛠️ Development Setup

### Prerequisites
- Node.js (version 14 or higher)
- Power Platform CLI (`pac`)
- A Dynamics 365 or Power Apps environment for testing

### Getting Started
```bash
# Clone your fork
git clone https://github.com/manojgaduparthi/DynamicGroupGrid.git
cd DynamicGroupGrid

# Install dependencies
npm install

# Build the control
npm run build

# Start development (optional)
npm start watch
```

### Testing Your Changes
```bash
# Build and test locally
npm run build

# Deploy to test environment
pac pcf push --environment your-test-environment-id
```

## 📝 Contribution Guidelines

### Code Style
- Follow existing TypeScript/JavaScript patterns
- Use meaningful variable and function names
- Add comments for complex logic
- Maintain consistent indentation (2 spaces)

### Commit Messages
Use clear, descriptive commit messages:
```
feat: add support for custom column sorting
fix: resolve selection sync issue with ribbon commands
docs: update installation instructions
style: improve responsive design for mobile devices
```

### Pull Request Process

1. **Create a feature branch**
   ```bash
   git checkout -b feature/your-feature-name
   ```

2. **Make your changes**
   - Write clear, tested code
   - Update documentation if needed
   - Test thoroughly in multiple environments

3. **Commit your changes**
   ```bash
   git add .
   git commit -m "feat: describe your changes"
   ```

4. **Push to your fork**
   ```bash
   git push origin feature/your-feature-name
   ```

5. **Create a Pull Request**
   - Use a clear title and description
   - Reference any related issues
   - Include screenshots for UI changes
   - Wait for review and respond to feedback

## 🧪 Testing Guidelines

### Manual Testing Checklist
Before submitting, ensure your changes work correctly:

- [ ] Control loads without errors
- [ ] Grouping functionality works
- [ ] Pagination operates correctly
- [ ] Selection syncs with host ribbon
- [ ] Column resizing persists
- [ ] Responsive design on different screen sizes
- [ ] Performance is acceptable with large datasets

### Test Environments
Test your changes in:
- Model-driven apps
- Different entity types
- Various view configurations
- Mobile and desktop browsers

## 📖 Documentation

When contributing, consider updating:
- README.md for new features
- Code comments for complex logic
- This CONTRIBUTING.md for process changes

## 🐛 Reporting Bugs

When reporting bugs, please include:
- **Clear description** of the issue
- **Steps to reproduce** the problem
- **Expected vs actual behavior**
- **Environment details** (browser, D365 version, etc.)
- **Screenshots or videos** if applicable

### Bug Report Template
```markdown
**Describe the bug**
A clear description of what the bug is.

**To Reproduce**
Steps to reproduce the behavior:
1. Go to '...'
2. Click on '....'
3. See error

**Expected behavior**
What you expected to happen.

**Environment:**
- Browser: [e.g. Chrome 96]
- D365 Version: [e.g. 9.2]
- Control Version: [e.g. 1.0.0]

**Additional context**
Any other context about the problem.
```

## 🎯 Feature Requests

For feature requests, please:
1. Check if the feature already exists or is planned
2. Describe the problem you're trying to solve
3. Explain how this feature would help users
4. Consider implementation complexity

## 📜 Code of Conduct

We expect all contributors to:
- Be respectful and inclusive
- Provide constructive feedback
- Focus on what's best for the community
- Show empathy towards other contributors

## ❓ Questions?

If you have questions about contributing:
- Create a [Discussion](https://github.com/manojgaduparthi/DynamicGroupGrid/discussions)
- Check existing [Issues](https://github.com/manojgaduparthi/DynamicGroupGrid/issues)
- Review this document and the README

Thank you for contributing to the Dynamic Group Grid PCF Control! 🙏