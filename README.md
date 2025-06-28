# SheetSense - AI-Powered Excel Add-in

An intelligent Excel Add-in that enables natural language interaction with spreadsheets through AI.

## 🎯 Overview

SheetSense transforms how users interact with Excel by enabling:
- **Natural Language Commands**: "Add a profit margin column" or "Explain this formula"
- **AI-Powered Analysis**: Intelligent explanations of formulas and data structures
- **Smart Editing**: Preview and apply changes with undo/redo support
- **Accessibility**: Full keyboard navigation and screen reader support

## 🏗️ Project Structure

```
excelauto/
├── docs/                    # Documentation
│   ├── design.md           # Product & Technical Design
│   ├── architecture.md     # System Architecture Details
│   └── api.md              # API Documentation
├── src/                    # Source Code
│   ├── manifest/           # Office Add-in Manifest
│   ├── ui/                 # React UI Components
│   ├── services/           # Business Logic Services
│   └── utils/              # Utility Functions
├── tests/                  # Test Files
├── tasks/                  # Development Tasks
└── fixes/                  # Bug Fix Documentation
```

## 🛠️ Development Setup

### Prerequisites
- Node.js 18+ 
- Excel (Desktop or Online)
- Office Add-in development tools

### Quick Start
```bash
# Install dependencies
npm install

# Start development server
npm run dev

# Build for production
npm run build

# Run tests
npm test
```

## 📋 Current Status

- [x] Project structure setup
- [ ] Office Add-in manifest configuration
- [ ] React UI framework setup
- [ ] Office.js integration
- [ ] AI service integration
- [ ] Core features implementation

## 🔗 Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Office.js API Reference](https://docs.microsoft.com/en-us/javascript/api/office)
- [Design Document](./docs/design.md)

## 📝 Development Guidelines

See `.cursorrules` for detailed development guidelines and collaboration practices.