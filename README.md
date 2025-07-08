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
- OpenAI API key (for production use)

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

## 🔑 API Key Setup (Production Must-Have)

### 1. Get OpenAI API Key
1. Go to [OpenAI Platform](https://platform.openai.com/api-keys)
2. Sign in or create an account
3. Click "Create new secret key"
4. Copy the generated key (starts with `sk-`)

### 2. Configure API Key

#### Option A: Environment Variables (Recommended for Production)
Create a `.env` file in the root directory:
```bash
# .env
OPENAI_API_KEY=sk-your-api-key-here
```

#### Option B: Azure Static Web Apps (Production Deployment)
1. Go to your Azure Static Web App in Azure Portal
2. Navigate to "Configuration" → "Application settings"
3. Add a new setting:
   - **Name**: `OPENAI_API_KEY`
   - **Value**: `sk-your-api-key-here`
4. Save the configuration

#### Option C: Local Storage (Development Only)
The add-in will automatically detect and use API keys stored in localStorage.

### 3. Verify API Key Setup
```bash
# Test the API key configuration
npm run test
```

## 🚀 Production Deployment

### Azure Static Web Apps (Current Setup)
1. **Configure Environment Variables**:
   - Set `OPENAI_API_KEY` in Azure Portal
   - Set `NODE_ENV=production`

2. **Update Manifest URLs**:
   - Update `src/manifest/manifest.xml` with your production URLs
   - Replace `https://empty-walls-mix.loca.lt` with your actual domain

3. **Deploy**:
   ```bash
   npm run build
   # Push to main branch (GitHub Actions will auto-deploy)
   ```

### Office Add-in Store (Future)
1. **Update Manifest**:
   - Set production URLs in manifest
   - Add privacy policy URL
   - Configure proper permissions

2. **Submit for Review**:
   - Follow [Office Store guidelines](https://docs.microsoft.com/en-us/office/dev/add-ins/publish/publish)
   - Include privacy policy and terms of service

## 🔒 Security & Privacy (Production Must-Haves)

### Data Privacy
- **User prompts** are sent to OpenAI for processing
- **Excel context** (selected range, worksheet name) is included in requests
- **No sensitive workbook data** is permanently stored
- **No user data** is logged or retained

### API Key Security
- API keys are stored client-side (localStorage/environment variables)
- Keys are sent directly to OpenAI from the client
- Consider implementing a backend proxy for enterprise use

### Privacy Policy Requirements
Create a privacy policy covering:
- What data is sent to OpenAI
- How data is used and protected
- User rights and data retention
- Contact information for privacy concerns

## 🧪 Testing & Quality Assurance

### Automated Testing
```bash
# Run all tests
npm test

# Run tests with coverage
npm run test:coverage

# Run tests in watch mode
npm run test:watch
```

### Manual Testing Checklist
- [ ] Test all Excel operations in real workbooks
- [ ] Verify API key configuration works
- [ ] Test error handling and fallbacks
- [ ] Check accessibility features
- [ ] Test in different Excel environments (desktop, web, mobile)

### Production Testing
- [ ] Test with real user prompts and data
- [ ] Verify rate limiting and error handling
- [ ] Check performance with large workbooks
- [ ] Test edge cases and error scenarios

## 📋 Production Readiness Checklist

### ✅ Completed
- [x] Core Excel operations (modify, insert, delete, copy, move, format, formula, sort, filter, chart, table, clean)
- [x] AI service integration with OpenAI API
- [x] Error handling and logging
- [x] Type safety and schema validation
- [x] Comprehensive test coverage
- [x] Modern, responsive UI
- [x] Azure Static Web Apps deployment
- [x] Office Add-in manifest configuration

### 🔄 In Progress
- [ ] API key management UI
- [ ] Operation history/undo functionality
- [ ] Advanced Excel features (pivot tables, conditional formatting)
- [ ] Internationalization support
- [ ] Onboarding and help system

### 📝 Future Enhancements
- [ ] Backend proxy for API key security
- [ ] User authentication and usage tracking
- [ ] Advanced analytics and insights
- [ ] Custom AI model fine-tuning
- [ ] Enterprise features (SSO, admin controls)

## 🐛 Troubleshooting

### Common Issues

#### API Key Not Working
```bash
# Check if API key is set
echo $OPENAI_API_KEY

# Test API key manually
curl -H "Authorization: Bearer sk-your-key" \
     -H "Content-Type: application/json" \
     -d '{"model":"gpt-3.5-turbo","messages":[{"role":"user","content":"test"}]}' \
     https://api.openai.com/v1/chat/completions
```

#### Excel Operations Failing
- Check Office.js is properly initialized
- Verify Excel permissions in manifest
- Check browser console for errors

#### Build Issues
```bash
# Clear cache and reinstall
rm -rf node_modules package-lock.json
npm install

# Check TypeScript errors
npm run type-check
```

## 📚 Documentation

- [API Documentation](./docs/api.md)
- [Architecture Overview](./docs/architecture.md)
- [Design Document](./docs/design.md)
- [Development Guidelines](./.cursorrules)

## 🔗 Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Office.js API Reference](https://docs.microsoft.com/en-us/javascript/api/office)
- [OpenAI API Documentation](https://platform.openai.com/docs/api-reference)
- [Azure Static Web Apps](https://docs.microsoft.com/en-us/azure/static-web-apps/)

## 📄 License

MIT License - see [LICENSE](LICENSE) file for details.

## 🤝 Contributing

See [CONTRIBUTING.md](CONTRIBUTING.md) for development guidelines and collaboration practices.