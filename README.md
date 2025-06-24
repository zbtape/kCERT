# Model Review Tool (MRT) - Excel Add-in

A professional Excel add-in designed for analyzing and reviewing Excel-based models with comprehensive formula analysis and audit trail capabilities.

## Overview

The Model Review Tool addresses the critical need for standardized, thorough review of Excel-based models used in business and risk management scenarios. It provides transparency into formula structures, creates audit trails, and helps ensure accuracy in model reviews.

## Features

### Phase 1 - Formula Analysis
- **Workbook Analysis**: Analyze all worksheets in a workbook for formula content
- **Unique Formula Detection**: Identify and count unique formulas across worksheets
- **Formula Complexity Assessment**: Automatically assess formula complexity levels
- **Professional Reporting**: Generate formatted analysis reports within Excel
- **Audit Trail Generation**: Create downloadable audit trails for compliance

### Key Benefits
- **Standardized Reviews**: Consistent, systematic approach to model review
- **Transparency**: Clear visibility into formula structures and relationships
- **Audit Trail**: Complete documentation of review activities
- **Time Efficiency**: Faster identification of formulas requiring review
- **Professional Output**: Enterprise-ready reports and documentation

## Installation

### Prerequisites
- Microsoft Excel (Office 365 or Excel 2016+)
- Node.js (v14 or higher)
- npm or yarn package manager

### Development Setup

1. **Clone the repository**
   ```bash
   git clone <repository-url>
   cd MRT-Tool
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Start development server**
   ```bash
   npm run dev-server
   ```

4. **Sideload the add-in in Excel**
   ```bash
   npm run sideload
   ```

### Building for Production

1. **Build the add-in**
   ```bash
   npm run build
   ```

2. **Validate the manifest**
   ```bash
   npm run validate
   ```

## Usage

### Running Formula Analysis

1. **Open Excel** with a workbook containing formulas
2. **Launch MRT** from the Home tab ribbon
3. **Configure Options**:
   - Include empty cells in analysis
   - Group similar formulas for better organization
4. **Click "Analyze Workbook Formulas"**
5. **Review Results** in the task pane showing:
   - Summary statistics (total worksheets, formulas, unique formulas)
   - Per-worksheet breakdown
   - List of unique formulas with usage counts

### Exporting Results

- **Analysis Report**: Creates a new worksheet with formatted analysis results
- **Audit Trail**: Downloads a JSON file with complete analysis metadata

## Project Structure

```
MRT-Tool/
├── src/
│   ├── taskpane/           # Main task pane interface
│   │   ├── taskpane.html   # UI layout
│   │   ├── taskpane.css    # Styling
│   │   └── taskpane.ts     # Main logic
│   ├── shared/             # Shared utilities
│   │   └── FormulaAnalyzer.ts  # Core analysis engine
│   └── commands/           # Office commands
├── assets/                 # Icons and resources
├── manifest.xml           # Add-in configuration
├── package.json           # Dependencies and scripts
└── webpack.config.js      # Build configuration
```

## Technical Architecture

### Core Components

- **FormulaAnalyzer Class**: Main analysis engine for processing Excel formulas
- **Professional UI**: Microsoft Fabric UI components for consistent Office experience
- **TypeScript**: Type-safe development with full Office.js API support
- **Webpack**: Modern build system with hot reload for development

### Formula Analysis Features

- **Formula Detection**: Identifies cells containing formulas vs. values
- **Normalization**: Groups structurally similar formulas for better analysis
- **Complexity Assessment**: Automatic scoring based on formula patterns
- **Cell Reference Mapping**: Tracks formula usage across worksheets

## Development

### Available Scripts

- `npm run dev-server`: Start development server with hot reload
- `npm run build`: Build for production
- `npm run build:dev`: Build for development
- `npm start`: Start Office debugging
- `npm run sideload`: Sideload add-in to Excel
- `npm run unload`: Remove add-in from Excel
- `npm run validate`: Validate manifest file

### Contributing

1. Create feature branch from main
2. Implement changes with appropriate tests
3. Update documentation as needed
4. Submit pull request for review

## Security & Compliance

- **Data Privacy**: All analysis is performed locally within Excel
- **Audit Trails**: Complete documentation of all analysis activities
- **No External Dependencies**: Core functionality works offline
- **Enterprise Ready**: Suitable for regulated environments

## Support

For technical support or feature requests, please:
1. Check the documentation
2. Review existing issues
3. Create new issue with detailed description

## License

MIT License - see LICENSE file for details

## Roadmap

### Future Enhancements
- Advanced formula validation rules
- Custom formula libraries
- Integration with version control systems
- Enhanced reporting templates
- Multi-language support

---

**Model Review Tool** - Professional Excel model analysis for enterprise environments. 