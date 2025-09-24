# kCERT - KPMG Comprehensive Excel Review Tool

A professional Excel add-in designed for analyzing and reviewing Excel-based models with comprehensive formula analysis and audit trail capabilities.

## Overview

kCERT (KPMG Comprehensive Excel Review Tool) addresses the critical need for standardized, thorough review of Excel-based models used in business and risk management scenarios. It provides transparency into formula structures, creates audit trails, and helps ensure accuracy in model reviews.

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

1. **Open Excel** with the workbook you want to review.
2. **Launch kCERT** from the Home tab ribbon to open the task pane.
3. **Configure options** (include empty cells, group similar formulas) as needed.
4. **Click "Analyze Workbook Formulas"**. The task pane displays live status updates:
   - Every streamed block shows `Streaming SheetName: rows a-b, cols c-d` so you know progress on large sheets.
   - When a worksheet exceeds the massive threshold (~150k cells), the pane announces the switch to "skim" mode.
5. **Review results** in the task pane:
   - Summary statistics (worksheets, total formulas, unique formulas, hard-coded values).
   - Per-worksheet cards now include `Mode: streaming | massive-skim | skipped` plus a fallback reason if applicable.
   - Each card lists the most common formulas, cell counts, and captured hard-coded literals (capped for performance).

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

- **FormulaAnalyzer (streaming engine)**: Processes worksheets in fixed-size blocks (default 200×120) to avoid loading entire sheets into memory.
- **Progress Reporter**: Emits status strings for each block so the task pane reflects real-time progress and fallbacks.
- **Task Pane UI**: Renders summary stats, per-worksheet cards, and exposes analysis modes / fallback reasons.
- **TypeScript + Webpack**: Modern toolchain with live reload for development.

### Streaming Analysis Highlights

- **Block-based Loading**: Each worksheet is scanned block by block. Data is discarded immediately after processing, keeping memory usage flat.
- **Massive-Skim Mode**: Workbooks above the cell threshold fall back to a lightweight count (formulas/constants) to avoid long waits. The UI marks these sheets as `Mode: massive-skim`.
- **Capped Sampling**: Unique formulas retain at most 200 sample addresses per formula, and hard-coded detection stores the first 400 findings per sheet.
- **Hard-Coded Detection**: Inline literal scanning (numbers, strings, arrays) runs during streaming without secondary passes.
- **Analysis Metadata**: Worksheet results include `analysisMode` and optional `fallbackReason` so downstream consumers know which path executed.

## Development

### Available Scripts

- `npm run dev-server`: Start development server with hot reload
- `npm run build`: Build for production
- `npm run build:dev`: Build for development
- `npm start`: Start Office debugging (launches dev-server, sideloads Excel, and opens DevTools)
- `npm run sideload`: Sideload add-in to Excel using the latest production build
- `npm run unload`: Remove add-in from Excel
- `npm run validate`: Validate manifest file

### Recommended Desktop Workflow

1. Run `npm start` for debugging: this launches the HTTPS dev server, sideloads the add-in, and opens Edge DevTools for the task pane WebView.
2. If you prefer the production bundle, run `npm run build` followed by `npm run sideload` (ensure no dev-server is running). Remove stale XLAM add-ins from Excel Options if Excel complains about missing files.
3. To stop debugging, run `npm run stop`.

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

**kCERT** - Professional Excel model analysis for enterprise environments. 