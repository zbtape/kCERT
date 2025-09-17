import './taskpane.css';
import { FormulaAnalyzer } from '../shared/FormulaAnalyzer';

// Wait for Office.js to load
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('analyzeFormulas')?.addEventListener('click', analyzeFormulas);
        document.getElementById('exportResults')?.addEventListener('click', exportResults);
        document.getElementById('generateAuditTrail')?.addEventListener('click', generateAuditTrail);
        
        // Components are now initialized
    }
});

let analysisResults: any = null;

// Fabric UI components are no longer needed - using standard HTML checkboxes

/**
 * Main function to analyze formulas in the workbook
 */
async function analyzeFormulas(): Promise<void> {
    try {
        showLoadingIndicator(true);
        showStatusMessage('Starting formula analysis...', 'info');
        
        const options = {
            includeEmptyCells: (document.getElementById('includeEmptyCells') as HTMLInputElement).checked,
            groupSimilarFormulas: (document.getElementById('groupSimilarFormulas') as HTMLInputElement).checked
        };
        
        await Excel.run(async (context) => {
            const analyzer = new FormulaAnalyzer();
            
            // Show progress for large models
            showStatusMessage('Analyzing workbook structure...', 'info');
            const results = await analyzer.analyzeWorkbook(context, options);
            
            // Show specific message for massive models
            if (results.totalCells > 50000) {
                showStatusMessage('MASSIVE MODEL detected - using minimal analysis mode for performance', 'info');
            }
            
            analysisResults = results;
            displayResults(results);
            
            const totalCells = results.totalCells;
            if (totalCells > 100000) {
                showStatusMessage(`Analysis completed! Processed ${totalCells.toLocaleString()} cells across ${results.totalWorksheets} worksheets.`, 'success');
            } else {
                showStatusMessage('Formula analysis completed successfully!', 'success');
            }
        });
        
    } catch (err) {
        console.error('Error analyzing formulas:', err);
        const errorMsg = getErrorMessage(err);
        if (errorMsg.includes('Maximum call stack size exceeded')) {
            showStatusMessage(`STACK OVERFLOW ERROR - Using optimized version v6.0 with NULL-SAFE HARD-CODED DETECTION. Please check if you're using the latest build. Error: ${errorMsg}`, 'error');
        } else {
            showStatusMessage(`Error during analysis: ${errorMsg}`, 'error');
        }
    } finally {
        showLoadingIndicator(false);
    }
}

/**
 * Display the analysis results in the UI
 */
function displayResults(results: any): void {
    // Update summary statistics
    document.getElementById('totalWorksheets')!.textContent = results.totalWorksheets.toString();
    document.getElementById('totalFormulas')!.textContent = results.totalFormulas.toString();
    document.getElementById('uniqueFormulas')!.textContent = results.uniqueFormulas.toString();
    document.getElementById('totalHardCodedValues')!.textContent = results.totalHardCodedValues.toString();
    
    // Update cell count analysis
    document.getElementById('totalCells')!.textContent = results.totalCells.toString();
    document.getElementById('cellsWithFormulas')!.textContent = results.totalCellsWithFormulas.toString();
    document.getElementById('cellsWithValues')!.textContent = results.totalCellsWithValues.toString();
    document.getElementById('emptyCells')!.textContent = (results.totalCells - results.totalCellsWithFormulas - results.totalCellsWithValues).toString();
    
    // Update hard-coded values analysis with null checks
    const highSeverity = results.worksheets.reduce((sum: number, ws: any) => {
        const analysis = ws.hardCodedValueAnalysis;
        return sum + (analysis?.highSeverityValues?.length || 0);
    }, 0);
    const mediumSeverity = results.worksheets.reduce((sum: number, ws: any) => {
        const analysis = ws.hardCodedValueAnalysis;
        return sum + (analysis?.mediumSeverityValues?.length || 0);
    }, 0);
    const lowSeverity = results.worksheets.reduce((sum: number, ws: any) => {
        const analysis = ws.hardCodedValueAnalysis;
        return sum + (analysis?.lowSeverityValues?.length || 0);
    }, 0);
    const infoSeverity = results.worksheets.reduce((sum: number, ws: any) => {
        const analysis = ws.hardCodedValueAnalysis;
        return sum + (analysis?.infoSeverityValues?.length || 0);
    }, 0);
    
    document.getElementById('highConfidenceValues')!.textContent = highSeverity.toString();
    document.getElementById('mediumConfidenceValues')!.textContent = mediumSeverity.toString();
    document.getElementById('lowConfidenceValues')!.textContent = lowSeverity.toString();
    
    // Display hard-coded values list
    displayHardCodedValues(results.worksheets);
    
    // Display worksheet details
    const worksheetDetails = document.getElementById('worksheetDetails')!;
    worksheetDetails.innerHTML = '';
    
    results.worksheets.forEach((worksheet: any) => {
        const worksheetCard = createWorksheetCard(worksheet);
        worksheetDetails.appendChild(worksheetCard);
    });
    
    // Show results section
    document.getElementById('resultsSection')!.style.display = 'block';
}

/**
 * Display hard-coded values in the UI
 */
function displayHardCodedValues(worksheets: any[]): void {
    const hardCodedValuesList = document.getElementById('hardCodedValuesList')!;
    hardCodedValuesList.innerHTML = '';
    
    // Collect all hard-coded values from all worksheets with null checks
    const allHardCodedValues: any[] = [];
    worksheets.forEach(worksheet => {
        const analysis = worksheet.hardCodedValueAnalysis;
        if (analysis) {
            allHardCodedValues.push(
                ...(analysis.highSeverityValues || []),
                ...(analysis.mediumSeverityValues || []),
                ...(analysis.lowSeverityValues || [])
            );
        }
    });
    
    // Sort by severity level (high first) and then by value
    allHardCodedValues.sort((a, b) => {
        const severityOrder: { [key: string]: number } = { 'High': 0, 'Medium': 1, 'Low': 2, 'Info': 3 };
        if (severityOrder[a.severity] !== severityOrder[b.severity]) {
            return severityOrder[a.severity] - severityOrder[b.severity];
        }
        return a.value.localeCompare(b.value);
    });
    
    // Display hard-coded values (limit to first 50 for performance)
    const displayValues = allHardCodedValues.slice(0, 50);
    displayValues.forEach(value => {
        const valueItem = createHardCodedValueItem(value);
        hardCodedValuesList.appendChild(valueItem);
    });
    
    if (allHardCodedValues.length > 50) {
        const moreItem = document.createElement('div');
        moreItem.className = 'hard-coded-value-item';
        moreItem.innerHTML = `<div class="hard-coded-value-content">... and ${allHardCodedValues.length - 50} more hard-coded values</div>`;
        hardCodedValuesList.appendChild(moreItem);
    }
}

/**
 * Create a hard-coded value item element with enhanced inconsistency information
 */
function createHardCodedValueItem(value: any): HTMLElement {
    const item = document.createElement('div');
    item.className = `hard-coded-value-item ${value.severity.toLowerCase()}-confidence`;
    
    const repetitionInfo = value.isRepeated ? ` (Repeated ${value.repetitionCount} times)` : '';
    const inconsistencyBadge = value.isInconsistent ? 
        `<span class="inconsistency-badge ${value.inconsistencyType}">${getInconsistencyLabel(value.inconsistencyType)}</span>` : '';
    
    let nearbyPatternsHtml = '';
    if (value.nearbyPatterns && value.nearbyPatterns.length > 0) {
        nearbyPatternsHtml = `<div class="nearby-patterns">
            <strong>Nearby formulas:</strong>
            <ul>${value.nearbyPatterns.map((p: string) => `<li>${escapeHtml(p)}</li>`).join('')}</ul>
        </div>`;
    }
    
    item.innerHTML = `
        <div class="hard-coded-value-header">
            <span class="hard-coded-value-address">${value.cellAddress}</span>
            <span class="hard-coded-value-confidence ${value.severity.toLowerCase()}">${value.severity}</span>
            ${inconsistencyBadge}
        </div>
        <div class="hard-coded-value-content">${escapeHtml(value.context)}</div>
        <div class="hard-coded-value-reason">
            <strong>Value:</strong> "${value.value}"${repetitionInfo}<br>
            <strong>Issue:</strong> ${value.rationale}<br>
            <strong>Suggested Fix:</strong> ${value.suggestedFix}
            ${value.expectedPattern ? `<br><strong>Expected:</strong> ${value.expectedPattern}` : ''}
        </div>
        ${nearbyPatternsHtml}
    `;
    
    return item;
}

/**
 * Get a user-friendly label for inconsistency types
 */
function getInconsistencyLabel(type: string): string {
    switch(type) {
        case 'value_mismatch':
            return 'Inconsistent Value';
        case 'range_endpoint':
            return 'Fixed Range';
        case 'pattern_deviation':
            return 'Pattern Deviation';
        case 'isolated_hardcode':
            return 'Isolated Hard-code';
        default:
            return 'Inconsistent';
    }
}

/**
 * Create a worksheet card element
 */
function createWorksheetCard(worksheet: any): HTMLElement {
    const card = document.createElement('div');
    card.className = 'worksheet-card';
    
    card.innerHTML = `
        <div class="worksheet-header">
            <div class="worksheet-name">${worksheet.name}</div>
        </div>
        <div class="worksheet-stats">
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${worksheet.totalFormulas}</span>
                <span class="worksheet-stat-label">Total Formulas</span>
            </div>
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${worksheet.uniqueFormulas}</span>
                <span class="worksheet-stat-label">Unique Formulas</span>
            </div>
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${worksheet.cellCountAnalysis.totalCells}</span>
                <span class="worksheet-stat-label">Total Cells</span>
            </div>
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${worksheet.formulaComplexity}</span>
                <span class="worksheet-stat-label">Complexity</span>
            </div>
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${worksheet.hardCodedValueAnalysis?.totalHardCodedValues || 0}</span>
                <span class="worksheet-stat-label">Hard-coded</span>
            </div>
        </div>
        <div class="formula-list" id="formulaList_${worksheet.name.replace(/\s+/g, '_')}">
            <h4>Unique Formulas:</h4>
            ${worksheet.uniqueFormulasList.map((formula: any) => `
                <div class="formula-item">
                    <div class="formula-text">${escapeHtml(formula.formula)}</div>
                    <div class="formula-count">Used ${formula.count} time(s)</div>
                </div>
            `).join('')}
        </div>
    `;
    
    return card;
}

/**
 * Export analysis results
 */
async function exportResults(): Promise<void> {
    if (!analysisResults) {
        showStatusMessage('No analysis results to export. Please run analysis first.', 'warning');
        return;
    }
    
    try {
        showStatusMessage('Exporting analysis results...', 'info');
        
        await Excel.run(async (context) => {
            // Remove existing report sheet if present to avoid double-counting
            const existing = context.workbook.worksheets.getItemOrNullObject('CERT_Analysis_Report');
            existing.load('name');
            await context.sync();
            if (!existing.isNullObject) {
                existing.delete();
                await context.sync();
            }

            // Create a new worksheet for the report
            const reportSheet = context.workbook.worksheets.add('CERT_Analysis_Report');

            // Helper to set a single cell value using A1 address
            const setCellValue = (address: string, value: any) => {
                const range = reportSheet.getRange(address);
                range.values = [[value]];
                return range;
            };

            // Add headers and data
            let row = 1;

            // Title and timestamp
            setCellValue(`A${row}`, 'Model Review Tool - Formula Analysis Report');
            reportSheet.getRange(`A${row}`).format.font.bold = true;
            reportSheet.getRange(`A${row}`).format.font.size = 16;
            row += 2;

            setCellValue(`A${row}`, `Generated: ${new Date().toLocaleString()}`);
            row += 2;

            // Summary statistics
            setCellValue(`A${row}`, 'Summary Statistics');
            reportSheet.getRange(`A${row}`).format.font.bold = true;
            row++;

            setCellValue(`A${row}`, 'Total Worksheets:');
            setCellValue(`B${row}`, analysisResults.totalWorksheets);
            row++;

            setCellValue(`A${row}`, 'Total Formulas:');
            setCellValue(`B${row}`, analysisResults.totalFormulas);
            row++;

            setCellValue(`A${row}`, 'Unique Formulas:');
            setCellValue(`B${row}`, analysisResults.uniqueFormulas);
            row++;

            setCellValue(`A${row}`, 'Total Cells:');
            setCellValue(`B${row}`, analysisResults.totalCells);
            row++;

            setCellValue(`A${row}`, 'Cells with Formulas:');
            setCellValue(`B${row}`, analysisResults.totalCellsWithFormulas);
            row++;

            setCellValue(`A${row}`, 'Cells with Values:');
            setCellValue(`B${row}`, analysisResults.totalCellsWithValues);
            row++;

            setCellValue(`A${row}`, 'Empty Cells:');
            setCellValue(`B${row}`, analysisResults.totalCells - analysisResults.totalCellsWithFormulas - analysisResults.totalCellsWithValues);
            row++;

            setCellValue(`A${row}`, 'Hard-coded Values:');
            setCellValue(`B${row}`, analysisResults.totalHardCodedValues);
            row += 2;

            // Worksheet details
            setCellValue(`A${row}`, 'Worksheet Details');
            reportSheet.getRange(`A${row}`).format.font.bold = true;
            row++;

            // Headers
            setCellValue(`A${row}`, 'Worksheet Name');
            setCellValue(`B${row}`, 'Total Formulas');
            setCellValue(`C${row}`, 'Unique Formulas');
            setCellValue(`D${row}`, 'Total Cells');
            setCellValue(`E${row}`, 'Cells with Formulas');
            setCellValue(`F${row}`, 'Cells with Values');
            setCellValue(`G${row}`, 'Empty Cells');
            setCellValue(`H${row}`, 'Hard-coded Values');
            setCellValue(`I${row}`, 'Complexity');

            // Make headers bold
            reportSheet.getRange(`A${row}:I${row}`).format.font.bold = true;
            row++;

            // Add worksheet data
            analysisResults.worksheets.forEach((worksheet: any) => {
                setCellValue(`A${row}`, worksheet.name);
                setCellValue(`B${row}`, worksheet.totalFormulas);
                setCellValue(`C${row}`, worksheet.uniqueFormulas);
                setCellValue(`D${row}`, worksheet.cellCountAnalysis.totalCells);
                setCellValue(`E${row}`, worksheet.cellCountAnalysis.cellsWithFormulas);
                setCellValue(`F${row}`, worksheet.cellCountAnalysis.cellsWithValues);
                setCellValue(`G${row}`, worksheet.cellCountAnalysis.emptyCells);
                setCellValue(`H${row}`, worksheet.hardCodedValueAnalysis.totalHardCodedValues);
                setCellValue(`I${row}`, worksheet.formulaComplexity);
                row++;
            });

            // Auto-fit columns
            const used = reportSheet.getUsedRange(true);
            used.format.autofitColumns();

            await context.sync();

            // Activate the report sheet
            reportSheet.activate();
        });
        
        showStatusMessage('Analysis report exported successfully!', 'success');
        
    } catch (err) {
        console.error('Error exporting results:', err);
        showStatusMessage(`Error exporting results: ${getErrorMessage(err)}`, 'error');
    }
}

/**
 * Generate audit trail
 */
async function generateAuditTrail(): Promise<void> {
    if (!analysisResults) {
        showStatusMessage('No analysis results to generate audit trail. Please run analysis first.', 'warning');
        return;
    }
    
    try {
        showStatusMessage('Generating audit trail...', 'info');
        
        const auditData = {
            timestamp: new Date().toISOString(),
            userAgent: navigator.userAgent,
            analysisResults: analysisResults,
            options: {
                includeEmptyCells: (document.getElementById('includeEmptyCells') as HTMLInputElement).checked,
                groupSimilarFormulas: (document.getElementById('groupSimilarFormulas') as HTMLInputElement).checked
            }
        };
        
        // Create downloadable JSON file
        const dataStr = JSON.stringify(auditData, null, 2);
        const dataBlob = new Blob([dataStr], { type: 'application/json' });
        
        const link = document.createElement('a');
        link.href = URL.createObjectURL(dataBlob);
        link.download = `MRT_Audit_Trail_${new Date().toISOString().split('T')[0]}.json`;
        
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        showStatusMessage('Audit trail generated and downloaded successfully!', 'success');
        
    } catch (err) {
        console.error('Error generating audit trail:', err);
        showStatusMessage(`Error generating audit trail: ${getErrorMessage(err)}`, 'error');
    }
}

/**
 * Show or hide loading indicator
 */
function showLoadingIndicator(show: boolean): void {
    const indicator = document.getElementById('loadingIndicator')!;
    indicator.style.display = show ? 'block' : 'none';
}

/**
 * Show status message
 */
function showStatusMessage(message: string, type: 'success' | 'error' | 'warning' | 'info'): void {
    const container = document.getElementById('statusMessages')!;
    
    const messageElement = document.createElement('div');
    messageElement.className = `status-message status-${type}`;
    messageElement.textContent = message;
    
    container.appendChild(messageElement);
    
    // Auto-remove after 5 seconds
    setTimeout(() => {
        if (messageElement.parentNode) {
            messageElement.parentNode.removeChild(messageElement);
        }
    }, 5000);
}

/**
 * Escape HTML characters
 */
function escapeHtml(text: string): string {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
} 

/**
 * Safely extract a message from an unknown error
 */
function getErrorMessage(err: unknown): string {
    if (err instanceof Error) {
        return err.message;
    }
    try {
        return JSON.stringify(err);
    } catch {
        return String(err);
    }
}