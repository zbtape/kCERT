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
            const results = await analyzer.analyzeWorkbook(context, options);
            
            analysisResults = results;
            displayResults(results);
            showStatusMessage('Formula analysis completed successfully!', 'success');
        });
        
    } catch (error) {
        console.error('Error analyzing formulas:', error);
        showStatusMessage(`Error during analysis: ${error.message}`, 'error');
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
                <span class="worksheet-stat-number">${worksheet.totalCells}</span>
                <span class="worksheet-stat-label">Total Cells</span>
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
            // Create a new worksheet for the report
            const reportSheet = context.workbook.worksheets.add('MRT_Analysis_Report');
            
            // Add headers and data
            let row = 1;
            
            // Title and timestamp
            reportSheet.getCell(row, 1).value = 'Model Review Tool - Formula Analysis Report';
            reportSheet.getCell(row, 1).format.font.bold = true;
            reportSheet.getCell(row, 1).format.font.size = 16;
            row += 2;
            
            reportSheet.getCell(row, 1).value = `Generated: ${new Date().toLocaleString()}`;
            row += 2;
            
            // Summary statistics
            reportSheet.getCell(row, 1).value = 'Summary Statistics';
            reportSheet.getCell(row, 1).format.font.bold = true;
            row++;
            
            reportSheet.getCell(row, 1).value = 'Total Worksheets:';
            reportSheet.getCell(row, 2).value = analysisResults.totalWorksheets;
            row++;
            
            reportSheet.getCell(row, 1).value = 'Total Formulas:';
            reportSheet.getCell(row, 2).value = analysisResults.totalFormulas;
            row++;
            
            reportSheet.getCell(row, 1).value = 'Unique Formulas:';
            reportSheet.getCell(row, 2).value = analysisResults.uniqueFormulas;
            row += 2;
            
            // Worksheet details
            reportSheet.getCell(row, 1).value = 'Worksheet Details';
            reportSheet.getCell(row, 1).format.font.bold = true;
            row++;
            
            // Headers
            reportSheet.getCell(row, 1).value = 'Worksheet Name';
            reportSheet.getCell(row, 2).value = 'Total Formulas';
            reportSheet.getCell(row, 3).value = 'Unique Formulas';
            reportSheet.getCell(row, 4).value = 'Total Cells';
            
            // Make headers bold
            reportSheet.getRange(`A${row}:D${row}`).format.font.bold = true;
            row++;
            
            // Add worksheet data
            analysisResults.worksheets.forEach((worksheet: any) => {
                reportSheet.getCell(row, 1).value = worksheet.name;
                reportSheet.getCell(row, 2).value = worksheet.totalFormulas;
                reportSheet.getCell(row, 3).value = worksheet.uniqueFormulas;
                reportSheet.getCell(row, 4).value = worksheet.totalCells;
                row++;
            });
            
            // Auto-fit columns
            reportSheet.getUsedRange().format.autofitColumns();
            
            await context.sync();
            
            // Activate the report sheet
            reportSheet.activate();
        });
        
        showStatusMessage('Analysis report exported successfully!', 'success');
        
    } catch (error) {
        console.error('Error exporting results:', error);
        showStatusMessage(`Error exporting results: ${error.message}`, 'error');
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
        
    } catch (error) {
        console.error('Error generating audit trail:', error);
        showStatusMessage(`Error generating audit trail: ${error.message}`, 'error');
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