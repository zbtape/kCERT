import './taskpane.css';
import { FormulaAnalyzer, AnalysisResult, WorksheetAnalysisResult, FormulaInfo } from '../shared/FormulaAnalyzer';
import { WorksheetMapGenerator, MapSymbol, MapCounts } from '../shared/WorksheetMapGenerator';

// Wait for Office.js to load
Office.onReady((info) => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById('analyzeFormulas')?.addEventListener('click', analyzeFormulas);
        document.getElementById('generateMaps')?.addEventListener('click', generateMaps);
        document.getElementById('refreshMapSheets')?.addEventListener('click', refreshMapSheetList);
        document.getElementById('selectAllMapSheets')?.addEventListener('click', () => toggleAllMapSheets(true));
        document.getElementById('clearAllMapSheets')?.addEventListener('click', () => toggleAllMapSheets(false));
        document.getElementById('exportResults')?.addEventListener('click', exportResults);
        document.getElementById('generateAuditTrail')?.addEventListener('click', generateAuditTrail);
        document.getElementById('analysisScope')?.addEventListener('change', handleScopeChange);
        document.getElementById('minutesPerFormula')?.addEventListener('change', handleMinutesPerFormulaChange);
        document.getElementById('listingReviewedFilter')?.addEventListener('change', () => {
            uniqueListingState.showReviewedOnly = (document.getElementById('listingReviewedFilter') as HTMLInputElement).checked;
            if (analysisResults) {
                renderUniqueFormulaListing(analysisResults);
            }
        });
        document.getElementById('listingSearch')?.addEventListener('input', (ev) => {
            uniqueListingState.searchTerm = (ev.target as HTMLInputElement).value.trim().toLowerCase();
            if (analysisResults) {
                renderUniqueFormulaListing(analysisResults);
            }
        });
        document.getElementById('exportListing')?.addEventListener('click', exportUniqueFormulaListing);
        document.getElementById('colourApply')?.addEventListener('click', () => applyColouring(false));
        document.getElementById('colourReset')?.addEventListener('click', () => applyColouring(true));
        ANALYSIS_VIEWS.forEach(viewId => {
            const tab = document.getElementById(`analysisTab_${viewId.replace('View','')}`);
            tab?.addEventListener('click', () => switchAnalysisView(viewId));
        });
        initializeScopePicker();
        
        // Components are now initialized
        refreshMapSheetList();
        renderMapRunSummary([]);
    }
});

let analysisResults: AnalysisResult | null = null;
let mapGenerator: WorksheetMapGenerator | null = null;
let mapSheetCache: string[] = [];
let minutesPerFormulaSetting = 2;
let uniqueListingState: { showReviewedOnly: boolean; searchTerm: string } = { showReviewedOnly: false, searchTerm: '' };

const ANALYSIS_VIEWS = ['summaryView', 'listingView', 'mapsView', 'colourView'] as const;
type AnalysisViewId = typeof ANALYSIS_VIEWS[number];

let activeView: AnalysisViewId = 'summaryView';
let sheetScope: string[] = [];

// Fabric UI components are no longer needed - using standard HTML checkboxes

const ESCAPE_MAP: Record<string, string> = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#039;'
};

function ensureCalloutContainer(): HTMLElement {
    const container = document.getElementById('statusMessages') as HTMLElement | null;
    if (container) {
        return container;
    }
    const created = document.createElement('div');
    created.id = 'statusMessages';
    document.body.appendChild(created);
    return created;
}

const formatNumber = (value: number | string | undefined | null): string => {
    if (value === undefined || value === null) {
        return '0';
    }
    const numeric = Number(value);
    if (!Number.isFinite(numeric)) {
        return '0';
    }
    return numeric.toLocaleString();
};

/**
 * Main function to analyze formulas in the workbook
 */
async function analyzeFormulas(): Promise<void> {
    try {
        showLoadingIndicator(true);
        showStatusMessage('Starting formula analysis (vNext massive-model fix)...', 'info');
        
        const options = {
            includeEmptyCells: (document.getElementById('includeEmptyCells') as HTMLInputElement).checked,
            groupSimilarFormulas: (document.getElementById('groupSimilarFormulas') as HTMLInputElement).checked,
            targetSheets: sheetScope.length ? sheetScope : undefined,
            minutesPerFormula: minutesPerFormulaSetting
        };
        
        await Excel.run(async (context) => {
            const analyzer = new FormulaAnalyzer();

            const progress = (message: string) => {
                showStatusMessage(message, 'info');
            };

            showStatusMessage('Analyzing workbook structure (streaming batches)...', 'info');
            const results = await analyzer.analyzeWorkbook(context, options, progress);
            
            // Show specific message for massive models
            if (results.totalCells > 50000) {
                showStatusMessage('MASSIVE MODEL detected - streaming + safe fallback engaged', 'info');
            }
            
            analysisResults = results;
            displayResults(results);
            renderUniqueFormulaSummary(results);
            renderUniqueFormulaListing(results);
            
            const totalCells = results.totalCells;
            if (totalCells > 100000) {
                showStatusMessage(`Analysis completed! Processed ${formatNumber(totalCells)} cells across ${formatNumber(results.totalWorksheets)} worksheets.`, 'success');
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
function displayResults(results: AnalysisResult): void {
    // Update summary statistics
    document.getElementById('totalWorksheets')!.textContent = formatNumber(results.totalWorksheets);
    document.getElementById('totalFormulas')!.textContent = formatNumber(results.totalFormulas);
    document.getElementById('uniqueFormulas')!.textContent = formatNumber(results.uniqueFormulas);
    document.getElementById('totalHardCodedValues')!.textContent = formatNumber(results.totalHardCodedValues);

    // Update cell count analysis
    document.getElementById('totalCells')!.textContent = formatNumber(results.totalCells);
    document.getElementById('cellsWithFormulas')!.textContent = formatNumber(results.totalCellsWithFormulas);
    document.getElementById('cellsWithValues')!.textContent = formatNumber(results.totalCellsWithValues);
    document.getElementById('emptyCells')!.textContent = formatNumber(results.totalCells - results.totalCellsWithFormulas - results.totalCellsWithValues);

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
    
    document.getElementById('highConfidenceValues')!.textContent = formatNumber(highSeverity);
    document.getElementById('mediumConfidenceValues')!.textContent = formatNumber(mediumSeverity);
    document.getElementById('lowConfidenceValues')!.textContent = formatNumber(lowSeverity);
    
    // Display hard-coded values list
    displayHardCodedValues(results.worksheets);
    
    // Display worksheet details
    const worksheetDetails = document.getElementById('worksheetDetails')!;
    worksheetDetails.innerHTML = '';
    
    results.worksheets.forEach((worksheet) => {
        const worksheetCard = createWorksheetCard(worksheet);
        worksheetDetails.appendChild(worksheetCard);
    });
    
    // Show results section
    document.getElementById('resultsSection')!.style.display = 'block';
}

/**
 * Display hard-coded values in the UI
 */
function displayHardCodedValues(worksheets: WorksheetAnalysisResult[]): void {
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
        moreItem.innerHTML = `<div class="hard-coded-value-content">... and ${formatNumber(allHardCodedValues.length - 50)} more hard-coded values</div>`;
        hardCodedValuesList.appendChild(moreItem);
    }
}

/**
 * Create a hard-coded value item element with enhanced inconsistency information
 */
function createHardCodedValueItem(value: any): HTMLElement {
    const item = document.createElement('div');
    item.className = `hard-coded-value-item ${value.severity.toLowerCase()}-confidence`;
    
    const repetitionInfo = value.isRepeated ? ` (Repeated ${formatNumber(value.repetitionCount)} times)` : '';
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
            <div class="worksheet-mode">
                <span class="mode-pill">${worksheet.analysisMode || 'streaming'}</span>
                ${worksheet.fallbackReason ? `<span class="fallback-note">${worksheet.fallbackReason}</span>` : ''}
            </div>
        </div>
        <div class="worksheet-stats">
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${formatNumber(worksheet.totalFormulas)}</span>
                <span class="worksheet-stat-label">Total Formulas</span>
            </div>
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${formatNumber(worksheet.uniqueFormulas)}</span>
                <span class="worksheet-stat-label">Unique Formulas</span>
            </div>
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${formatNumber(worksheet.cellCountAnalysis.totalCells)}</span>
                <span class="worksheet-stat-label">Total Cells</span>
            </div>
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${worksheet.formulaComplexity}</span>
                <span class="worksheet-stat-label">Complexity</span>
            </div>
            <div class="worksheet-stat">
                <span class="worksheet-stat-number">${formatNumber(worksheet.hardCodedValueAnalysis?.totalHardCodedValues || 0)}</span>
                <span class="worksheet-stat-label">Hard-coded</span>
            </div>
        </div>
        <div class="formula-list" id="formulaList_${worksheet.name.replace(/\s+/g, '_')}">
            <h4>Unique Formulas:</h4>
            <div class="worksheet-meta">
                <span class="worksheet-mode">Mode: ${worksheet.analysisMode || 'standard'}${worksheet.fallbackReason ? ` (${worksheet.fallbackReason})` : ''}</span>
            </div>
            ${worksheet.uniqueFormulasList.map((formula: any) => `
                <div class="formula-item">
                    <div class="formula-text">${escapeHtml(formula.formula)}</div>
                    <div class="formula-count">Used ${formatNumber(formula.count)} time(s)</div>
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
            const existing = context.workbook.worksheets.getItemOrNullObject('kCERT_Analysis_Report');
            existing.load('name');
            await context.sync();
            if (!existing.isNullObject) {
                existing.delete();
                await context.sync();
            }

            // Create a new worksheet for the report
            const reportSheet = context.workbook.worksheets.add('kCERT_Analysis_Report');

            // Helper to set a single cell value using A1 address
            const setCellValue = (address: string, value: any) => {
                const range = reportSheet.getRange(address);
                range.values = [[value]];
                return range;
            };

            // Add headers and data
            let row = 1;

            // Title and timestamp
            setCellValue(`A${row}`, 'kCERT - Formula Analysis Report');
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

            setCellValue(`A${row}`, 'Total Hard-coded Values:');
            setCellValue(`B${row}`, analysisResults.totalHardCodedValues);
            row++;

            // Cell count analysis
            setCellValue(`A${row}`, 'Cell Count Analysis');
            reportSheet.getRange(`A${row}`).format.font.bold = true;
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

            // Hard-coded values analysis
            setCellValue(`A${row}`, 'Hard-coded Values Analysis');
            reportSheet.getRange(`A${row}`).format.font.bold = true;
            row++;

            setCellValue(`A${row}`, 'High Severity:');
            setCellValue(`B${row}`, analysisResults.highSeverityValues?.length || 0);
            row++;

            setCellValue(`A${row}`, 'Medium Severity:');
            setCellValue(`B${row}`, analysisResults.mediumSeverityValues?.length || 0);
            row++;

            setCellValue(`A${row}`, 'Low Severity:');
            setCellValue(`B${row}`, analysisResults.lowSeverityValues?.length || 0);
            row++;

            setCellValue(`A${row}`, 'Info Severity:');
            setCellValue(`B${row}`, analysisResults.infoSeverityValues?.length || 0);
            row++;

            // Worksheet details
            setCellValue(`A${row}`, 'Worksheet Details');
            reportSheet.getRange(`A${row}`).format.font.bold = true;
            row++;

            analysisResults.worksheets.forEach((worksheet: any) => {
                setCellValue(`A${row}`, worksheet.name);
                setCellValue(`B${row}`, worksheet.totalFormulas);
                setCellValue(`C${row}`, worksheet.uniqueFormulas);
                setCellValue(`D${row}`, worksheet.cellCountAnalysis.totalCells);
                setCellValue(`E${row}`, worksheet.formulaComplexity);
                setCellValue(`F${row}`, worksheet.hardCodedValueAnalysis?.totalHardCodedValues || 0);
                row++;
            });

            // Formula list (example for one worksheet, adjust for all)
            // This part needs to be generalized for all worksheets
            // For now, we'll just add a placeholder or skip if not applicable
            // setCellValue(`A${row}`, 'Formula List (Example)');
            // reportSheet.getRange(`A${row}`).format.font.bold = true;
            // row++;
            // analysisResults.worksheets.forEach((worksheet: any) => {
            //     setCellValue(`A${row}`, worksheet.name);
            //     worksheet.uniqueFormulasList.forEach((formula: any) => {
            //         setCellValue(`B${row}`, formula.formula);
            //         setCellValue(`C${row}`, formula.count);
            //         row++;
            //     });
            //     row++; // Add a row for the next worksheet
            // });

            await context.sync();
            showStatusMessage('Analysis results exported to new worksheet.', 'success');
        });
    } catch (err) {
        console.error('Error exporting results:', err);
        const errorMsg = getErrorMessage(err);
        showStatusMessage(`Error exporting results: ${errorMsg}`, 'error');
    }
}

/**
 * Generate audit trail for the current workbook
 */
async function generateAuditTrail(): Promise<void> {
    try {
        showLoadingIndicator(true);
        showStatusMessage('Generating audit trail...', 'info');

        await Excel.run(async (context) => {
            const auditTrail = await FormulaAnalyzer.generateAuditTrail(context);
            await context.sync();

            // Assuming auditTrail is an array of strings or a single string
            // For now, we'll just display the first few lines
            const auditTrailText = auditTrail.join('\n');
            showStatusMessage(`Audit Trail generated. Total lines: ${auditTrail.length}`, 'success');
            alert(auditTrailText); // For simplicity, we'll show in an alert
        });
    } catch (err) {
        console.error('Error generating audit trail:', err);
        const errorMsg = getErrorMessage(err);
        showStatusMessage(`Error generating audit trail: ${errorMsg}`, 'error');
    } finally {
        showLoadingIndicator(false);
    }
}

/**
 * Helper to show loading indicator
 */
function showLoadingIndicator(isLoading: boolean): void {
    const loadingIndicator = document.getElementById('loadingIndicator');
    if (loadingIndicator) {
        loadingIndicator.style.display = isLoading ? 'block' : 'none';
    }
}

function getSelectedMapSheets(): string[] {
    const listContainer = document.getElementById('mapSheetChecklist');
    if (!listContainer) {
        return [];
    }

    const selected: string[] = [];
    listContainer.querySelectorAll('input[type="checkbox"]').forEach(box => {
        const input = box as HTMLInputElement;
        if (input.checked) {
            selected.push(input.value);
        }
    });
    return selected;
}

function toggleAllMapSheets(shouldSelect: boolean): void {
    const listContainer = document.getElementById('mapSheetChecklist');
    if (!listContainer) {
        return;
    }

    listContainer.querySelectorAll('input[type="checkbox"]').forEach(box => {
        (box as HTMLInputElement).checked = shouldSelect;
    });
}

async function generateMaps(): Promise<void> {
    try {
        showLoadingIndicator(true);
        showStatusMessage('Generating worksheet maps...', 'info');
        renderMapRunSummary([]);

        const selectedSheets = getSelectedMapSheets();
        if (selectedSheets.length === 0) {
            showStatusMessage('Please select at least one worksheet to map.', 'warning');
            renderMapRunSummary([]);
            return;
        }

        await Excel.run(async (context) => {
            const workbook = context.workbook;
            const worksheets = workbook.worksheets;
            worksheets.load('items/name');
            await context.sync();

            const skipSheets = new Set(['kCERT_Analysis_Report']);
            const targets = worksheets.items.filter(ws => selectedSheets.includes(ws.name) && !skipSheets.has(ws.name));

            if (targets.length === 0) {
                showStatusMessage('No worksheets selected for mapping.', 'warning');
                return;
            }

            if (!mapGenerator) {
                mapGenerator = new WorksheetMapGenerator({ includeHidden: true });
            }

            const summary: string[] = [];
            for (const worksheet of targets) {
                showStatusMessage(`Mapping worksheet "${worksheet.name}"...`, 'info');
                const result = await mapGenerator.generate(context, worksheet, message => {
                    showStatusMessage(message, 'info');
                });

                if (result.skipped) {
                    showStatusMessage(`Skipping ${worksheet.name}: ${result.skipReason}`, 'warning');
                    summary.push(`${worksheet.name}: skipped (${result.skipReason})`);
                    continue;
                }

                await writeMapSheet(context, worksheet.name, result);
                summary.push(`${worksheet.name}: generated (${formatNumber(result.rowCount)}×${formatNumber(result.columnCount)})`);
            }

            showStatusMessage('Worksheet maps generated successfully.', 'success');
            renderMapRunSummary(summary);
        });

        await refreshMapSheetList();
    } catch (error) {
        console.error('Error generating maps', error);
        showStatusMessage(`Error generating maps: ${getErrorMessage(error)}`, 'error');
        renderMapRunSummary([]);
    } finally {
        showLoadingIndicator(false);
    }
}

async function writeMapSheet(
    context: Excel.RequestContext,
    worksheetName: string,
    map: Awaited<ReturnType<WorksheetMapGenerator['generate']>>
): Promise<void> {
    const workbook = context.workbook;
    const sheetName = generateMapSheetName(worksheetName);

    const existing = workbook.worksheets.getItemOrNullObject(sheetName);
    existing.load('isNullObject');
    await context.sync();
    if (!existing.isNullObject) {
        existing.delete();
        await context.sync();
    }

    const mapSheet = workbook.worksheets.add(sheetName);

    const { rowCount, columnCount, symbols, counts, arrayAreas, usedRangeAddress } = map;

    if (rowCount === 0 || columnCount === 0) {
        mapSheet.getRange('A1').values = [[`No used range detected for ${worksheetName}.`]];
        mapSheet.getRange('A2').values = [[`Reason: ${map.skipReason ?? 'unknown'}`]];
        await context.sync();
        return;
    }

    const targetRange = mapSheet.getRangeByIndexes(0, 0, rowCount, columnCount);
    targetRange.values = symbols;
    targetRange.format.font.name = 'Consolas';
    targetRange.format.font.size = 10;
    targetRange.format.columnWidth = 12;
    targetRange.format.rowHeight = 18;
    targetRange.format.horizontalAlignment = 'Center';
    targetRange.format.verticalAlignment = 'Center';

    applyMapColors(targetRange, symbols);

    arrayAreas.forEach(area => {
        const borderRange = mapSheet.getRangeByIndexes(area.top, area.left, area.bottom - area.top + 1, area.right - area.left + 1);
        borderRange.format.borders.getItem('EdgeTop').style = 'Continuous';
        borderRange.format.borders.getItem('EdgeTop').weight = 'Thick';
        borderRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
        borderRange.format.borders.getItem('EdgeBottom').weight = 'Thick';
        borderRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
        borderRange.format.borders.getItem('EdgeLeft').weight = 'Thick';
        borderRange.format.borders.getItem('EdgeRight').style = 'Continuous';
        borderRange.format.borders.getItem('EdgeRight').weight = 'Thick';
    });

    const legendStartRow = rowCount + 2;
    const legendRange = mapSheet.getRange(`A${legendStartRow}:D${legendStartRow + 8}`);

    legendRange.values = buildLegendRows(worksheetName, usedRangeAddress, counts, map.anomalies);
    legendRange.format.font.bold = true;
    legendRange.format.columnWidth = 30;

    await context.sync();
}

function generateMapSheetName(base: string): string {
    const truncated = `${base}_maps`.substring(0, 31);
    if (truncated === base) {
        return `${base}_maps`.substring(0, 31);
    }
    return truncated;
}

function applyMapColors(range: Excel.Range, symbols: MapSymbol[][]): void {
    const rows = symbols.length;
    if (rows === 0) {
        return;
    }
    const cols = symbols[0].length;

    const backgrounds: string[][] = Array.from({ length: rows }, () => new Array<string>(cols).fill(''));
    for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
            const symbol = symbols[r][c];
            backgrounds[r][c] = colorForSymbol(symbol);
        }
    }
    range.format.fill.color = 'white';
    range.format.font.color = '#1b1a19';

    for (let r = 0; r < rows; r++) {
        for (let c = 0; c < cols; c++) {
            const color = backgrounds[r][c];
            if (!color || color === '#FFFFFF') {
                continue;
            }
            range.getCell(r, c).format.fill.color = color;
        }
    }
}

function colorForSymbol(symbol: MapSymbol): string {
    switch (symbol) {
        case 'F':
            return '#7030A0';
        case '<':
            return '#4472C4';
        case '^':
            return '#5B9BD5';
        case '+':
            return '#2F5597';
        case 'L':
            return '#D9D9D9';
        case 'N':
            return '#FFC000';
        case 'A':
            return '#9E480E';
        default:
            return '#FFFFFF';
    }
}

function buildLegendRows(
    worksheetName: string,
    usedRange: string,
    counts: MapCounts,
    anomalies: { changeOfDirection: number; horizontalBreaks: number; verticalBreaks: number }
): string[][] {
    return [
        [`Worksheet`, worksheetName, 'Used Range', usedRange],
        [`Legend`, `F: Unique formula (${counts['F']})`, `^: Copy from above (${counts['^']})`, `+: Copy both (${counts['+']})`],
        ['', `<: Copy from left (${counts['<']})`, `L: Label (${counts['L']})`, `N: Numeric input (${counts['N']})`],
        ['', `A: Array formula (${counts['A']})`, '', ''],
        ['Anomalies', `Direction switches: ${anomalies.changeOfDirection}`, `Horizontal breaks: ${anomalies.horizontalBreaks}`, `Vertical breaks: ${anomalies.verticalBreaks}`],
    ];
}

/**
 * Helper to show status messages
 */
function showStatusMessage(message: string, type: 'info' | 'success' | 'warning' | 'error' = 'info'): void {
    const statusMessage = ensureCalloutContainer();
    statusMessage.textContent = message;
    statusMessage.className = `status-message ${type}`;
    statusMessage.style.display = 'block';
}

/**
 * Helper to get error message from an Office.js error object
 */
function getErrorMessage(err: any): string {
    if (err.message) {
        return err.message;
    }
    if (err.error) {
        return err.error.message;
    }
    return String(err);
}

/**
 * Helper to escape HTML for display
 */
function escapeHtml(value: string): string {
    return value.replace(/[&<>"']/g, (ch) => ESCAPE_MAP[ch]);
}

async function refreshMapSheetList(): Promise<void> {
    const container = document.getElementById('mapSheetChecklist');
    if (!container) {
        return;
    }
    container.innerHTML = '<div class="map-sheet-placeholder">Loading worksheets...</div>';

    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load('items/name');
            await context.sync();

            mapSheetCache = worksheets.items
                .map(ws => ws.name)
                .filter(name => name !== 'kCERT_Analysis_Report' && !name.endsWith('_maps'));

            renderMapSheetChecklist(mapSheetCache);
        });
    } catch (error) {
        console.error('Failed to refresh worksheet list', error);
        container.innerHTML = '<div class="map-sheet-placeholder">Unable to load worksheets</div>';
    }
}

function renderMapSheetChecklist(sheets: string[]): void {
    const container = document.getElementById('mapSheetChecklist');
    if (!container) {
        return;
    }

    if (sheets.length === 0) {
        container.innerHTML = '<div class="map-sheet-placeholder">No worksheets available</div>';
        return;
    }

    const fragment = document.createDocumentFragment();
    sheets.forEach(sheetName => {
        const label = document.createElement('label');
        label.className = 'map-sheet-item';

        const input = document.createElement('input');
        input.type = 'checkbox';
        input.value = sheetName;
        input.checked = true;

        const span = document.createElement('span');
        span.textContent = sheetName;

        label.appendChild(input);
        label.appendChild(span);
        fragment.appendChild(label);
    });

    container.innerHTML = '';
    container.appendChild(fragment);
}

function renderMapRunSummary(summary: string[]): void {
    const summaryContainer = document.getElementById('mapRunSummary');
    if (!summaryContainer) {
        return;
    }

    if (!summary || summary.length === 0) {
        summaryContainer.style.display = 'none';
        summaryContainer.innerHTML = '';
        return;
    }

    summaryContainer.style.display = 'block';
    summaryContainer.innerHTML = `
        <strong>Recent map run:</strong>
        <ul class="map-summary-list">
            ${summary.map(entry => `<li>${escapeHtml(entry)}</li>`).join('')}
        </ul>
    `;
}

function renderUniqueFormulaSummary(results: AnalysisResult): void {
    const overview = document.getElementById('summaryOverview');
    if (!overview) {
        return;
    }
    overview.hidden = false;
    const ufCount = results.uniqueSummary.ufCount;
    minutesPerFormulaSetting = results.uniqueSummary.minutesPerFormula ?? minutesPerFormulaSetting;
    const estimatedMinutes = Math.round(ufCount * minutesPerFormulaSetting);
    const estimatedHours = estimatedMinutes / 60;

    document.getElementById('summaryUniqueCount')!.textContent = formatNumber(ufCount);
    document.getElementById('summaryEstimatedMinutes')!.textContent = formatNumber(estimatedMinutes);
    document.getElementById('summaryEstimatedHours')!.textContent = estimatedHours.toFixed(1);
    document.getElementById('summaryWorksheetsScanned')!.textContent = formatNumber(results.totalWorksheets);
    const minutesField = document.getElementById('minutesPerFormula') as HTMLInputElement | null;
    if (minutesField) {
        minutesField.value = String(minutesPerFormulaSetting);
    }
}

function renderUniqueFormulaListing(results: AnalysisResult): void {
    const tbody = document.getElementById('uniqueFormulaTableBody');
    if (!tbody) {
        return;
    }
    tbody.innerHTML = '';
    const applyFilter = uniqueListingState.showReviewedOnly || uniqueListingState.searchTerm.length > 0;
    const rows: HTMLTableRowElement[] = [];

    results.worksheets.forEach((worksheet) => {
        worksheet.uniqueFormulasList.forEach((formulaInfo, index) => {
            const status = '';
            if (uniqueListingState.showReviewedOnly && !status) {
                return;
            }
            if (uniqueListingState.searchTerm) {
                const search = uniqueListingState.searchTerm;
                const haystack = `${formulaInfo.ufIndicator} ${formulaInfo.normalizedFormula} ${worksheet.name}`.toLowerCase();
                if (!haystack.includes(search)) {
                    return;
                }
            }

            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${formulaInfo.ufIndicator}</td>
                <td>${generateFinPlaceholder(index)}</td>
                <td>${escapeHtml(worksheet.name)}</td>
                <td class="formula-cell">${escapeHtml(formulaInfo.exampleFormula)}</td>
                <td class="formula-cell">${escapeHtml(formulaInfo.normalizedFormula)}</td>
                <td>${formatNumber(formulaInfo.count)}</td>
                <td>${formatNumber(formulaInfo.fScore)}</td>
                <td>${formulaInfo.complexity}</td>
                <td contenteditable="true" data-field="status"></td>
                <td contenteditable="true" data-field="priority"></td>
                <td contenteditable="true" data-field="comment"></td>
                <td contenteditable="true" data-field="clientResponse"></td>
            `;
            rows.push(row);
        });
    });

    if (!rows.length && applyFilter) {
        const emptyRow = document.createElement('tr');
        emptyRow.innerHTML = `<td colspan="11" class="listing-empty">No formulas match current filters.</td>`;
        tbody.appendChild(emptyRow);
        return;
    }

    rows.sort((a, b) => {
        const scoreA = Number(a.children[5].textContent || '0');
        const scoreB = Number(b.children[5].textContent || '0');
        return scoreB - scoreA;
    });

    rows.forEach(row => tbody.appendChild(row));
}

function generateFinPlaceholder(index: number): string {
    return `FIN-${(index + 1).toString().padStart(4, '0')}`;
}

function switchAnalysisView(target: AnalysisViewId): void {
    if (activeView === target) {
        return;
    }
    activeView = target;
    ANALYSIS_VIEWS.forEach(viewId => {
        const panel = document.getElementById(viewId);
        const toolbar = document.querySelector(`[data-view="${viewId}"]`);
        const tab = document.querySelector(`[data-target="${viewId}"]`);
        if (!panel || !tab) {
            return;
        }
        const isActive = viewId === target;
        panel.toggleAttribute('hidden', !isActive);
        (tab as HTMLElement).setAttribute('aria-selected', isActive ? 'true' : 'false');
        if (toolbar instanceof HTMLElement) {
            toolbar.hidden = !isActive;
        }
    });
}

function handleMinutesPerFormulaChange(): void {
    const minutesField = document.getElementById('minutesPerFormula') as HTMLInputElement | null;
    if (!minutesField) {
        return;
    }
    const value = Number(minutesField.value);
    if (!Number.isFinite(value) || value <= 0) {
        minutesField.value = String(minutesPerFormulaSetting);
        return;
    }
    minutesPerFormulaSetting = value;
    if (analysisResults) {
        renderUniqueFormulaSummary({
            ...analysisResults,
            uniqueSummary: {
                ufCount: analysisResults.uniqueSummary.ufCount,
                estimatedMinutes: analysisResults.uniqueSummary.ufCount * minutesPerFormulaSetting,
                minutesPerFormula: minutesPerFormulaSetting
            }
        });
    }
}

function handleScopeChange(): void {
    const scopeSelect = document.getElementById('analysisScope') as HTMLSelectElement | null;
    const picker = document.getElementById('sheetScopePicker') as HTMLElement | null;
    if (!scopeSelect || !picker) {
        return;
    }
    const value = scopeSelect.value;
    if (value === 'sheet') {
        picker.hidden = false;
        populateScopePicker();
    } else {
        picker.hidden = true;
        sheetScope = [];
    }
}

function initializeScopePicker(): void {
    const picker = document.getElementById('sheetScopePicker');
    if (picker) {
        picker.hidden = true;
    }
}

async function populateScopePicker(): Promise<void> {
    const list = document.getElementById('sheetScopeList');
    if (!list) {
        return;
    }
    list.innerHTML = '<div class="sheet-scope-placeholder">Loading worksheets...</div>';
    try {
        await Excel.run(async (context) => {
            const worksheets = context.workbook.worksheets;
            worksheets.load('items/name');
            await context.sync();

            const items = worksheets.items.filter(ws => !ws.name.endsWith('_maps') && ws.name !== 'kCERT_Analysis_Report');
            const fragment = document.createDocumentFragment();
            sheetScope = sheetScope.filter(name => items.some(ws => ws.name === name));
            items.forEach(ws => {
                const label = document.createElement('label');
                label.className = 'sheet-scope-item';
                const input = document.createElement('input');
                input.type = 'checkbox';
                input.value = ws.name;
                input.checked = sheetScope.includes(ws.name);
                input.addEventListener('change', () => {
                    if (input.checked) {
                        if (!sheetScope.includes(ws.name)) {
                            sheetScope.push(ws.name);
                        }
                    } else {
                        sheetScope = sheetScope.filter(name => name !== ws.name);
                    }
                });
                const span = document.createElement('span');
                span.textContent = ws.name;
                label.append(input, span);
                fragment.appendChild(label);
            });
            list.innerHTML = '';
            list.appendChild(fragment);
        });
    } catch (error) {
        console.error('Failed to populate scope picker', error);
        list.innerHTML = '<div class="sheet-scope-placeholder">Unable to load worksheets.</div>';
    }
}

async function exportUniqueFormulaListing(): Promise<void> {
    if (!analysisResults) {
        showStatusMessage('No analysis results to export. Please run analysis first.', 'warning');
        return;
    }

    try {
        await Excel.run(async (context) => {
            const reportSheetName = 'kCERT_UFL';
            const existing = context.workbook.worksheets.getItemOrNullObject(reportSheetName);
            existing.load('isNullObject');
            await context.sync();
            if (!existing.isNullObject) {
                existing.delete();
                await context.sync();
            }

            const sheet = context.workbook.worksheets.add(reportSheetName);
            let row = 1;
            sheet.getRange(`A${row}`).values = [['Unique Formula Listing']];
            sheet.getRange(`A${row}`).format.font.bold = true;
            row += 2;

            sheet.getRange(`A${row}:K${row}`).values = TABLE_HEADER;
            sheet.getRange(`A${row}:K${row}`).format.font.bold = true;
            row++;

            const firstDataRow = row;

            analysisResults.worksheets.forEach(ws => {
                ws.uniqueFormulasList.forEach((info, idx) => {
                    sheet.getRange(`A${row}:K${row}`).values = [[
                        info.ufIndicator,
                        generateFinPlaceholder(idx),
                        ws.name,
                        info.exampleFormula,
                        info.normalizedFormula,
                        info.count,
                        info.fScore,
                        info.complexity,
                        '',
                        '',
                        ''
                    ]];
                    row++;
                });
            });

            const lastDataRow = row - 1;
            if (lastDataRow >= firstDataRow) {
                sheet.tables.add(`A${firstDataRow}:K${lastDataRow}`, true).name = 'UFLTable';
            }
            await context.sync();
        });
        showStatusMessage('Unique Formula Listing exported to worksheet "kCERT_UFL".', 'success');
    } catch (error) {
        console.error('Failed to export listing', error);
        showStatusMessage(`Failed to export listing: ${getErrorMessage(error)}`, 'error');
    }
}

async function applyColouring(resetOnly: boolean): Promise<void> {
    const resetFills = (document.getElementById('colourResetFills') as HTMLInputElement)?.checked ?? false;
    const resetFont = (document.getElementById('colourResetFont') as HTMLInputElement)?.checked ?? false;
    const targetSheets = sheetScope.length ? [...sheetScope] : undefined;

    try {
        await Excel.run(async (context) => {
            const workbook = context.workbook;
            const sheets = workbook.worksheets;
            sheets.load('items/name');
            await context.sync();

            const targets = sheets.items.filter(ws => {
                if (ws.name === 'kCERT_Analysis_Report' || ws.name.endsWith('_maps')) {
                    return false;
                }
                if (targetSheets) {
                    return targetSheets.includes(ws.name);
                }
                return true;
            });

            for (const sheet of targets) {
                const usedRange = sheet.getUsedRangeOrNullObject();
                usedRange.load(['isNullObject', 'address']);
                await context.sync();
                if (usedRange.isNullObject) {
                    continue;
                }
                if (resetOnly) {
                    await resetColours(usedRange, resetFills, resetFont);
                    continue;
                }
                if (resetFills || resetFont) {
                    await resetColours(usedRange, resetFills, resetFont);
                }
                await colourUniqueAndInputs(usedRange);
            }
            await context.sync();
        });
        showStatusMessage(resetOnly ? 'Model colours removed.' : 'Unique and input colouring applied.', 'success');
    } catch (error) {
        console.error('Colouring failed', error);
        showStatusMessage(`Colouring failed: ${getErrorMessage(error)}`, 'error');
    }
}

async function resetColours(range: Excel.Range, resetFills: boolean, resetFont: boolean): Promise<void> {
    if (resetFills) {
        range.format.fill.clear();
    }
    if (resetFont) {
        range.format.font.color = 'Automatic';
    }
}

async function colourUniqueAndInputs(range: Excel.Range): Promise<void> {
    range.load(['formulas', 'values']);
    await range.context.sync();
    const formulas = range.formulas as (string | null)[][];
    const values = range.values as any[][];

    for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
            const formula = formulas[r][c];
            const value = values[r][c];
            const cell = range.getCell(r, c);
            if (typeof formula === 'string' && formula.startsWith('=')) {
                cell.format.fill.color = '#7030A0';
            } else if (typeof value === 'number') {
                cell.format.fill.color = '#FFC000';
            }
        }
    }
}