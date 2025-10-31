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
        document.getElementById('gptSettings')?.addEventListener('click', openGptSettings);
        document.getElementById('saveGptSettings')?.addEventListener('click', saveGptSettings);
        document.getElementById('cancelGptSettings')?.addEventListener('click', closeGptSettings);
        document.getElementById('closeGptSettings')?.addEventListener('click', closeGptSettings);
        document.getElementById('closeGptResponse')?.addEventListener('click', closeGptResponse);
        document.getElementById('closeGptResponseBtn')?.addEventListener('click', closeGptResponse);
        ANALYSIS_VIEWS.forEach(viewId => {
            const tab = document.getElementById(`analysisTab_${viewId.replace('View','')}`);
            tab?.addEventListener('click', () => switchAnalysisView(viewId));
        });
        initializeScopePicker();
        
        // Components are now initialized
        refreshMapSheetList();
        renderMapRunSummary([]);
        
        // Load GPT settings and update button states
        updateAllGptButtons();
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

function setText(id: string, value: string): void {
    const element = document.getElementById(id);
    if (element) {
        element.textContent = value;
    }
}

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
    setText('totalWorksheets', formatNumber(results.totalWorksheets));
    setText('totalFormulas', formatNumber(results.totalFormulas));
    setText('uniqueFormulas', formatNumber(results.uniqueFormulas));
    setText('totalHardCodedValues', formatNumber(results.totalHardCodedValues));

    // Update cell count analysis
    setText('totalCells', formatNumber(results.totalCells));
    setText('cellsWithFormulas', formatNumber(results.totalCellsWithFormulas));
    setText('cellsWithValues', formatNumber(results.totalCellsWithValues));
    setText('emptyCells', formatNumber(results.totalCells - results.totalCellsWithFormulas - results.totalCellsWithValues));

    // Update hard-coded values analysis with null checks
    const highSeverity = results.worksheets.reduce((sum: number, ws: WorksheetAnalysisResult) => {
        const analysis = ws.hardCodedValueAnalysis;
        return sum + (analysis?.highSeverityValues?.length || 0);
    }, 0);
    const mediumSeverity = results.worksheets.reduce((sum: number, ws: WorksheetAnalysisResult) => {
        const analysis = ws.hardCodedValueAnalysis;
        return sum + (analysis?.mediumSeverityValues?.length || 0);
    }, 0);
    const lowSeverity = results.worksheets.reduce((sum: number, ws: WorksheetAnalysisResult) => {
        const analysis = ws.hardCodedValueAnalysis;
        return sum + (analysis?.lowSeverityValues?.length || 0);
    }, 0);
    const infoSeverity = results.worksheets.reduce((sum: number, ws: any) => {
        const analysis = ws.hardCodedValueAnalysis;
        return sum + (analysis?.infoSeverityValues?.length || 0);
    }, 0);
    
    setText('highConfidenceValues', formatNumber(highSeverity));
    setText('mediumConfidenceValues', formatNumber(mediumSeverity));
    setText('lowConfidenceValues', formatNumber(lowSeverity));
    
    // Display hard-coded values list
    displayHardCodedValues(results.worksheets);
    
    // Display worksheet details
    const worksheetDetails = document.getElementById('worksheetDetails');
    if (worksheetDetails) {
        worksheetDetails.innerHTML = '';
        results.worksheets.forEach((worksheet) => {
            const worksheetCard = createWorksheetCard(worksheet);
            worksheetDetails.appendChild(worksheetCard);
        });
    }
    
    // Show results section
    document.getElementById('resultsSection')!.style.display = 'block';
}

/**
 * Display hard-coded values in the UI
 */
function displayHardCodedValues(worksheets: WorksheetAnalysisResult[]): void {
    const hardCodedValuesList = document.getElementById('hardCodedValuesList');
    if (!hardCodedValuesList) {
        return;
    }
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

    setText('summaryUniqueCount', formatNumber(ufCount));
    setText('summaryEstimatedMinutes', formatNumber(estimatedMinutes));
    setText('summaryEstimatedHours', estimatedHours.toFixed(1));
    setText('summaryWorksheetsScanned', formatNumber(results.totalWorksheets));
    const minutesField = document.getElementById('minutesPerFormula') as HTMLInputElement | null;
    if (minutesField) {
        minutesField.value = String(minutesPerFormulaSetting);
    }
}

interface FinRowState {
    ufi: string;
    finIndex: number;
}

let finRowStates = new Map<string, FinRowState[]>(); // Maps UFI to array of FIN states

function renderUniqueFormulaListing(results: AnalysisResult): void {
    const tbody = document.getElementById('uniqueFormulaTableBody');
    if (!tbody) {
        return;
    }
    // Reset FIN states when re-rendering from new analysis
    finRowStates.clear();
    tbody.innerHTML = '';
    const applyFilter = uniqueListingState.showReviewedOnly || uniqueListingState.searchTerm.length > 0;
    const rows: HTMLTableRowElement[] = [];

    results.worksheets.forEach((worksheet) => {
        worksheet.uniqueFormulasList.forEach((formulaInfo) => {
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

            // Create main UFI row
            const row = document.createElement('tr');
            row.setAttribute('data-ufi', formulaInfo.ufIndicator);
            row.setAttribute('data-row-type', 'ufi');
            row.setAttribute('data-worksheet', worksheet.name);
            row.setAttribute('data-cell-address', formulaInfo.cells[0] || '');
            row.setAttribute('data-count', formulaInfo.count.toString());
            row.setAttribute('data-fscore', formulaInfo.fScore.toString());
            row.setAttribute('data-complexity', formulaInfo.complexity);
            
            const ufiCell = document.createElement('td');
            const addFinButton = document.createElement('button');
            addFinButton.type = 'button';
            addFinButton.className = 'add-fin-button';
            addFinButton.textContent = '+';
            addFinButton.title = 'Add FIN (Findings/Issues/Notes)';
            addFinButton.addEventListener('click', () => addFinRow(formulaInfo.ufIndicator, tbody, row));
            ufiCell.appendChild(addFinButton);
            const ufiText = document.createElement('span');
            ufiText.textContent = formulaInfo.ufIndicator;
            ufiCell.appendChild(ufiText);
            
            const finCell = document.createElement('td');
            finCell.textContent = ''; // Empty - FIN only appears in subrows
            
            const sheetCell = document.createElement('td');
            sheetCell.textContent = worksheet.name;
            
            const formulaCell = document.createElement('td');
            formulaCell.className = 'formula-cell clickable-formula';
            formulaCell.textContent = formulaInfo.exampleFormula;
            formulaCell.title = `Click to navigate to ${formulaInfo.cells[0] || 'cell'}`;
            formulaCell.addEventListener('click', () => navigateToCell(worksheet.name, formulaInfo.cells[0]));
            
            const normalizedCell = document.createElement('td');
            normalizedCell.className = 'formula-cell';
            normalizedCell.textContent = formulaInfo.normalizedFormula;
            
            const countCell = document.createElement('td');
            countCell.textContent = formatNumber(formulaInfo.count);
            
            const fScoreCell = document.createElement('td');
            fScoreCell.textContent = formatNumber(formulaInfo.fScore);
            
            const complexityCell = document.createElement('td');
            complexityCell.textContent = formulaInfo.complexity;
            
            const priorityCell = document.createElement('td');
            const prioritySelect = document.createElement('select');
            prioritySelect.className = 'priority-dropdown';
            prioritySelect.setAttribute('data-field', 'priority');
            prioritySelect.innerHTML = `
                <option value="">--</option>
                <option value="High">High</option>
                <option value="Medium">Medium</option>
                <option value="Low">Low</option>
            `;
            prioritySelect.addEventListener('change', (e) => {
                const select = e.target as HTMLSelectElement;
                select.setAttribute('data-selected', select.value);
                if (select.value) {
                    select.style.backgroundColor = 
                        select.value === 'High' ? '#d13438' :
                        select.value === 'Medium' ? '#ffaa44' : '#107c10';
                    select.style.color = 'white';
                } else {
                    select.style.backgroundColor = '';
                    select.style.color = '';
                }
            });
            priorityCell.appendChild(prioritySelect);
            
            const statusCell = document.createElement('td');
            statusCell.contentEditable = 'true';
            statusCell.setAttribute('data-field', 'status');
            
            const commentCell = document.createElement('td');
            commentCell.contentEditable = 'true';
            commentCell.setAttribute('data-field', 'comment');
            
            const clientResponseCell = document.createElement('td');
            clientResponseCell.contentEditable = 'true';
            clientResponseCell.setAttribute('data-field', 'clientResponse');
            
            const askGptCell = document.createElement('td');
            const askGptButton = document.createElement('button');
            askGptButton.type = 'button';
            askGptButton.className = 'ask-gpt-button';
            askGptButton.textContent = 'Ask GPT';
            askGptButton.title = 'Get explanation of this formula';
            askGptButton.addEventListener('click', () => askGpt(formulaInfo.exampleFormula, formulaInfo.ufIndicator));
            askGptCell.appendChild(askGptButton);
            
            // Check if GPT settings are configured and update button state
            updateGptButtonState(askGptButton);
            
            row.appendChild(ufiCell);
            row.appendChild(finCell);
            row.appendChild(sheetCell);
            row.appendChild(formulaCell);
            row.appendChild(normalizedCell);
            row.appendChild(countCell);
            row.appendChild(fScoreCell);
            row.appendChild(complexityCell);
            row.appendChild(priorityCell);
            row.appendChild(statusCell);
            row.appendChild(commentCell);
            row.appendChild(clientResponseCell);
            row.appendChild(askGptCell);
            
            rows.push(row);
            
            // Add any existing FIN rows for this UFI
            const existingFins = finRowStates.get(formulaInfo.ufIndicator) || [];
            existingFins.forEach(finState => {
                const finRow = createFinRow(formulaInfo.ufIndicator, finState.finIndex, tbody, row);
                rows.push(finRow);
            });
        });
    });

    if (!rows.length && applyFilter) {
        const emptyRow = document.createElement('tr');
        emptyRow.innerHTML = `<td colspan="13" class="listing-empty">No formulas match current filters.</td>`;
        tbody.appendChild(emptyRow);
        return;
    }

    // Sort rows, keeping FIN rows with their parent UFI rows
    // Default sort by F-Score
    const sortedUfiRows = rows.filter(r => r.getAttribute('data-row-type') === 'ufi');
    sortedUfiRows.sort((a, b) => {
        const scoreA = Number(a.children[6].textContent || '0');
        const scoreB = Number(b.children[6].textContent || '0');
        return scoreB - scoreA;
    });

    // Insert rows in order: UFI row followed by its FIN rows
    sortedUfiRows.forEach(ufiRow => {
        tbody.appendChild(ufiRow);
        const ufi = ufiRow.getAttribute('data-ufi');
        if (ufi) {
            const finRows = rows.filter(r => 
                r.getAttribute('data-row-type') === 'fin' && 
                r.getAttribute('data-ufi') === ufi
            );
            finRows.forEach(finRow => tbody.appendChild(finRow));
        }
    });
    
    // Attach sort handlers to sortable headers
    attachSortHandlers(tbody);
    
    // Auto-size columns based on content, then attach resize handlers
    autoSizeColumns();
    attachColumnResizeHandlers();
}

let currentSort: { key: string; direction: 'asc' | 'desc' } | null = null;

function attachSortHandlers(tbody: HTMLElement): void {
    const table = tbody.closest('table');
    if (!table) return;
    
    const headers = table.querySelectorAll('th[data-sortable="true"]');
    headers.forEach(header => {
        header.addEventListener('click', () => {
            const sortKey = header.getAttribute('data-sort-key');
            if (!sortKey) return;
            
            // Toggle sort direction
            if (currentSort?.key === sortKey) {
                currentSort.direction = currentSort.direction === 'asc' ? 'desc' : 'asc';
            } else {
                currentSort = { key: sortKey, direction: 'asc' };
            }
            
            sortTable(tbody, sortKey, currentSort.direction);
            updateSortIndicators(table);
        });
    });
}

function sortTable(tbody: HTMLElement, sortKey: string, direction: 'asc' | 'desc'): void {
    // Store FIN rows grouped by UFI before clearing
    const finRowsByUfi = new Map<string, HTMLTableRowElement[]>();
    const allFinRows = Array.from(tbody.querySelectorAll('tr[data-row-type="fin"]')) as HTMLTableRowElement[];
    allFinRows.forEach(finRow => {
        const ufi = finRow.getAttribute('data-ufi');
        if (ufi) {
            if (!finRowsByUfi.has(ufi)) {
                finRowsByUfi.set(ufi, []);
            }
            finRowsByUfi.get(ufi)!.push(finRow);
        }
    });
    
    const allRows = Array.from(tbody.querySelectorAll('tr[data-row-type="ufi"]')) as HTMLTableRowElement[];
    
    allRows.sort((a, b) => {
        let valueA: any;
        let valueB: any;
        
        switch (sortKey) {
            case 'count':
                valueA = Number(a.getAttribute('data-count') || a.children[5].textContent || '0');
                valueB = Number(b.getAttribute('data-count') || b.children[5].textContent || '0');
                break;
            case 'fscore':
                valueA = Number(a.getAttribute('data-fscore') || a.children[6].textContent || '0');
                valueB = Number(b.getAttribute('data-fscore') || b.children[6].textContent || '0');
                break;
            case 'complexity':
                const complexityOrder = { 'High': 3, 'Medium': 2, 'Low': 1 };
                const complexityA = a.getAttribute('data-complexity') || a.children[7].textContent || '';
                const complexityB = b.getAttribute('data-complexity') || b.children[7].textContent || '';
                valueA = complexityOrder[complexityA as keyof typeof complexityOrder] || 0;
                valueB = complexityOrder[complexityB as keyof typeof complexityOrder] || 0;
                break;
            case 'sheet':
                valueA = (a.children[2].textContent || '').toLowerCase();
                valueB = (b.children[2].textContent || '').toLowerCase();
                break;
            default:
                return 0;
        }
        
        if (typeof valueA === 'number' && typeof valueB === 'number') {
            return direction === 'asc' ? valueA - valueB : valueB - valueA;
        } else {
            if (valueA < valueB) return direction === 'asc' ? -1 : 1;
            if (valueA > valueB) return direction === 'asc' ? 1 : -1;
            return 0;
        }
    });
    
    // Clear tbody and re-insert sorted rows with their FIN rows
    tbody.innerHTML = '';
    allRows.forEach(ufiRow => {
        tbody.appendChild(ufiRow);
        const ufi = ufiRow.getAttribute('data-ufi');
        if (ufi) {
            const finRows = finRowsByUfi.get(ufi) || [];
            finRows.forEach(finRow => tbody.appendChild(finRow));
        }
    });
}

function updateSortIndicators(table: HTMLTableElement): void {
    if (!currentSort) return;
    
    const headers = table.querySelectorAll('th[data-sortable="true"]');
    headers.forEach(header => {
        header.classList.remove('sort-asc', 'sort-desc');
        if (header.getAttribute('data-sort-key') === currentSort!.key) {
            header.classList.add(`sort-${currentSort!.direction}`);
        }
    });
}

function autoSizeColumns(): void {
    const table = document.getElementById('uniqueFormulaTable');
    const colgroup = document.getElementById('tableColgroup');
    if (!table || !colgroup) return;
    
    // Clear existing columns
    colgroup.innerHTML = '';
    
    // Get all headers to determine number of columns
    const headers = table.querySelectorAll('thead th');
    const numColumns = headers.length;
    
    // Calculate optimal widths for each column
    const columnWidths: number[] = [];
    const minWidths = [80, 100, 120, 250, 250, 70, 80, 100, 120, 100, 150, 150, 100]; // Minimum widths per column
    const maxWidths = [120, 150, 200, 400, 400, 100, 120, 150, 150, 150, 300, 300, 150]; // Maximum widths per column
    
    // For each column, calculate width based on header text and sample cell content
    for (let colIndex = 0; colIndex < numColumns; colIndex++) {
        let maxWidth = 0;
        
        // Measure header width
        const header = headers[colIndex] as HTMLElement;
        const headerText = header.textContent || '';
        // Estimate: ~8px per character + padding
        const headerWidth = (headerText.length * 8) + 32; // 16px padding on each side
        maxWidth = Math.max(maxWidth, headerWidth);
        
        // Sample a few cells in this column to find the widest content
        const sampleCells = table.querySelectorAll(`tbody td:nth-child(${colIndex + 1})`);
        const sampleSize = Math.min(10, sampleCells.length); // Sample up to 10 cells
        
        for (let i = 0; i < sampleSize; i++) {
            const cell = sampleCells[i] as HTMLElement;
            const cellText = cell.textContent || '';
            
            // For formula cells, use more space
            if (cell.classList.contains('formula-cell')) {
                const estimatedWidth = Math.min(cellText.length * 7 + 32, 400);
                maxWidth = Math.max(maxWidth, estimatedWidth);
            } else {
                // For regular cells, estimate width
                const estimatedWidth = Math.min(cellText.length * 8 + 32, 300);
                maxWidth = Math.max(maxWidth, estimatedWidth);
            }
        }
        
        // Ensure width is within min/max bounds
        const finalWidth = Math.max(minWidths[colIndex] || 80, Math.min(maxWidth, maxWidths[colIndex] || 300));
        columnWidths.push(finalWidth);
        
        // Create col element
        const col = document.createElement('col');
        col.style.width = `${finalWidth}px`;
        colgroup.appendChild(col);
    }
    
    // Apply widths to all cells
    for (let colIndex = 0; colIndex < numColumns; colIndex++) {
        const width = columnWidths[colIndex];
        const cells = table.querySelectorAll(`th:nth-child(${colIndex + 1}), td:nth-child(${colIndex + 1})`);
        cells.forEach(cell => {
            const cellElement = cell as HTMLElement;
            cellElement.style.width = `${width}px`;
            cellElement.style.minWidth = `${width}px`;
        });
    }
}

function attachColumnResizeHandlers(): void {
    const table = document.getElementById('uniqueFormulaTable');
    if (!table) return;
    
    // Remove any existing resize handles
    table.querySelectorAll('.column-resize-handle').forEach(handle => handle.remove());
    
    const headers = table.querySelectorAll('thead th');
    headers.forEach((header, index) => {
        let isResizing = false;
        let startX = 0;
        let startWidth = 0;
        
        const handle = document.createElement('div');
        handle.className = 'column-resize-handle';
        handle.style.cssText = 'position:absolute;right:0;top:0;bottom:0;width:4px;cursor:col-resize;z-index:10;';
        header.appendChild(handle);
        
        handle.addEventListener('mousedown', (e) => {
            isResizing = true;
            startX = e.clientX;
            const headerElement = header as HTMLElement;
            startWidth = headerElement.offsetWidth;
            document.body.style.cursor = 'col-resize';
            document.body.style.userSelect = 'none';
            e.preventDefault();
            e.stopPropagation(); // Prevent triggering sort
        });
        
        const mouseMoveHandler = (e: MouseEvent) => {
            if (!isResizing) return;
            const diff = e.clientX - startX;
            const newWidth = Math.max(40, startWidth + diff); // Minimum width 40px
            
            const headerElement = header as HTMLElement;
            const columnIndex = index;
            
            // Set width on header
            headerElement.style.width = `${newWidth}px`;
            headerElement.style.minWidth = `${newWidth}px`;
            headerElement.style.maxWidth = `${newWidth}px`;
            
            // Apply width to all cells in this column using col element or direct styling
            const colElements = table.querySelectorAll('col');
            if (colElements.length > columnIndex) {
                const colElement = colElements[columnIndex] as HTMLElement;
                colElement.style.width = `${newWidth}px`;
            }
            
            // Also apply to cells directly as fallback
            const cells = table.querySelectorAll(`tbody td:nth-child(${columnIndex + 1}), thead th:nth-child(${columnIndex + 1})`);
            cells.forEach(cell => {
                const cellElement = cell as HTMLElement;
                cellElement.style.width = `${newWidth}px`;
                cellElement.style.minWidth = `${newWidth}px`;
            });
        };
        
        document.addEventListener('mousemove', mouseMoveHandler);
        
        const mouseUpHandler = () => {
            if (isResizing) {
                isResizing = false;
                document.body.style.cursor = '';
                document.body.style.userSelect = '';
                document.removeEventListener('mousemove', mouseMoveHandler);
                document.removeEventListener('mouseup', mouseUpHandler);
            }
        };
        
        document.addEventListener('mouseup', mouseUpHandler);
    });
}

function addFinRow(ufi: string, tbody: HTMLElement, parentRow: HTMLTableRowElement): void {
    const existingFins = finRowStates.get(ufi) || [];
    const finIndex = existingFins.length + 1;
    
    const finState: FinRowState = {
        ufi,
        finIndex
    };
    
    if (!finRowStates.has(ufi)) {
        finRowStates.set(ufi, []);
    }
    finRowStates.get(ufi)!.push(finState);
    
    const finRow = createFinRow(ufi, finIndex, tbody, parentRow);
    tbody.insertBefore(finRow, parentRow.nextSibling);
}

function createFinRow(ufi: string, finIndex: number, tbody: HTMLElement, parentRow: HTMLTableRowElement): HTMLTableRowElement {
    const row = document.createElement('tr');
    row.setAttribute('data-ufi', ufi);
    row.setAttribute('data-row-type', 'fin');
    row.setAttribute('data-fin-index', finIndex.toString());
    row.className = 'fin-subrow';
    
    const ufiCell = document.createElement('td');
    ufiCell.textContent = ''; // Empty for nested rows
    
    const finCell = document.createElement('td');
    finCell.className = 'fin-cell';
    const finText = document.createElement('span');
    finText.textContent = `FIN-${finIndex.toString().padStart(4, '0')}`;
    finCell.appendChild(finText);
    
    // Add delete button for FIN row
    const deleteButton = document.createElement('button');
    deleteButton.type = 'button';
    deleteButton.className = 'delete-fin-button';
    deleteButton.textContent = '×';
    deleteButton.title = 'Delete this FIN';
    deleteButton.addEventListener('click', () => deleteFinRow(ufi, finIndex, row, tbody));
    finCell.appendChild(deleteButton);
    
    const sheetCell = document.createElement('td');
    sheetCell.textContent = ''; // Empty for nested rows
    
    const formulaCell = document.createElement('td');
    formulaCell.textContent = ''; // Empty for nested rows
    
    const normalizedCell = document.createElement('td');
    normalizedCell.textContent = ''; // Empty for nested rows
    
    const countCell = document.createElement('td');
    countCell.textContent = ''; // Empty for nested rows
    
    const fScoreCell = document.createElement('td');
    fScoreCell.textContent = ''; // Empty for nested rows
    
    const complexityCell = document.createElement('td');
    complexityCell.textContent = ''; // Empty for nested rows
    
    const priorityCell = document.createElement('td');
    priorityCell.textContent = ''; // Empty for nested rows
    
    const statusCell = document.createElement('td');
    statusCell.textContent = ''; // Empty for nested rows
    
    const commentCell = document.createElement('td');
    commentCell.contentEditable = 'true';
    commentCell.setAttribute('data-field', 'comment');
    commentCell.setAttribute('data-ufi', ufi);
    commentCell.setAttribute('data-fin', finIndex.toString());
    
    const clientResponseCell = document.createElement('td');
    clientResponseCell.contentEditable = 'true';
    clientResponseCell.setAttribute('data-field', 'clientResponse');
    clientResponseCell.setAttribute('data-ufi', ufi);
    clientResponseCell.setAttribute('data-fin', finIndex.toString());
    
    const askGptCell = document.createElement('td');
    askGptCell.textContent = ''; // Empty for nested rows
    
    row.appendChild(ufiCell);
    row.appendChild(finCell);
    row.appendChild(sheetCell);
    row.appendChild(formulaCell);
    row.appendChild(normalizedCell);
    row.appendChild(countCell);
    row.appendChild(fScoreCell);
    row.appendChild(complexityCell);
    row.appendChild(priorityCell);
    row.appendChild(statusCell);
    row.appendChild(commentCell);
    row.appendChild(clientResponseCell);
    row.appendChild(askGptCell);
    
    return row;
}

function deleteFinRow(ufi: string, finIndex: number, finRow: HTMLTableRowElement, tbody: HTMLElement): void {
    // Remove from state
    const finStates = finRowStates.get(ufi) || [];
    const updatedStates = finStates.filter(state => state.finIndex !== finIndex);
    finRowStates.set(ufi, updatedStates);
    
    // Remove from DOM
    finRow.remove();
    
    // Update remaining FIN indices if needed (for display purposes, but we keep original indices for consistency)
}

// GPT Settings Management
interface GptSettings {
    bearerToken: string;
    engagementCode: string;
    isValidated: boolean;
}

const GPT_SETTINGS_KEY = 'kCERT_gpt_settings';

function getGptSettings(): GptSettings | null {
    try {
        const stored = localStorage.getItem(GPT_SETTINGS_KEY);
        if (stored) {
            return JSON.parse(stored);
        }
    } catch (e) {
        console.error('Error reading GPT settings:', e);
    }
    return null;
}

function saveGptSettingsToStorage(settings: GptSettings): void {
    try {
        localStorage.setItem(GPT_SETTINGS_KEY, JSON.stringify(settings));
    } catch (e) {
        console.error('Error saving GPT settings:', e);
    }
}

function openGptSettings(): void {
    const modal = document.getElementById('gptSettingsModal');
    const bearerInput = document.getElementById('bearerToken') as HTMLInputElement;
    const engagementInput = document.getElementById('engagementCode') as HTMLInputElement;
    const statusDiv = document.getElementById('gptSettingsStatus');
    
    if (modal && bearerInput && engagementInput && statusDiv) {
        const settings = getGptSettings();
        if (settings) {
            bearerInput.value = settings.bearerToken || '';
            engagementInput.value = settings.engagementCode || '';
        } else {
            bearerInput.value = '';
            engagementInput.value = '';
        }
        statusDiv.textContent = '';
        statusDiv.className = 'settings-status';
        modal.style.display = 'flex';
        
        // Close modal when clicking outside
        const clickOutsideHandler = (e: MouseEvent) => {
            if (e.target === modal) {
                closeGptSettings();
                modal.removeEventListener('click', clickOutsideHandler);
            }
        };
        modal.addEventListener('click', clickOutsideHandler);
    }
}

function closeGptSettings(): void {
    const modal = document.getElementById('gptSettingsModal');
    if (modal) {
        modal.style.display = 'none';
    }
}

async function saveGptSettings(): Promise<void> {
    const bearerInput = document.getElementById('bearerToken') as HTMLInputElement;
    const engagementInput = document.getElementById('engagementCode') as HTMLInputElement;
    const statusDiv = document.getElementById('gptSettingsStatus');
    
    if (!bearerInput || !engagementInput || !statusDiv) return;
    
    const bearerToken = bearerInput.value.trim();
    const engagementCode = engagementInput.value.trim();
    
    if (!bearerToken || !engagementCode) {
        statusDiv.textContent = 'Please enter both Bearer Token and Engagement Code.';
        statusDiv.className = 'settings-status error';
        return;
    }
    
    // Validate by making a test request
    statusDiv.textContent = 'Validating credentials...';
    statusDiv.className = 'settings-status info';
    
    try {
        const testResponse = await callGptApi(bearerToken, engagementCode, 'Test validation');
        
        if (testResponse && !testResponse.error) {
            const settings: GptSettings = {
                bearerToken,
                engagementCode,
                isValidated: true
            };
            saveGptSettingsToStorage(settings);
            statusDiv.textContent = 'Settings saved and validated successfully!';
            statusDiv.className = 'settings-status success';
            
            // Update all Ask GPT buttons
            updateAllGptButtons();
            
            // Close modal after a brief delay
            setTimeout(() => {
                closeGptSettings();
            }, 1500);
        } else {
            statusDiv.textContent = `Validation failed: ${testResponse?.error || 'Unknown error'}`;
            statusDiv.className = 'settings-status error';
        }
    } catch (error: any) {
        statusDiv.textContent = `Validation failed: ${error.message || 'Unknown error'}`;
        statusDiv.className = 'settings-status error';
    }
}

function updateGptButtonState(button: HTMLButtonElement): void {
    const settings = getGptSettings();
    if (settings && settings.isValidated) {
        button.disabled = false;
        button.style.opacity = '1';
        button.style.cursor = 'pointer';
    } else {
        button.disabled = true;
        button.style.opacity = '0.6';
        button.style.cursor = 'not-allowed';
        button.title = 'Get explanation of this formula (Configure GPT Settings first)';
    }
}

function updateAllGptButtons(): void {
    const buttons = document.querySelectorAll('.ask-gpt-button');
    buttons.forEach(button => updateGptButtonState(button as HTMLButtonElement));
}

async function callGptApi(bearerToken: string, engagementCode: string, formula: string): Promise<any> {
    // Use local proxy to bypass CORS, or direct API if CORS is configured
    // To use proxy: run 'node proxy-server.js' and set useProxy = true
    const useProxy = true; // Set to false if CORS is configured on the API server
    const url = useProxy 
        ? 'http://localhost:3001/api/gpt'
        : 'https://digitalmatrix-cat.kpmgcloudops.com/workspace/api/v1/generativeai/chat';
    
    // Limit context length
    const context = formula.substring(0, 25000);
    
    const prompt = `Explain this Excel formula in plain English without any formatting. Formula: ${context}`;
    
    const headers = {
        'Content-Type': 'application/json',
        'Accept': '*/*',
        'Authorization': `Bearer ${bearerToken}`
    };
    
    const requestBody = {
        engagementCode: engagementCode,
        modelName: 'GPT 4o Omni',
        provider: 'AzureOpenAI',
        providerModelName: 'gpt-4o',
        parameters: {
            max_tokens: 1024,
            frequency_penalty: 2,
            presence_penalty: 2,
            temperature: 0.9,
            top_p: 1
        },
        prompt: prompt,
        messages: [
            {
                role: 'system',
                content: 'You are a helpful assistant that explains Excel formulas in plain English.'
            },
            {
                role: 'user',
                content: prompt
            }
        ]
    };
    
    try {
        // Log the request for debugging
        console.log('Making GPT API request:', {
            url: url,
            headers: { ...headers, Authorization: 'Bearer ***' }, // Hide token in logs
            body: requestBody
        });
        
        // If using proxy, include bearerToken in body so proxy can forward it in headers
        // Otherwise, use Authorization header directly
        let requestPayload: any;
        let requestHeaders: any;
        
        if (useProxy) {
            // For proxy: send bearerToken in body
            requestPayload = {
                ...requestBody,
                bearerToken: bearerToken
            };
            requestHeaders = {
                'Content-Type': 'application/json'
            };
        } else {
            // Direct API call: use Authorization header
            requestPayload = requestBody;
            requestHeaders = headers;
        }
        
        const response = await fetch(url, {
            method: 'POST',
            headers: requestHeaders,
            body: JSON.stringify(requestPayload),
            mode: 'cors',
            credentials: 'omit'
        });
        
        if (response.status === 401) {
            return { error: 'Authentication failed. Please check your bearer token.' };
        }
        
        if (!response.ok) {
            const errorText = await response.text().catch(() => '');
            return { error: `Request failed with status ${response.status}. ${errorText ? `Response: ${errorText.substring(0, 200)}` : ''}` };
        }
        
        const data = await response.json();
        
        // Log the full response structure for debugging
        console.log('GPT API Response:', JSON.stringify(data, null, 2));
        
        return data;
    } catch (error: any) {
        console.error('GPT API error:', error);
        
        // Provide more specific error messages
        if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
            // This is typically a CORS or network issue
            // Check browser console for more details
            console.error('CORS or Network Error Details:', {
                url: url,
                method: 'POST',
                headers: headers,
                error: error
            });
            
            return { 
                error: 'Failed to connect to GPT API. This is likely a CORS (Cross-Origin Resource Sharing) issue. The API server needs to allow requests from this Excel Add-in domain. Please contact IT support to: 1) Configure CORS headers on the API to allow requests from your add-in domain, or 2) Set up a proxy service to handle the API calls. Check the browser console (F12) for more details.' 
            };
        }
        
        return { error: error.message || 'Failed to connect to GPT API' };
    }
}

async function askGpt(formula: string, ufi: string): Promise<void> {
    const settings = getGptSettings();
    if (!settings || !settings.isValidated) {
        showStatusMessage('Please configure GPT settings first.', 'warning');
        openGptSettings();
        return;
    }
    
    const modal = document.getElementById('gptResponseModal');
    const contentDiv = document.getElementById('gptResponseContent');
    const errorDiv = document.getElementById('gptResponseError');
    
    if (!modal || !contentDiv || !errorDiv) return;
    
    // Show modal with loading state
    modal.style.display = 'flex';
    contentDiv.textContent = 'Getting explanation...';
    errorDiv.style.display = 'none';
    
    // Close modal when clicking outside
    const clickOutsideHandler = (e: MouseEvent) => {
        if (e.target === modal) {
            modal.style.display = 'none';
            modal.removeEventListener('click', clickOutsideHandler);
        }
    };
    modal.addEventListener('click', clickOutsideHandler);
    
    try {
        const response = await callGptApi(settings.bearerToken, settings.engagementCode, formula);
        
        if (response.error) {
            errorDiv.textContent = response.error;
            errorDiv.style.display = 'block';
            contentDiv.textContent = '';
        } else {
            // Extract the explanation from the response
            // Log the response structure to help debug
            console.log('Parsing GPT response. Response type:', typeof response);
            console.log('Response keys:', Object.keys(response || {}));
            console.log('Full response object:', response);
            
            let explanation = '';
            
            // Try different response structures (most common first)
            
            // Structure 1: response.choices[0].message.content (OpenAI format)
            if (response.choices && Array.isArray(response.choices) && response.choices.length > 0) {
                const choice = response.choices[0];
                if (choice.message && choice.message.content) {
                    explanation = choice.message.content;
                    console.log('Found explanation in response.choices[0].message.content');
                } else if (choice.text) {
                    explanation = choice.text;
                    console.log('Found explanation in response.choices[0].text');
                } else if (choice.content) {
                    explanation = choice.content;
                    console.log('Found explanation in response.choices[0].content');
                }
            }
            
            // Structure 2: response.response or response.data
            if (!explanation) {
                if (response.response) {
                    explanation = typeof response.response === 'string' ? response.response : JSON.stringify(response.response);
                    console.log('Found explanation in response.response');
                } else if (response.data) {
                    explanation = typeof response.data === 'string' ? response.data : JSON.stringify(response.data);
                    console.log('Found explanation in response.data');
                }
            }
            
            // Structure 3: Direct text properties
            if (!explanation) {
                if (response.text) {
                    explanation = response.text;
                    console.log('Found explanation in response.text');
                } else if (response.content) {
                    explanation = response.content;
                    console.log('Found explanation in response.content');
                } else if (response.message) {
                    explanation = response.message;
                    console.log('Found explanation in response.message');
                } else if (response.result) {
                    explanation = typeof response.result === 'string' ? response.result : JSON.stringify(response.result);
                    console.log('Found explanation in response.result');
                } else if (response.output) {
                    explanation = typeof response.output === 'string' ? response.output : JSON.stringify(response.output);
                    console.log('Found explanation in response.output');
                }
            }
            
            // Structure 4: String response
            if (!explanation && typeof response === 'string') {
                explanation = response;
                console.log('Response is a string');
            }
            
            // Structure 5: Nested message/content in various paths
            if (!explanation && response.message) {
                const msg = response.message;
                if (typeof msg === 'string') {
                    explanation = msg;
                } else if (msg.content) {
                    explanation = msg.content;
                } else if (msg.text) {
                    explanation = msg.text;
                }
            }
            
            if (explanation && explanation.trim()) {
                // Clean up the explanation - remove any JSON wrapper if it's double-encoded
                let cleanedExplanation = explanation.trim();
                try {
                    // Try to parse if it looks like JSON
                    const parsed = JSON.parse(cleanedExplanation);
                    if (typeof parsed === 'string') {
                        cleanedExplanation = parsed;
                    } else if (parsed.text || parsed.content || parsed.message) {
                        cleanedExplanation = parsed.text || parsed.content || parsed.message;
                    }
                } catch (e) {
                    // Not JSON, use as-is
                }
                
                contentDiv.textContent = cleanedExplanation;
                errorDiv.style.display = 'none';
            } else {
                // Show the raw response structure for debugging
                const responsePreview = JSON.stringify(response, null, 2).substring(0, 1000);
                errorDiv.innerHTML = `
                    <strong>No explanation text found in response.</strong><br><br>
                    <strong>Response structure:</strong><br>
                    <pre style="font-size: 11px; overflow: auto; max-height: 200px;">${escapeHtml(responsePreview)}</pre><br>
                    Please check the browser console (F12) for the full response structure.
                `;
                errorDiv.style.display = 'block';
                contentDiv.textContent = '';
                console.error('Could not extract explanation from response:', response);
            }
        }
    } catch (error: any) {
        errorDiv.textContent = `Error: ${error.message || 'Failed to get explanation'}`;
        errorDiv.style.display = 'block';
        contentDiv.textContent = '';
    }
}

function closeGptResponse(): void {
    const modal = document.getElementById('gptResponseModal');
    if (modal) {
        modal.style.display = 'none';
    }
}

async function navigateToCell(worksheetName: string, cellAddress: string): Promise<void> {
    if (!cellAddress) {
        showStatusMessage('No cell address available for this formula.', 'warning');
        return;
    }
    
    try {
        await Excel.run(async (context) => {
            const worksheet = context.workbook.worksheets.getItem(worksheetName);
            const range = worksheet.getRange(cellAddress);
            range.select();
            await context.sync();
            showStatusMessage(`Navigated to ${worksheetName}!${cellAddress}`, 'success');
        });
    } catch (error) {
        console.error('Error navigating to cell:', error);
        showStatusMessage(`Failed to navigate to cell: ${getErrorMessage(error)}`, 'error');
    }
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

            // Define table headers with correct column order
            const TABLE_HEADER = [
                ['UFI', 'FIN', 'Sheet', 'Formula', 'Formula (normalized)', 'Count', 'F-Score', 'Complexity', 'Priority', 'Status', 'Comment', 'Client Response']
            ];

            sheet.getRange(`A${row}:L${row}`).values = TABLE_HEADER;
            sheet.getRange(`A${row}:L${row}`).format.font.bold = true;
            row++;

            const firstDataRow = row;
            const tbody = document.getElementById('uniqueFormulaTableBody');
            
            if (tbody) {
                // Export from DOM to capture FIN rows and edited data
                const rows = tbody.querySelectorAll('tr[data-row-type="ufi"], tr[data-row-type="fin"]');
                rows.forEach(tr => {
                    const cells = tr.querySelectorAll('td');
                    if (cells.length >= 13) {
                        const rowType = tr.getAttribute('data-row-type');
                        let ufi = '';
                        if (rowType === 'ufi') {
                            // Extract UFI from span or data attribute
                            const ufiSpan = cells[0].querySelector('span');
                            ufi = ufiSpan ? ufiSpan.textContent || '' : tr.getAttribute('data-ufi') || '';
                        }
                        const fin = rowType === 'fin'
                            ? cells[1].textContent || ''
                            : '';
                        const sheetName = rowType === 'ufi' ? cells[2].textContent || '' : '';
                        const formula = rowType === 'ufi' ? cells[3].textContent || '' : '';
                        const normalized = rowType === 'ufi' ? cells[4].textContent || '' : '';
                        const count = rowType === 'ufi' ? cells[5].textContent || '' : '';
                        const fScore = rowType === 'ufi' ? cells[6].textContent || '' : '';
                        const complexity = rowType === 'ufi' ? cells[7].textContent || '' : '';
                        // Extract priority from dropdown
                        const prioritySelect = cells[8].querySelector('select');
                        const priority = prioritySelect ? prioritySelect.value : cells[8].textContent || '';
                        const status = cells[9].textContent || '';
                        const comment = cells[10].textContent || '';
                        const clientResponse = cells[11].textContent || '';
                        
                        sheet.getRange(`A${row}:L${row}`).values = [[
                            ufi,
                            fin,
                            sheetName,
                            formula,
                            normalized,
                            count,
                            fScore,
                            complexity,
                            priority,
                            status,
                            comment,
                            clientResponse
                        ]];
                        row++;
                    }
                });
            } else {
                // Fallback: export from analysis results (no FIN rows)
            analysisResults.worksheets.forEach(ws => {
                    ws.uniqueFormulasList.forEach(info => {
                    sheet.getRange(`A${row}:L${row}`).values = [[
                        info.ufIndicator,
                            '',
                        ws.name,
                        info.exampleFormula,
                        info.normalizedFormula,
                        info.count,
                        info.fScore,
                        info.complexity,
                        '',
                        '',
                        '',
                        ''
                    ]];
                    row++;
                });
            });
            }

            const lastDataRow = row - 1;
            if (lastDataRow >= firstDataRow) {
                sheet.tables.add(`A${firstDataRow}:L${lastDataRow}`, true).name = 'UFLTable';
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