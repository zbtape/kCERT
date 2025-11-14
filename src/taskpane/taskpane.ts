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
        document.getElementById('colorApply')?.addEventListener('click', () => applyColoring(false));
        document.getElementById('colorReset')?.addEventListener('click', () => applyColoring(true));
        document.getElementById('gptSettings')?.addEventListener('click', openGptSettings);
        document.getElementById('saveGptSettings')?.addEventListener('click', saveGptSettings);
        document.getElementById('cancelGptSettings')?.addEventListener('click', closeGptSettings);
        document.getElementById('closeGptSettings')?.addEventListener('click', closeGptSettings);
        document.getElementById('closeGptResponse')?.addEventListener('click', closeGptResponse);
        document.getElementById('closeGptResponseBtn')?.addEventListener('click', closeGptResponse);
        document.getElementById('askGptWorkbook')?.addEventListener('click', () => {
            if (analysisResults) {
                askGptForWorkbook(analysisResults);
            }
        });
        ANALYSIS_VIEWS.forEach(viewId => {
            const tab = document.getElementById(`analysisTab_${viewId.replace('View','')}`);
            tab?.addEventListener('click', () => switchAnalysisView(viewId));
        });
        initializeScopePicker();
        
        // Initialize the view state to ensure toolbars are shown/hidden correctly
        switchAnalysisView('summaryView');
        
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

const ANALYSIS_VIEWS = ['summaryView', 'listingView', 'mapsView', 'colorView'] as const;
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

    // Render workbook-level statistics
    renderWorkbookStats(results);
    
    // Render sheet-by-sheet breakdown
    renderSheetBreakdown(results);
    
    // Update GPT button states after rendering
    updateAllGptButtons();
}

function renderWorkbookStats(results: AnalysisResult): void {
    const workbookStatsSection = document.getElementById('workbookStats');
    if (!workbookStatsSection) {
        return;
    }
    workbookStatsSection.hidden = false;

    // Calculate workbook-level statistics
    const totalFormulas = results.totalFormulas;
    const uniqueFormulas = results.uniqueFormulas;
    const totalWorksheets = results.totalWorksheets;
    
    // Average formulas per sheet
    const avgFormulasPerSheet = totalWorksheets > 0 ? (totalFormulas / totalWorksheets).toFixed(0) : '0';
    
    // Complexity distribution
    const complexityCounts = results.worksheets.reduce((acc: { [key: string]: number }, ws: WorksheetAnalysisResult) => {
        const complexity = ws.formulaComplexity || 'Low';
        acc[complexity] = (acc[complexity] || 0) + 1;
        return acc;
    }, {});
    
    // Update workbook stats
    setText('wbTotalFormulas', formatNumber(totalFormulas));
    setText('wbUniqueFormulas', formatNumber(uniqueFormulas));
    setText('wbAvgFormulasPerSheet', formatNumber(avgFormulasPerSheet));
    
    setText('wbComplexityHigh', formatNumber(complexityCounts['High'] || 0));
    setText('wbComplexityMedium', formatNumber(complexityCounts['Medium'] || 0));
    setText('wbComplexityLow', formatNumber(complexityCounts['Low'] || 0));
}

function renderSheetBreakdown(results: AnalysisResult): void {
    const sheetBreakdownSection = document.getElementById('sheetBreakdown');
    const sheetBreakdownContent = document.getElementById('sheetBreakdownContent');
    if (!sheetBreakdownSection || !sheetBreakdownContent) {
        return;
    }
    
    if (results.worksheets.length === 0) {
        sheetBreakdownSection.hidden = true;
        return;
    }
    
    sheetBreakdownSection.hidden = false;
    sheetBreakdownContent.innerHTML = '';
    
    // Create a table for sheet breakdown
    const table = document.createElement('table');
    table.className = 'sheet-breakdown-table';
    
    // Table header
    const thead = document.createElement('thead');
    thead.innerHTML = `
        <tr>
            <th>Sheet Name</th>
            <th>Formulas</th>
            <th>Unique</th>
            <th>Cells</th>
            <th>Formula %</th>
            <th>Hard-coded</th>
            <th>Complexity</th>
            <th>Ask GPT</th>
        </tr>
    `;
    table.appendChild(thead);
    
    // Table body
    const tbody = document.createElement('tbody');
    results.worksheets.forEach((worksheet: WorksheetAnalysisResult) => {
        const totalCells = worksheet.cellCountAnalysis?.totalCells || worksheet.totalCells || 0;
        const formulas = worksheet.totalFormulas || 0;
        const formulaPercentage = totalCells > 0 ? ((formulas / totalCells) * 100).toFixed(1) : '0.0';
        const hardCoded = worksheet.hardCodedValueAnalysis?.totalHardCodedValues || 0;
        const complexity = worksheet.formulaComplexity || 'Low';
        
        const row = document.createElement('tr');
        
        // Create Ask GPT button cell
        const askGptCell = document.createElement('td');
        askGptCell.className = 'ask-gpt-cell';
        const askGptButton = document.createElement('button');
        askGptButton.type = 'button';
        askGptButton.className = 'ask-gpt-button';
        askGptButton.textContent = 'Ask GPT';
        askGptButton.title = 'Get explanation of this sheet';
        askGptButton.addEventListener('click', () => askGptForSheet(worksheet));
        askGptCell.appendChild(askGptButton);
        updateGptButtonState(askGptButton);
        
        row.innerHTML = `
            <td class="sheet-name-cell">${escapeHtml(worksheet.name)}</td>
            <td class="number-cell">${formatNumber(formulas)}</td>
            <td class="number-cell">${formatNumber(worksheet.uniqueFormulas || 0)}</td>
            <td class="number-cell">${formatNumber(totalCells)}</td>
            <td class="number-cell">${formulaPercentage}%</td>
            <td class="number-cell">${formatNumber(hardCoded)}</td>
            <td class="complexity-cell complexity-${escapeHtml(complexity.toLowerCase())}">${escapeHtml(complexity)}</td>
        `;
        row.appendChild(askGptCell);
        tbody.appendChild(row);
    });
    
    table.appendChild(tbody);
    sheetBreakdownContent.appendChild(table);
    
    // Update GPT button states after rendering
    updateAllGptButtons();
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
            
            // Combined Complexity / F-Score cell
            const complexityFScoreCell = document.createElement('td');
            complexityFScoreCell.className = 'complexity-fscore-cell';
            const complexitySpan = document.createElement('div');
            complexitySpan.className = `complexity-label complexity-${formulaInfo.complexity.toLowerCase()}`;
            complexitySpan.textContent = formulaInfo.complexity;
            const fScoreSpan = document.createElement('div');
            fScoreSpan.className = 'fscore-label';
            fScoreSpan.textContent = `F-Score: ${formatNumber(formulaInfo.fScore)}`;
            complexityFScoreCell.appendChild(complexitySpan);
            complexityFScoreCell.appendChild(fScoreSpan);
            
            // UFI rows have priority/status that are greyed out until FIN exists
            const priorityCell = document.createElement('td');
            const prioritySelect = document.createElement('select');
            prioritySelect.className = 'priority-dropdown';
            prioritySelect.setAttribute('data-field', 'priority');
            prioritySelect.setAttribute('data-ufi', formulaInfo.ufIndicator);
            prioritySelect.disabled = true; // Disabled until FIN exists
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
            priorityCell.classList.add('ufi-priority-status-cell');
            
            const statusCell = document.createElement('td');
            statusCell.setAttribute('data-field', 'status');
            statusCell.setAttribute('data-ufi', formulaInfo.ufIndicator);
            statusCell.classList.add('ufi-priority-status-cell');
            // Will be enabled/disabled based on FIN existence
            // Initially disabled (greyed out) until FIN exists
            
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
            row.appendChild(complexityFScoreCell);
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
            
            // Enable Priority/Status in UFI row if FIN rows exist
            if (existingFins.length > 0) {
                prioritySelect.disabled = false;
                statusCell.contentEditable = 'true';
                priorityCell.classList.remove('ufi-priority-status-greyed');
                statusCell.classList.remove('ufi-priority-status-greyed');
            } else {
                prioritySelect.disabled = true;
                statusCell.removeAttribute('contenteditable');
                priorityCell.classList.add('ufi-priority-status-greyed');
                statusCell.classList.add('ufi-priority-status-greyed');
            }
        });
    });

    if (!rows.length && applyFilter) {
        const emptyRow = document.createElement('tr');
        emptyRow.innerHTML = `<td colspan="13" class="listing-empty">No formulas match current filters.</td>`;
        tbody.appendChild(emptyRow);
        return;
    }

    // Group rows by sheet, then sort within each sheet
    const rowsBySheet = new Map<string, HTMLTableRowElement[]>();
    const ufiRows = rows.filter(r => r.getAttribute('data-row-type') === 'ufi');
    
    // Group UFI rows by sheet
    ufiRows.forEach(ufiRow => {
        const sheetName = ufiRow.getAttribute('data-worksheet') || '';
        if (!rowsBySheet.has(sheetName)) {
            rowsBySheet.set(sheetName, []);
        }
        rowsBySheet.get(sheetName)!.push(ufiRow);
    });
    
    // Sort sheets alphabetically
    const sortedSheets = Array.from(rowsBySheet.keys()).sort();
    
    // Sort UFI rows within each sheet by F-Score (descending)
    sortedSheets.forEach(sheetName => {
        const sheetUfiRows = rowsBySheet.get(sheetName)!;
        sheetUfiRows.sort((a, b) => {
            const scoreA = Number(a.getAttribute('data-fscore') || a.children[6].textContent || '0');
            const scoreB = Number(b.getAttribute('data-fscore') || b.children[6].textContent || '0');
            return scoreB - scoreA;
        });
    });
    
    // Insert rows grouped by sheet with headers
    sortedSheets.forEach((sheetName, sheetIndex) => {
        // Add sheet header row before each sheet group (except first)
        if (sheetIndex > 0) {
            const separatorRow = document.createElement('tr');
            separatorRow.className = 'sheet-separator-row';
            separatorRow.innerHTML = `<td colspan="13" class="sheet-separator-cell"></td>`;
            tbody.appendChild(separatorRow);
        }
        
        // Add sheet header row
        const headerRow = document.createElement('tr');
        headerRow.className = 'sheet-header-row';
        const headerCell = document.createElement('td');
        headerCell.colSpan = 13;
        headerCell.className = 'sheet-header-cell';
        headerCell.textContent = `Sheet: ${sheetName}`;
        headerRow.appendChild(headerCell);
        tbody.appendChild(headerRow);
        
        // Insert UFI rows for this sheet with their FIN rows
        const sheetUfiRows = rowsBySheet.get(sheetName)!;
        sheetUfiRows.forEach(ufiRow => {
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
    // Store sheet headers and separators
    const sheetHeaders = Array.from(tbody.querySelectorAll('tr.sheet-header-row')) as HTMLTableRowElement[];
    const sheetSeparators = Array.from(tbody.querySelectorAll('tr.sheet-separator-row')) as HTMLTableRowElement[];
    
    // Store FIN rows grouped by UFI
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
    
    // Group UFI rows by sheet
    const rowsBySheet = new Map<string, HTMLTableRowElement[]>();
    const allUfiRows = Array.from(tbody.querySelectorAll('tr[data-row-type="ufi"]')) as HTMLTableRowElement[];
    
    allUfiRows.forEach(ufiRow => {
        const sheetName = ufiRow.getAttribute('data-worksheet') || '';
        if (!rowsBySheet.has(sheetName)) {
            rowsBySheet.set(sheetName, []);
        }
        rowsBySheet.get(sheetName)!.push(ufiRow);
    });
    
    // Sort UFI rows within each sheet
    rowsBySheet.forEach((sheetUfiRows, sheetName) => {
        sheetUfiRows.sort((a, b) => {
            let valueA: any;
            let valueB: any;
            
            switch (sortKey) {
                case 'count':
                    valueA = Number(a.getAttribute('data-count') || a.children[5].textContent || '0');
                    valueB = Number(b.getAttribute('data-count') || b.children[5].textContent || '0');
                    break;
                case 'fscore':
                case 'complexity':
                    // Both sort by F-Score (complexity is derived from F-Score)
                    valueA = Number(a.getAttribute('data-fscore') || '0');
                    valueB = Number(b.getAttribute('data-fscore') || '0');
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
    });
    
    // Sort sheets alphabetically (unless sorting by sheet)
    const sortedSheets = Array.from(rowsBySheet.keys()).sort();
    if (sortKey === 'sheet') {
        sortedSheets.sort((a, b) => {
            return direction === 'asc' ? a.localeCompare(b) : b.localeCompare(a);
        });
    }
    
    // Clear tbody and re-insert sorted rows grouped by sheet with headers
    tbody.innerHTML = '';
    sortedSheets.forEach((sheetName, sheetIndex) => {
        // Add separator before each sheet group (except first)
        if (sheetIndex > 0) {
            const separatorRow = document.createElement('tr');
            separatorRow.className = 'sheet-separator-row';
            separatorRow.innerHTML = `<td colspan="13" class="sheet-separator-cell"></td>`;
            tbody.appendChild(separatorRow);
        }
        
        // Add sheet header
        const headerRow = document.createElement('tr');
        headerRow.className = 'sheet-header-row';
        const headerCell = document.createElement('td');
        headerCell.colSpan = 13;
        headerCell.className = 'sheet-header-cell';
        headerCell.textContent = `Sheet: ${sheetName}`;
        headerRow.appendChild(headerCell);
        tbody.appendChild(headerRow);
        
        // Insert UFI rows for this sheet with their FIN rows
        const sheetUfiRows = rowsBySheet.get(sheetName)!;
        sheetUfiRows.forEach(ufiRow => {
            tbody.appendChild(ufiRow);
            const ufi = ufiRow.getAttribute('data-ufi');
            if (ufi) {
                const finRows = finRowsByUfi.get(ufi) || [];
                finRows.forEach(finRow => tbody.appendChild(finRow));
            }
        });
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
    
    // Enable Priority/Status in parent UFI row when FIN is added
    updateUfiRowPriorityStatusState(ufi, tbody);
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
    
    const complexityFScoreCell = document.createElement('td');
    complexityFScoreCell.textContent = ''; // Empty for nested rows
    
    // FIN rows have priority and status
    const priorityCell = document.createElement('td');
    const prioritySelect = document.createElement('select');
    prioritySelect.className = 'priority-dropdown';
    prioritySelect.setAttribute('data-field', 'priority');
    prioritySelect.setAttribute('data-ufi', ufi);
    prioritySelect.setAttribute('data-fin', finIndex.toString());
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
        // Update parent UFI row Priority/Status state when FIN priority changes
        updateUfiRowPriorityStatusState(ufi, tbody);
    });
    priorityCell.appendChild(prioritySelect);
    
    const statusCell = document.createElement('td');
    statusCell.contentEditable = 'true';
    statusCell.setAttribute('data-field', 'status');
    statusCell.setAttribute('data-ufi', ufi);
    statusCell.setAttribute('data-fin', finIndex.toString());
    
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
    row.appendChild(complexityFScoreCell);
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
    
    // Update UFI row Priority/Status state - grey out if no FIN rows remain
    updateUfiRowPriorityStatusState(ufi, tbody);
}

/**
 * Update the Priority/Status editable state of a UFI row based on whether FIN rows exist
 */
function updateUfiRowPriorityStatusState(ufi: string, tbody: HTMLElement): void {
    const ufiRow = tbody.querySelector(`tr[data-row-type="ufi"][data-ufi="${ufi}"]`) as HTMLTableRowElement;
    if (!ufiRow) return;
    
    const cells = ufiRow.querySelectorAll('td');
    if (cells.length < 10) return;
    
    const priorityCell = cells[8]; // Priority is column 9 (0-indexed: 8)
    const statusCell = cells[9]; // Status is column 10 (0-indexed: 9)
    const prioritySelect = priorityCell.querySelector('select') as HTMLSelectElement;
    
    const finRows = tbody.querySelectorAll(`tr[data-row-type="fin"][data-ufi="${ufi}"]`);
    
    if (finRows.length > 0) {
        // Enable Priority/Status when FIN exists
        if (prioritySelect) {
            prioritySelect.disabled = false;
        }
        statusCell.contentEditable = 'true';
        priorityCell.classList.remove('ufi-priority-status-greyed');
        statusCell.classList.remove('ufi-priority-status-greyed');
    } else {
        // Grey out Priority/Status when no FIN exists
        if (prioritySelect) {
            prioritySelect.disabled = true;
        }
        statusCell.removeAttribute('contenteditable');
        priorityCell.classList.add('ufi-priority-status-greyed');
        statusCell.classList.add('ufi-priority-status-greyed');
    }
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
    const buttons = document.querySelectorAll('.ask-gpt-button, .ask-gpt-workbook-button');
    buttons.forEach(button => updateGptButtonState(button as HTMLButtonElement));
}

async function callGptApi(bearerToken: string, engagementCode: string, formula: string): Promise<any> {
    // Use local proxy to bypass CORS, or direct API if CORS is configured
    // To use proxy: run 'node proxy-server.js' and set useProxy = true
    const useProxy = true; // Set to false if CORS is configured on the API server
    const url = useProxy 
        ? 'http://localhost:3001/api/gpt'
        : 'https://digitalmatrix-cat.kpmgcloudops.com/workspace/api/v1/generativeai/chat';
    
    // Trim all inputs to avoid whitespace errors
    const trimmedBearerToken = bearerToken.trim();
    const trimmedEngagementCode = engagementCode.trim();
    const trimmedFormula = formula.trim();
    
    // Limit context length after trimming
    const context = trimmedFormula.substring(0, 25000);
    
    const prompt = `Explain this Excel formula in plain English without any formatting. Formula: ${context}`.trim();
    
    const headers = {
        'Content-Type': 'application/json',
        'Accept': '*/*',
        'Authorization': `Bearer ${trimmedBearerToken}`
    };
    
    const requestBody = {
        engagementCode: trimmedEngagementCode,
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
                bearerToken: trimmedBearerToken
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
                // Check if message is a string (direct message)
                if (choice.message && typeof choice.message === 'string') {
                    explanation = choice.message;
                    console.log('Found explanation in response.choices[0].message (string)');
                } else if (choice.message && choice.message.content) {
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
                
                // Convert markdown to HTML for better formatting
                contentDiv.innerHTML = markdownToHtml(cleanedExplanation);
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
        // Reset modal title
        const modalTitle = modal.querySelector('.modal-header h3');
        if (modalTitle) {
            modalTitle.textContent = 'Formula Explanation';
        }
    }
}

/**
 * Convert markdown to HTML for display
 */
function markdownToHtml(markdown: string): string {
    let html = markdown.trim();
    
    // Escape HTML first to prevent XSS
    html = escapeHtml(html);
    
    // Clean up common markdown artifacts and malformed patterns
    // Remove patterns like "4.*aaa*" or similar malformed markdown
    html = html.replace(/\d+\.\*[^*]*\*/g, '');
    
    // Process in order: headers, code, bold, italic, lists, paragraphs
    
    // Headers: # Header, ## Header, ### Header (must be at start of line)
    html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
    html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
    html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');
    
    // Code blocks: `code` (inline code)
    html = html.replace(/`([^`\n]+?)`/g, '<code>$1</code>');
    
    // Bold: **text** or __text__ (process before italic to avoid conflicts)
    html = html.replace(/\*\*([^*\n]+?)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/__([^_\n]+?)__/g, '<strong>$1</strong>');
    
    // Italic: *text* or _text_ (single asterisk/underscore, not part of bold)
    // More careful matching to avoid false positives
    html = html.replace(/(^|[^*\s])\*([^*\s\n][^*\n]*?[^*\s\n])\*([^*\s]|$)/g, '$1<em>$2</em>$3');
    html = html.replace(/(^|[^_\s])_([^_\s\n][^_\n]*?[^_\s\n])_([^_\s]|$)/g, '$1<em>$2</em>$3');
    
    // Handle special case: *Purpose* at start of line or after colon
    html = html.replace(/(:\s*|\n\s*)\*([^*\n]+?)\*/g, '$1<em>$2</em>');
    
    // Lists: - item or * item (must be at start of line or after whitespace)
    // Process lists line by line
    const lines = html.split('\n');
    const processedLines: string[] = [];
    let inList = false;
    let listType: 'ul' | 'ol' | null = null;
    
    for (let i = 0; i < lines.length; i++) {
        let line = lines[i];
        const trimmedLine = line.trim();
        
        // Check for bullet list item
        const bulletMatch = trimmedLine.match(/^[-*]\s+(.+)$/);
        // Check for numbered list item
        const numberedMatch = trimmedLine.match(/^\d+\.\s+(.+)$/);
        
        if (bulletMatch) {
            if (!inList || listType !== 'ul') {
                if (inList && listType === 'ol') {
                    processedLines.push('</ol>');
                } else if (inList) {
                    processedLines.push('</ul>');
                }
                processedLines.push('<ul>');
                inList = true;
                listType = 'ul';
            }
            processedLines.push('<li>' + bulletMatch[1] + '</li>');
        } else if (numberedMatch) {
            if (!inList || listType !== 'ol') {
                if (inList && listType === 'ul') {
                    processedLines.push('</ul>');
                } else if (inList) {
                    processedLines.push('</ol>');
                }
                processedLines.push('<ol>');
                inList = true;
                listType = 'ol';
            }
            processedLines.push('<li>' + numberedMatch[1] + '</li>');
        } else {
            // Not a list item
            if (inList) {
                if (listType === 'ul') {
                    processedLines.push('</ul>');
                } else {
                    processedLines.push('</ol>');
                }
                inList = false;
                listType = null;
            }
            // Only add non-empty lines
            if (trimmedLine || line.includes('<')) {
                processedLines.push(line);
            }
        }
    }
    
    // Close any open list
    if (inList) {
        if (listType === 'ul') {
            processedLines.push('</ul>');
        } else {
            processedLines.push('</ol>');
        }
    }
    
    html = processedLines.join('\n');
    
    // Convert double newlines to paragraph breaks, but preserve HTML structure
    // Split by double newlines, but be careful not to break HTML tags
    const sections = html.split(/\n\n+/);
    html = sections.map(section => {
        section = section.trim();
        if (!section) return '';
        
        // Don't wrap if it's already a block-level HTML element
        if (section.match(/^<(h[1-6]|ul|ol|li|p|div|code|pre)/)) {
            return section;
        }
        
        // Don't wrap if it contains block-level elements
        if (section.includes('<h') || section.includes('<ul') || section.includes('<ol') || section.includes('<li')) {
            return section;
        }
        
        // Wrap in paragraph
        return '<p>' + section + '</p>';
    }).join('\n\n');
    
    // Convert remaining single newlines to <br>, but not inside HTML tags
    html = html.replace(/([^>])\n([^<])/g, '$1<br>$2');
    
    // Clean up any double <br> tags
    html = html.replace(/<br>\s*<br>/g, '<br>');
    
    // Clean up any empty paragraphs
    html = html.replace(/<p>\s*<\/p>/g, '');
    html = html.replace(/<p><br><\/p>/g, '');
    
    return html;
}

/**
 * Gather sheet data for GPT analysis
 */
async function gatherSheetData(worksheet: WorksheetAnalysisResult): Promise<string> {
    let sheetData = `Sheet Name: ${worksheet.name}\n`;
    sheetData += `Total Formulas: ${worksheet.totalFormulas}\n`;
    sheetData += `Unique Formulas: ${worksheet.uniqueFormulas}\n`;
    sheetData += `Total Cells: ${worksheet.cellCountAnalysis?.totalCells || worksheet.totalCells || 0}\n`;
    sheetData += `Complexity: ${worksheet.formulaComplexity || 'Low'}\n`;
    sheetData += `Hard-coded Values: ${worksheet.hardCodedValueAnalysis?.totalHardCodedValues || 0}\n\n`;
    
    // Get sample formulas from the sheet
    try {
        await Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getItem(worksheet.name);
            const usedRange = sheet.getUsedRange();
            usedRange.load(['address', 'formulas']);
            await context.sync();
            
            if (usedRange.formulas) {
                sheetData += `Sample Formulas:\n`;
                const formulas: string[] = [];
                const maxFormulas = 50; // Limit to avoid token limits
                
                for (let i = 0; i < usedRange.formulas.length && formulas.length < maxFormulas; i++) {
                    for (let j = 0; j < usedRange.formulas[i].length && formulas.length < maxFormulas; j++) {
                        const formula = usedRange.formulas[i][j];
                        if (formula && typeof formula === 'string' && formula.startsWith('=')) {
                            formulas.push(formula);
                        }
                    }
                }
                
                // Get unique formulas
                const uniqueFormulas = Array.from(new Set(formulas));
                uniqueFormulas.slice(0, 30).forEach((formula, idx) => {
                    sheetData += `${idx + 1}. ${formula.trim()}\n`;
                });
                
                if (uniqueFormulas.length > 30) {
                    sheetData += `... and ${uniqueFormulas.length - 30} more unique formulas\n`;
                }
            }
        });
    } catch (error) {
        console.warn('Could not gather sheet formulas:', error);
        // Use unique formulas list from analysis if available
        if (worksheet.uniqueFormulasList && worksheet.uniqueFormulasList.length > 0) {
            sheetData += `Sample Formulas:\n`;
            worksheet.uniqueFormulasList.slice(0, 30).forEach((formulaInfo, idx) => {
                sheetData += `${idx + 1}. ${formulaInfo.formula.trim()}\n`;
            });
        }
    }
    
    // Final trim to remove any trailing whitespace
    return sheetData.trim();
}

/**
 * Gather workbook summary data for GPT analysis
 */
function gatherWorkbookData(results: AnalysisResult): string {
    let workbookData = `Workbook Summary:\n`;
    workbookData += `Total Worksheets: ${results.totalWorksheets}\n`;
    workbookData += `Total Formulas: ${results.totalFormulas}\n`;
    workbookData += `Unique Formulas: ${results.uniqueFormulas}\n`;
    workbookData += `Total Cells: ${results.totalCells}\n\n`;
    
    workbookData += `Worksheets:\n`;
    results.worksheets.forEach((ws, idx) => {
        workbookData += `${idx + 1}. ${ws.name}\n`;
        workbookData += `   - Formulas: ${ws.totalFormulas}, Unique: ${ws.uniqueFormulas}\n`;
        workbookData += `   - Cells: ${ws.cellCountAnalysis?.totalCells || ws.totalCells || 0}\n`;
        workbookData += `   - Complexity: ${ws.formulaComplexity || 'Low'}\n`;
        workbookData += `   - Hard-coded Values: ${ws.hardCodedValueAnalysis?.totalHardCodedValues || 0}\n`;
        
        // Add sample formulas from each sheet
        if (ws.uniqueFormulasList && ws.uniqueFormulasList.length > 0) {
            workbookData += `   - Sample Formulas:\n`;
            ws.uniqueFormulasList.slice(0, 5).forEach((formulaInfo, fIdx) => {
                const formula = formulaInfo.formula.trim();
                workbookData += `     ${fIdx + 1}. ${formula.substring(0, 100)}${formula.length > 100 ? '...' : ''}\n`;
            });
            if (ws.uniqueFormulasList.length > 5) {
                workbookData += `     ... and ${ws.uniqueFormulasList.length - 5} more\n`;
            }
        }
        workbookData += `\n`;
    });
    
    // Final trim to remove any trailing whitespace
    return workbookData.trim();
}

/**
 * Ask GPT to explain a specific sheet
 */
async function askGptForSheet(worksheet: WorksheetAnalysisResult): Promise<void> {
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
    
    // Update modal title
    const modalTitle = modal.querySelector('.modal-header h3');
    if (modalTitle) {
        modalTitle.textContent = `Sheet Explanation: ${worksheet.name}`;
    }
    
    // Show modal with loading state
    modal.style.display = 'flex';
    contentDiv.textContent = 'Analyzing sheet and getting explanation...';
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
        // Gather sheet data
        const sheetData = await gatherSheetData(worksheet);
        
        // Call GPT API with sheet explanation prompt
        const response = await callGptApiForExplanation(
            settings.bearerToken,
            settings.engagementCode,
            sheetData,
            `Explain what this Excel worksheet does based on its structure, formulas, and data. Worksheet: ${worksheet.name}`
        );
        
        if (response.error) {
            errorDiv.textContent = response.error;
            errorDiv.style.display = 'block';
            contentDiv.textContent = '';
        } else {
            const explanation = extractExplanationFromResponse(response);
            if (explanation && explanation.trim()) {
                // Convert markdown to HTML for better formatting
                contentDiv.innerHTML = markdownToHtml(explanation.trim());
                errorDiv.style.display = 'none';
            } else {
                errorDiv.textContent = 'Could not extract explanation from response. Please check the browser console for details.';
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

/**
 * Ask GPT to explain the entire workbook
 */
async function askGptForWorkbook(results: AnalysisResult): Promise<void> {
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
    
    // Update modal title
    const modalTitle = modal.querySelector('.modal-header h3');
    if (modalTitle) {
        modalTitle.textContent = 'Workbook Explanation';
    }
    
    // Show modal with loading state
    modal.style.display = 'flex';
    contentDiv.textContent = 'Analyzing workbook and getting explanation...';
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
        // Gather workbook data
        const workbookData = gatherWorkbookData(results);
        
        // Call GPT API with workbook explanation prompt
        const response = await callGptApiForExplanation(
            settings.bearerToken,
            settings.engagementCode,
            workbookData,
            'Explain what this Excel workbook does at a high level based on its structure, worksheets, formulas, and data. Provide an overview of the purpose and functionality of the workbook.'
        );
        
        if (response.error) {
            errorDiv.textContent = response.error;
            errorDiv.style.display = 'block';
            contentDiv.textContent = '';
        } else {
            const explanation = extractExplanationFromResponse(response);
            if (explanation && explanation.trim()) {
                // Convert markdown to HTML for better formatting
                contentDiv.innerHTML = markdownToHtml(explanation.trim());
                errorDiv.style.display = 'none';
            } else {
                errorDiv.textContent = 'Could not extract explanation from response. Please check the browser console for details.';
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

/**
 * Extract explanation from GPT response (reusable function)
 */
function extractExplanationFromResponse(response: any): string {
    let explanation = '';
    
    // Try different response structures (same as askGpt function)
    if (response.choices && Array.isArray(response.choices) && response.choices.length > 0) {
        const choice = response.choices[0];
        if (choice.message && typeof choice.message === 'string') {
            explanation = choice.message;
        } else if (choice.message && choice.message.content) {
            explanation = choice.message.content;
        } else if (choice.text) {
            explanation = choice.text;
        } else if (choice.content) {
            explanation = choice.content;
        }
    }
    
    if (!explanation) {
        if (response.response) {
            explanation = typeof response.response === 'string' ? response.response : JSON.stringify(response.response);
        } else if (response.data) {
            explanation = typeof response.data === 'string' ? response.data : JSON.stringify(response.data);
        }
    }
    
    if (!explanation) {
        if (response.text) {
            explanation = response.text;
        } else if (response.content) {
            explanation = response.content;
        } else if (response.message) {
            explanation = response.message;
        } else if (response.result) {
            explanation = typeof response.result === 'string' ? response.result : JSON.stringify(response.result);
        } else if (response.output) {
            explanation = typeof response.output === 'string' ? response.output : JSON.stringify(response.output);
        }
    }
    
    if (!explanation && typeof response === 'string') {
        explanation = response;
    }
    
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
    
    // Clean up the explanation
    if (explanation && explanation.trim()) {
        let cleanedExplanation = explanation.trim();
        try {
            const parsed = JSON.parse(cleanedExplanation);
            if (typeof parsed === 'string') {
                cleanedExplanation = parsed;
            } else if (parsed.text || parsed.content || parsed.message) {
                cleanedExplanation = parsed.text || parsed.content || parsed.message;
            }
        } catch (e) {
            // Not JSON, use as-is
        }
        return cleanedExplanation;
    }
    
    return '';
}

/**
 * Call GPT API for explanations (sheet or workbook)
 */
async function callGptApiForExplanation(bearerToken: string, engagementCode: string, context: string, prompt: string): Promise<any> {
    const useProxy = true;
    const url = useProxy 
        ? 'http://localhost:3001/api/gpt'
        : 'https://digitalmatrix-cat.kpmgcloudops.com/workspace/api/v1/generativeai/chat';
    
    // Trim all inputs to avoid whitespace errors
    const trimmedBearerToken = bearerToken.trim();
    const trimmedEngagementCode = engagementCode.trim();
    const trimmedContext = context.trim();
    const trimmedPrompt = prompt.trim();
    
    // Limit context length after trimming
    const limitedContext = trimmedContext.substring(0, 20000);
    
    const fullPrompt = `${trimmedPrompt}\n\nContext:\n${limitedContext}`.trim();
    
    const headers = {
        'Content-Type': 'application/json',
        'Accept': '*/*',
        'Authorization': `Bearer ${trimmedBearerToken}`
    };
    
    const requestBody = {
        engagementCode: trimmedEngagementCode,
        modelName: 'GPT 4o Omni',
        provider: 'AzureOpenAI',
        providerModelName: 'gpt-4o',
        parameters: {
            max_tokens: 2048,
            frequency_penalty: 2,
            presence_penalty: 2,
            temperature: 0.7,
            top_p: 1
        },
        prompt: fullPrompt,
        messages: [
            {
                role: 'system',
                content: 'You are a helpful assistant that explains Excel workbooks and worksheets in plain English. Provide clear, concise explanations that help reviewers understand the purpose and functionality of the workbook or worksheet.'
            },
            {
                role: 'user',
                content: fullPrompt
            }
        ]
    };
    
    try {
        let requestPayload: any;
        let requestHeaders: any;
        
        if (useProxy) {
            requestPayload = {
                ...requestBody,
                bearerToken: trimmedBearerToken
            };
            requestHeaders = {
                'Content-Type': 'application/json'
            };
        } else {
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
        console.log('GPT API Response:', JSON.stringify(data, null, 2));
        
        return data;
    } catch (error: any) {
        console.error('GPT API error:', error);
        
        if (error.name === 'TypeError' && error.message.includes('Failed to fetch')) {
            return { 
                error: 'Failed to connect to GPT API. This is likely a CORS (Cross-Origin Resource Sharing) issue. The API server needs to allow requests from this Excel Add-in domain. Please contact IT support to: 1) Configure CORS headers on the API to allow requests from your add-in domain, or 2) Set up a proxy service to handle the API calls. Check the browser console (F12) for more details.' 
            };
        }
        
        return { error: error.message || 'Failed to connect to GPT API' };
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
    // Always update the state, even if it's the same view (needed for initial load)
    activeView = target;
    
    // Update panels and tabs
    ANALYSIS_VIEWS.forEach(viewId => {
        const panel = document.getElementById(viewId);
        const tab = document.querySelector(`[data-target="${viewId}"]`);
        if (!panel || !tab) {
            return;
        }
        const isActive = viewId === target;
        panel.toggleAttribute('hidden', !isActive);
        (tab as HTMLElement).setAttribute('aria-selected', isActive ? 'true' : 'false');
    });
    
    // Update toolbars - handle both single and multiple data-view values
    const allToolbars = document.querySelectorAll('[data-view]');
    allToolbars.forEach(toolbar => {
        if (!(toolbar instanceof HTMLElement)) {
            return;
        }
        const dataView = toolbar.getAttribute('data-view');
        if (!dataView) {
            return;
        }
        
        // Check if this toolbar should be visible for the current view
        // Support comma-separated list of views
        const allowedViews = dataView.split(',').map(v => v.trim());
        const shouldShow = allowedViews.includes(target);
        
        // Use both hidden attribute and display style for reliability
        if (shouldShow) {
            toolbar.removeAttribute('hidden');
            toolbar.style.display = '';
        } else {
            toolbar.setAttribute('hidden', '');
            toolbar.style.display = 'none';
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

            const firstDataRow = row;
            const tbody = document.getElementById('uniqueFormulaTableBody');
            
            // Group rows by sheet name
            const rowsBySheet = new Map<string, Array<{
                ufi: string;
                fin: string;
                sheetName: string;
                formula: string;
                normalized: string;
                count: string;
                fScore: string;
                complexity: string;
                priority: string;
                status: string;
                comment: string;
                clientResponse: string;
                isFin: boolean;
            }>>();
            
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
                        } else if (rowType === 'fin') {
                            // For FIN rows, get UFI from data attribute
                            ufi = tr.getAttribute('data-ufi') || '';
                        }
                        const fin = rowType === 'fin'
                            ? cells[1].textContent || ''
                            : '';
                        const sheetName = rowType === 'ufi' ? cells[2].textContent || '' : '';
                        
                        // For FIN rows, get sheet name from parent UFI row
                        let actualSheetName = sheetName;
                        if (!actualSheetName && rowType === 'fin') {
                            const parentUfiRow = tbody.querySelector(`tr[data-row-type="ufi"][data-ufi="${ufi}"]`);
                            if (parentUfiRow) {
                                const parentCells = parentUfiRow.querySelectorAll('td');
                                if (parentCells.length >= 3) {
                                    actualSheetName = parentCells[2].textContent || '';
                                }
                            }
                        }
                        
                        const formula = rowType === 'ufi' ? cells[3].textContent || '' : '';
                        const normalized = rowType === 'ufi' ? cells[4].textContent || '' : '';
                        const count = rowType === 'ufi' ? cells[5].textContent || '' : '';
                        // Combined Complexity / F-Score cell (index 6)
                        const complexityFScoreCell = rowType === 'ufi' ? cells[6] : null;
                        let complexity = '';
                        let fScore = '';
                        if (complexityFScoreCell) {
                            const complexityLabel = complexityFScoreCell.querySelector('.complexity-label');
                            const fScoreLabel = complexityFScoreCell.querySelector('.fscore-label');
                            complexity = complexityLabel ? complexityLabel.textContent || '' : '';
                            fScore = fScoreLabel ? fScoreLabel.textContent.replace('F-Score: ', '') || '' : '';
                        }
                        // Extract priority from dropdown (both UFI and FIN rows can have priority)
                        const prioritySelect = cells[7].querySelector('select');
                        const priority = prioritySelect ? prioritySelect.value : cells[7].textContent || '';
                        // Extract status (both UFI and FIN rows can have status)
                        const status = cells[8].textContent || '';
                        const comment = cells[9].textContent || '';
                        const clientResponse = cells[10].textContent || '';
                        
                        if (actualSheetName) {
                            if (!rowsBySheet.has(actualSheetName)) {
                                rowsBySheet.set(actualSheetName, []);
                            }
                            rowsBySheet.get(actualSheetName)!.push({
                                ufi,
                                fin,
                                sheetName: actualSheetName,
                                formula,
                                normalized,
                                count,
                                fScore,
                                complexity,
                                priority,
                                status,
                                comment,
                                clientResponse,
                                isFin: rowType === 'fin'
                            });
                        }
                    }
                });
            } else {
                // Fallback: export from analysis results (no FIN rows)
                analysisResults.worksheets.forEach(ws => {
                    ws.uniqueFormulasList.forEach(info => {
                        if (!rowsBySheet.has(ws.name)) {
                            rowsBySheet.set(ws.name, []);
                        }
                        rowsBySheet.get(ws.name)!.push({
                            ufi: info.ufIndicator,
                            fin: '',
                            sheetName: ws.name,
                            formula: info.exampleFormula,
                            normalized: info.normalizedFormula,
                            count: String(info.count),
                            fScore: String(info.fScore),
                            complexity: info.complexity,
                            priority: '',
                            status: '',
                            comment: '',
                            clientResponse: '',
                            isFin: false
                        });
                    });
                });
            }
            
            // Sort sheets alphabetically and export grouped by sheet
            const sortedSheets = Array.from(rowsBySheet.keys()).sort();
            
            sortedSheets.forEach((sheetName, sheetIndex) => {
                const sectionStartRow = row;
                
                // Add separator row before new sheet section (except first)
                if (sheetIndex > 0) {
                    // Add a blank row with top border for visual separation
                    const separatorRange = sheet.getRange(`A${row}:L${row}`);
                    separatorRange.format.borders.getItem('EdgeTop').style = 'Continuous';
                    separatorRange.format.borders.getItem('EdgeTop').weight = 'Medium';
                    separatorRange.format.borders.getItem('EdgeTop').color = '#8a8886';
                    row++;
                }
                
                // Add sheet header row with merged cells
                const headerRange = sheet.getRange(`A${row}:K${row}`);
                headerRange.merge();
                headerRange.values = [[`Sheet: ${sheetName}`]];
                headerRange.format.font.bold = true;
                headerRange.format.font.size = 13;
                headerRange.format.fill.color = '#0078d4';
                headerRange.format.font.color = '#ffffff';
                headerRange.format.horizontalAlignment = 'Left';
                headerRange.format.verticalAlignment = 'Center';
                headerRange.format.borders.getItem('EdgeTop').style = 'Continuous';
                headerRange.format.borders.getItem('EdgeTop').weight = 'Thick';
                headerRange.format.borders.getItem('EdgeTop').color = '#005a9e';
                headerRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
                headerRange.format.borders.getItem('EdgeBottom').weight = 'Thick';
                headerRange.format.borders.getItem('EdgeBottom').color = '#005a9e';
                headerRange.format.borders.getItem('EdgeLeft').style = 'Continuous';
                headerRange.format.borders.getItem('EdgeLeft').weight = 'Thick';
                headerRange.format.borders.getItem('EdgeLeft').color = '#005a9e';
                headerRange.format.borders.getItem('EdgeRight').style = 'Continuous';
                headerRange.format.borders.getItem('EdgeRight').weight = 'Thick';
                headerRange.format.borders.getItem('EdgeRight').color = '#005a9e';
                row++;
                
                // Add header row for this section
                const columnHeaderRange = sheet.getRange(`A${row}:K${row}`);
                columnHeaderRange.values = TABLE_HEADER;
                columnHeaderRange.format.font.bold = true;
                columnHeaderRange.format.fill.color = '#f3f2f1';
                columnHeaderRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
                columnHeaderRange.format.borders.getItem('EdgeBottom').weight = 'Medium';
                columnHeaderRange.format.borders.getItem('EdgeBottom').color = '#8a8886';
                row++;
                
                // Export rows for this sheet (group UFI with its FIN rows)
                const sheetRows = rowsBySheet.get(sheetName)!;
                const ufiRows = sheetRows.filter(r => !r.isFin);
                const finRows = sheetRows.filter(r => r.isFin);
                
                // Group FIN rows by their UFI
                const finRowsByUfi = new Map<string, typeof finRows>();
                finRows.forEach(finRow => {
                    if (!finRowsByUfi.has(finRow.ufi)) {
                        finRowsByUfi.set(finRow.ufi, []);
                    }
                    finRowsByUfi.get(finRow.ufi)!.push(finRow);
                });
                
                // Export UFI rows with their FIN rows
                ufiRows.forEach(ufiRow => {
                    sheet.getRange(`A${row}:K${row}`).values = [[
                        ufiRow.ufi,
                        ufiRow.fin,
                        ufiRow.sheetName,
                        ufiRow.formula,
                        ufiRow.normalized,
                        ufiRow.count,
                        `${ufiRow.complexity} / ${ufiRow.fScore}`,
                        ufiRow.priority,
                        ufiRow.status,
                        ufiRow.comment,
                        ufiRow.clientResponse
                    ]];
                    row++;
                    
                    // Add FIN rows for this UFI
                    const relatedFins = finRowsByUfi.get(ufiRow.ufi) || [];
                    relatedFins.forEach(finRow => {
                        sheet.getRange(`A${row}:K${row}`).values = [[
                            finRow.ufi,
                            finRow.fin,
                            finRow.sheetName,
                            finRow.formula,
                            finRow.normalized,
                            finRow.count,
                            `${finRow.complexity} / ${finRow.fScore}`,
                            finRow.priority,
                            finRow.status,
                            finRow.comment,
                            finRow.clientResponse
                        ]];
                        row++;
                    });
                });
                
                // Add bottom border to the last row of this section for visual separation
                if (row > sectionStartRow + 2) { // Only if we have data rows
                    const lastRowRange = sheet.getRange(`A${row - 1}:K${row - 1}`);
                    lastRowRange.format.borders.getItem('EdgeBottom').style = 'Continuous';
                    lastRowRange.format.borders.getItem('EdgeBottom').weight = 'Medium';
                    lastRowRange.format.borders.getItem('EdgeBottom').color = '#8a8886';
                }
            });

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

async function applyColoring(resetOnly: boolean): Promise<void> {
    const resetFills = (document.getElementById('colorResetFills') as HTMLInputElement)?.checked ?? false;
    const resetFont = (document.getElementById('colorResetFont') as HTMLInputElement)?.checked ?? false;
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
                    await resetColors(usedRange, resetFills, resetFont);
                    continue;
                }
                if (resetFills || resetFont) {
                    await resetColors(usedRange, resetFills, resetFont);
                }
                await colorUniqueAndInputs(usedRange);
            }
            await context.sync();
        });
        showStatusMessage(resetOnly ? 'Model colors removed.' : 'Unique and input coloring applied.', 'success');
    } catch (error) {
        console.error('Coloring failed', error);
        showStatusMessage(`Coloring failed: ${getErrorMessage(error)}`, 'error');
    }
}

async function resetColors(range: Excel.Range, resetFills: boolean, resetFont: boolean): Promise<void> {
    if (resetFills) {
        range.format.fill.clear();
    }
    if (resetFont) {
        range.format.font.color = 'Automatic';
    }
}

async function colorUniqueAndInputs(range: Excel.Range): Promise<void> {
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