/**
 * Interface for analysis options
 */
export interface AnalysisOptions {
    includeEmptyCells: boolean;
    groupSimilarFormulas: boolean;
}

/**
 * Interface for formula information
 */
export interface FormulaInfo {
    formula: string;
    count: number;
    cells: string[];
}

/**
 * Interface for cell count analysis
 */
export interface CellCountAnalysis {
    totalCells: number;
    cellsWithFormulas: number;
    cellsWithValues: number;
    emptyCells: number;
    formulaPercentage: number;
    valuePercentage: number;
}

/**
 * Interface for hard-coded value detection with severity scoring
 */
export interface HardCodedValue {
    value: string;
    context: string;
    cellAddress: string;
    severity: 'High' | 'Medium' | 'Low' | 'Info';
    rationale: string;
    suggestedFix: string;
    isRepeated: boolean;
    repetitionCount: number;
    // New fields for row-wise comparison
    isInconsistent?: boolean;
    inconsistencyType?: 'value_mismatch' | 'range_endpoint' | 'pattern_deviation' | 'isolated_hardcode';
    expectedPattern?: string;
    nearbyPatterns?: string[];
}

/**
 * Interface for hard-coded value analysis
 */
export interface HardCodedValueAnalysis {
    totalHardCodedValues: number;
    highSeverityValues: HardCodedValue[];
    mediumSeverityValues: HardCodedValue[];
    lowSeverityValues: HardCodedValue[];
    infoSeverityValues: HardCodedValue[];
    repeatedValues: HardCodedValue[];
    undocumentedParameters: HardCodedValue[];
}

/**
 * Interface for worksheet analysis results
 */
export interface WorksheetAnalysisResult {
    name: string;
    totalCells: number;
    totalFormulas: number;
    uniqueFormulas: number;
    uniqueFormulasList: FormulaInfo[];
    formulaComplexity: 'Low' | 'Medium' | 'High';
    cellCountAnalysis: CellCountAnalysis;
    hardCodedValueAnalysis: HardCodedValueAnalysis;
}

/**
 * Interface for complete analysis results
 */
export interface AnalysisResult {
    totalWorksheets: number;
    totalFormulas: number;
    uniqueFormulas: number;
    worksheets: WorksheetAnalysisResult[];
    analysisTimestamp: string;
    totalCells: number;
    totalCellsWithFormulas: number;
    totalCellsWithValues: number;
    totalHardCodedValues: number;
}

/**
 * Main class for analyzing Excel formulas
 */
export class FormulaAnalyzer {
    
    /**
     * Analyze all worksheets in the workbook with performance optimizations for large models
     */
    async analyzeWorkbook(context: Excel.RequestContext, options: AnalysisOptions): Promise<AnalysisResult> {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        
        worksheets.load('items/name');
        await context.sync();
        
        const worksheetResults: WorksheetAnalysisResult[] = [];
        let totalFormulas = 0;
        let totalCells = 0;
        let totalCellsWithFormulas = 0;
        let totalCellsWithValues = 0;
        let totalHardCodedValues = 0;
        const allUniqueFormulas = new Set<string>();
        
        // Filter out add-in generated report sheets
        const reportSheetNames = new Set(["CERT_Analysis_Report", "MRT_Analysis_Report"]);
        const worksheetsToAnalyze = worksheets.items.filter(ws => !reportSheetNames.has(ws.name));

        // Process worksheets in batches to avoid memory issues
        const BATCH_SIZE = 1; // Process 1 worksheet at a time for massive models
        for (let i = 0; i < worksheetsToAnalyze.length; i += BATCH_SIZE) {
            const batch = worksheetsToAnalyze.slice(i, i + BATCH_SIZE);
            
            // Process batch in parallel but with controlled concurrency
            const batchPromises = batch.map(worksheet => this.analyzeWorksheetOptimized(context, worksheet, options));
            const batchResults = await Promise.all(batchPromises);
            
            // Accumulate results
            batchResults.forEach(worksheetResult => {
            worksheetResults.push(worksheetResult);
            totalFormulas += worksheetResult.totalFormulas;
                totalCells += worksheetResult.cellCountAnalysis.totalCells;
                totalCellsWithFormulas += worksheetResult.cellCountAnalysis.cellsWithFormulas;
                totalCellsWithValues += worksheetResult.cellCountAnalysis.cellsWithValues;
                totalHardCodedValues += worksheetResult.hardCodedValueAnalysis.totalHardCodedValues;
            
            // Add unique formulas to global set
            worksheetResult.uniqueFormulasList.forEach(formulaInfo => {
                allUniqueFormulas.add(formulaInfo.formula);
            });
            });
            
            // Yield control back to the browser to prevent UI freezing
            await new Promise(resolve => setTimeout(resolve, 0));
        }
        
        return {
            totalWorksheets: worksheetsToAnalyze.length,
            totalFormulas,
            uniqueFormulas: allUniqueFormulas.size,
            worksheets: worksheetResults,
            analysisTimestamp: new Date().toISOString(),
            totalCells,
            totalCellsWithFormulas,
            totalCellsWithValues,
            totalHardCodedValues
        };
    }
    
    /**
     * Optimized worksheet analysis for large models with chunked processing
     */
    async analyzeWorksheetOptimized(
        context: Excel.RequestContext, 
        worksheet: Excel.Worksheet, 
        options: AnalysisOptions
    ): Promise<WorksheetAnalysisResult> {
        
        worksheet.load('name');
        await context.sync();
        
        // Get the used range
        const usedRange = worksheet.getUsedRange();
        
        try {
            usedRange.load(['formulas', 'values', 'address', 'rowCount', 'columnCount']);
            await context.sync();
            
            const formulas = usedRange.formulas;
            const values = usedRange.values;
            const totalCells = usedRange.rowCount * usedRange.columnCount;
            
            // For massive models, use minimal analysis to prevent stack overflow
            const MASSIVE_MODEL_THRESHOLD = 50000; // 50k+ cells = massive model
            if (totalCells > MASSIVE_MODEL_THRESHOLD) {
                return await this.analyzeMassiveWorksheet(context, worksheet, formulas, values, totalCells, options);
            }
            
        // For very large worksheets, process in chunks
        const CHUNK_SIZE = 100; // Process 100 rows at a time for massive models
        const totalRows = formulas.length;
        
        if (totalRows > CHUNK_SIZE) {
            return await this.analyzeLargeWorksheet(context, worksheet, formulas, values, totalCells, options);
        }
            
            // For smaller worksheets, use the original method
            return await this.analyzeWorksheet(context, worksheet, options);
            
        } catch (error) {
            console.warn(`Warning: Could not analyze worksheet "${worksheet.name}":`, error);
            
            return {
                name: worksheet.name,
                totalCells: 0,
                totalFormulas: 0,
                uniqueFormulas: 0,
                uniqueFormulasList: [],
                formulaComplexity: 'Low',
                cellCountAnalysis: {
                    totalCells: 0,
                    cellsWithFormulas: 0,
                    cellsWithValues: 0,
                    emptyCells: 0,
                    formulaPercentage: 0,
                    valuePercentage: 0
                },
                hardCodedValueAnalysis: {
                    totalHardCodedValues: 0,
                    highSeverityValues: [],
                    mediumSeverityValues: [],
                    lowSeverityValues: [],
                    infoSeverityValues: [],
                    repeatedValues: [],
                    undocumentedParameters: []
                }
            };
        }
    }
    
    /**
     * Analyze a single worksheet
     */
    async analyzeWorksheet(
        context: Excel.RequestContext, 
        worksheet: Excel.Worksheet, 
        options: AnalysisOptions
    ): Promise<WorksheetAnalysisResult> {
        
        worksheet.load('name');
        await context.sync();
        
        // Get the used range
        const usedRange = worksheet.getUsedRange();
        
        try {
            usedRange.load(['formulas', 'values', 'address', 'rowCount', 'columnCount']);
            await context.sync();
            
            const formulas = usedRange.formulas;
            const values = usedRange.values;
            const totalCells = usedRange.rowCount * usedRange.columnCount;
            
            // Analyze formulas
            const formulaMap = new Map<string, FormulaInfo>();
            let totalFormulas = 0;
            
            // Cell count analysis
            const cellCountAnalysis = this.analyzeCellCounts(formulas, values, totalCells);
            
            // Hard-coded value analysis
            const hardCodedValueAnalysis = this.analyzeHardCodedValues(formulas, values);
            
            for (let row = 0; row < formulas.length; row++) {
                for (let col = 0; col < formulas[row].length; col++) {
                    const formula = formulas[row][col] as string;
                    const value = values[row][col];
                    
                    // Skip if not a formula or if it's an empty cell and we're not including them
                    if (!this.isFormula(formula) || (!options.includeEmptyCells && this.isEmpty(value))) {
                        continue;
                    }
                    
                    totalFormulas++;
                    
                    let normalizedFormula = formula;
                    if (options.groupSimilarFormulas) {
                        normalizedFormula = this.normalizeFormula(formula);
                    }
                    
                    const cellAddress = this.getCellAddress(row, col);
                    
                    if (formulaMap.has(normalizedFormula)) {
                        const existing = formulaMap.get(normalizedFormula)!;
                        existing.count++;
                        existing.cells.push(cellAddress);
                    } else {
                        formulaMap.set(normalizedFormula, {
                            formula: normalizedFormula,
                            count: 1,
                            cells: [cellAddress]
                        });
                    }
                }
            }
            
            const uniqueFormulasList = Array.from(formulaMap.values())
                .sort((a, b) => b.count - a.count); // Sort by frequency
            
            return {
                name: worksheet.name,
                totalCells,
                totalFormulas,
                uniqueFormulas: uniqueFormulasList.length,
                uniqueFormulasList,
                formulaComplexity: this.assessComplexity(uniqueFormulasList),
                cellCountAnalysis,
                hardCodedValueAnalysis
            };
            
        } catch (error) {
            // Handle case where worksheet is empty or has no used range
            console.warn(`Warning: Could not analyze worksheet "${worksheet.name}":`, error);
            
            return {
                name: worksheet.name,
                totalCells: 0,
                totalFormulas: 0,
                uniqueFormulas: 0,
                uniqueFormulasList: [],
                formulaComplexity: 'Low',
                cellCountAnalysis: {
                    totalCells: 0,
                    cellsWithFormulas: 0,
                    cellsWithValues: 0,
                    emptyCells: 0,
                    formulaPercentage: 0,
                    valuePercentage: 0
                },
                hardCodedValueAnalysis: {
                    totalHardCodedValues: 0,
                    highSeverityValues: [],
                    mediumSeverityValues: [],
                    lowSeverityValues: [],
                    infoSeverityValues: [],
                    repeatedValues: [],
                    undocumentedParameters: []
                }
            };
        }
    }
    
    /**
     * Check if a cell value is a formula
     */
    private isFormula(value: any): boolean {
        return typeof value === 'string' && value.startsWith('=');
    }
    
    /**
     * Check if a cell value is empty
     */
    private isEmpty(value: any): boolean {
        return value === null || value === undefined || value === '';
    }
    
    /**
     * Normalize formula for grouping similar formulas
     * This replaces cell references with placeholders to group structurally similar formulas
     */
    private normalizeFormula(formula: string): string {
        if (!this.isFormula(formula)) {
            return formula;
        }
        
        // Replace cell references (like A1, B2, $A$1) with placeholders
        let normalized = formula;
        
        // Absolute references ($A$1, $B$2, etc.)
        normalized = normalized.replace(/\$[A-Z]+\$\d+/g, '<ABS_REF>');
        
        // Mixed references ($A1, A$1, etc.)
        normalized = normalized.replace(/\$?[A-Z]+\$?\d+/g, '<REF>');
        
        // Range references (A1:B2, etc.)
        normalized = normalized.replace(/<REF>:<REF>/g, '<RANGE>');
        normalized = normalized.replace(/<ABS_REF>:<ABS_REF>/g, '<ABS_RANGE>');
        
        // Named ranges and table references might be preserved
        // This is a basic implementation - could be enhanced further
        
        return normalized;
    }
    
    /**
     * Convert row/column indices to Excel cell address
     */
    private getCellAddress(row: number, col: number): string {
        const columnLetter = this.numberToColumnLetter(col + 1);
        return `${columnLetter}${row + 1}`;
    }
    
    /**
     * Convert column number to Excel column letter(s)
     */
    private numberToColumnLetter(num: number): string {
        let result = '';
        while (num > 0) {
            num--;
            result = String.fromCharCode(65 + (num % 26)) + result;
            num = Math.floor(num / 26);
        }
        return result;
    }
    
    /**
     * Assess the complexity of formulas in a worksheet
     */
    private assessComplexity(formulas: FormulaInfo[]): 'Low' | 'Medium' | 'High' {
        if (formulas.length === 0) return 'Low';
        
        const totalFormulas = formulas.reduce((sum, f) => sum + f.count, 0);
        const uniqueFormulas = formulas.length;
        const complexFormulas = formulas.filter(f => this.isComplexFormula(f.formula)).length;
        
        // Calculate complexity score
        let score = 0;
        
        // Factor 1: Total number of formulas
        if (totalFormulas > 100) score += 2;
        else if (totalFormulas > 20) score += 1;
        
        // Factor 2: Number of unique formulas (high uniqueness can indicate complexity)
        const uniquenessRatio = uniqueFormulas / totalFormulas;
        if (uniquenessRatio > 0.7) score += 2;
        else if (uniquenessRatio > 0.4) score += 1;
        
        // Factor 3: Presence of complex formulas
        const complexityRatio = complexFormulas / uniqueFormulas;
        if (complexityRatio > 0.3) score += 2;
        else if (complexityRatio > 0.1) score += 1;
        
        // Determine final complexity
        if (score >= 4) return 'High';
        if (score >= 2) return 'Medium';
        return 'Low';
    }
    
    /**
     * Check if a formula is considered complex
     */
    private isComplexFormula(formula: string): boolean {
        const complexPatterns = [
            /INDEX\s*\(/i,           // INDEX function
            /MATCH\s*\(/i,           // MATCH function
            /VLOOKUP\s*\(/i,         // VLOOKUP function
            /HLOOKUP\s*\(/i,         // HLOOKUP function
            /XLOOKUP\s*\(/i,         // XLOOKUP function
            /SUMIFS\s*\(/i,          // SUMIFS function
            /COUNTIFS\s*\(/i,        // COUNTIFS function
            /AVERAGEIFS\s*\(/i,      // AVERAGEIFS function
            /INDIRECT\s*\(/i,        // INDIRECT function
            /OFFSET\s*\(/i,          // OFFSET function
            /ARRAY\s*\(/i,           // Array formulas
            /\{[^}]*\}/,              // Array constants (non-greedy, bounded)
        ];
        
        if (complexPatterns.some(pattern => pattern.test(formula))) {
            return true;
        }
        
        // Simple and safe nested function heuristic: count parentheses
        const openParens = (formula.match(/\(/g) || []).length;
        const closeParens = (formula.match(/\)/g) || []).length;
        const nestedDepthEstimate = Math.min(openParens, closeParens);
        return nestedDepthEstimate >= 3;
    }
    
    /**
     * Analyze cell counts in a worksheet
     */
    private analyzeCellCounts(formulas: any[][], values: any[][], totalCells: number): CellCountAnalysis {
        let cellsWithFormulas = 0;
        let cellsWithValues = 0;
        let emptyCells = 0;
        
        for (let row = 0; row < formulas.length; row++) {
            for (let col = 0; col < formulas[row].length; col++) {
                const formula = formulas[row][col] as string;
                const value = values[row][col];
                
                if (this.isFormula(formula)) {
                    cellsWithFormulas++;
                } else if (!this.isEmpty(value)) {
                    cellsWithValues++;
                } else {
                    emptyCells++;
                }
            }
        }
        
        const formulaPercentage = totalCells > 0 ? (cellsWithFormulas / totalCells) * 100 : 0;
        const valuePercentage = totalCells > 0 ? (cellsWithValues / totalCells) * 100 : 0;
        
        return {
            totalCells,
            cellsWithFormulas,
            cellsWithValues,
            emptyCells,
            formulaPercentage: Math.round(formulaPercentage * 100) / 100,
            valuePercentage: Math.round(valuePercentage * 100) / 100
        };
    }
    
    /**
     * Analyze hard-coded values in formulas with row-wise pattern detection
     * This simulates how a top-level model reviewer would identify issues
     */
    private analyzeHardCodedValues(formulas: any[][], values: any[][]): HardCodedValueAnalysis {
        const hardCodedValues: HardCodedValue[] = [];
        
        // Step 1: Build a map of formulas by column for pattern analysis
        const columnFormulas = new Map<number, Array<{row: number, formula: string, normalized: string}>>();
        
        for (let row = 0; row < formulas.length; row++) {
            for (let col = 0; col < formulas[row].length; col++) {
                const formula = formulas[row][col] as string;
                
                if (this.isFormula(formula)) {
                    if (!columnFormulas.has(col)) {
                        columnFormulas.set(col, []);
                    }
                    
                    const normalized = this.normalizeFormulaForComparison(formula);
                    columnFormulas.get(col)!.push({ row, formula, normalized });
                }
            }
        }
        
        // Step 2: Analyze each formula with awareness of its column context
        for (let row = 0; row < formulas.length; row++) {
            for (let col = 0; col < formulas[row].length; col++) {
                const formula = formulas[row][col] as string;
                
                if (this.isFormula(formula)) {
                    const cellAddress = this.getCellAddress(row, col);
                    const columnContext = columnFormulas.get(col) || [];
                    
                    // Detect hard-coded values with row-wise context
                    const detectedValues = this.detectHardCodedValuesWithContext(
                        formula, 
                        cellAddress, 
                        row, 
                        col,
                        columnContext,
                        formulas
                    );
                    
                    hardCodedValues.push(...detectedValues);
                }
            }
        }
        
        // Step 3: Post-process to identify patterns and adjust severity
        this.analyzePatternConsistency(hardCodedValues, formulas);
        
        // Categorize by severity (not confidence anymore)
        const highSeverity = hardCodedValues.filter(v => v.severity === 'High');
        const mediumSeverity = hardCodedValues.filter(v => v.severity === 'Medium');
        const lowSeverity = hardCodedValues.filter(v => v.severity === 'Low');
        const infoSeverity = hardCodedValues.filter(v => v.severity === 'Info');
        
        // Find repeated values
        const valueMap = new Map<string, HardCodedValue[]>();
        hardCodedValues.forEach(hcv => {
            const key = hcv.value;
            if (!valueMap.has(key)) {
                valueMap.set(key, []);
            }
            valueMap.get(key)!.push(hcv);
        });
        
        valueMap.forEach((values, key) => {
            if (values.length > 1) {
                values.forEach(v => {
                    v.isRepeated = true;
                    v.repetitionCount = values.length;
                });
            }
        });
        
        const repeatedValues = hardCodedValues.filter(v => v.isRepeated);
        const undocumentedParameters = hardCodedValues.filter(v => 
            (v.severity === 'High' || v.severity === 'Medium') && v.isRepeated
        );
        
        return {
            totalHardCodedValues: hardCodedValues.length,
            highSeverityValues: highSeverity,
            mediumSeverityValues: mediumSeverity,
            lowSeverityValues: lowSeverity,
            infoSeverityValues: infoSeverity,
            repeatedValues: repeatedValues,
            undocumentedParameters: undocumentedParameters
        };
    }
    
    /**
     * Normalize formula for comparison by removing row-specific references
     * This helps identify patterns across rows
     */
    private normalizeFormulaForComparison(formula: string): string {
        if (!formula) return '';
        
        // Replace row numbers with placeholder to identify column patterns
        let normalized = formula;
        
        // Replace cell references like A1, B2 with A#, B#
        normalized = normalized.replace(/([A-Z]+)(\d+)/g, '$1#');
        
        // Replace absolute references like $A$1 with $A$#
        normalized = normalized.replace(/\$([A-Z]+)\$(\d+)/g, '$$1$#');
        
        // Replace mixed references
        normalized = normalized.replace(/\$([A-Z]+)(\d+)/g, '$$1#');
        normalized = normalized.replace(/([A-Z]+)\$(\d+)/g, '$1$#');
        
        return normalized;
    }
    
    /**
     * Detect hard-coded values with row-wise context awareness
     */
    private detectHardCodedValuesWithContext(
        formula: string, 
        cellAddress: string,
        row: number,
        col: number,
        columnContext: Array<{row: number, formula: string, normalized: string}>,
        allFormulas: any[][]
    ): HardCodedValue[] {
        const hardCodedValues: HardCodedValue[] = [];
        
        // Get nearby rows for comparison (±3 rows)
        const nearbyRows = columnContext.filter(ctx => 
            Math.abs(ctx.row - row) <= 3 && ctx.row !== row
        );
        
        // Define patterns for different types of literals
        const literalPatterns = [
            // Fixed range endpoints (highest priority)
            { pattern: /([A-Z]+\$\d+)/g, type: 'fixed_row_ref' },
            { pattern: /(\$[A-Z]+\d+)/g, type: 'fixed_col_ref' },
            { pattern: /(\$[A-Z]+\$\d+)/g, type: 'fixed_absolute' },
            // Numeric literals
            { pattern: /(?<![A-Z])(\d+\.\d+)(?![0-9])/g, type: 'decimal' },
            { pattern: /(?<![A-Z0-9])(\d{2,})(?![0-9])/g, type: 'multi_digit' },
            { pattern: /(?<![A-Z0-9])([2-9])(?![0-9])/g, type: 'single_digit' },
            // String literals
            { pattern: /"([^"]+)"/g, type: 'string' },
            // Array literals
            { pattern: /\{([^}]+)\}/g, type: 'array' }
        ];
        
        literalPatterns.forEach(({ pattern, type }) => {
            let match;
            pattern.lastIndex = 0; // Reset regex state
            
            while ((match = pattern.exec(formula)) !== null) {
                const value = match[0];
                const position = match.index;
                
                // Check if this value appears inconsistently in nearby rows
                const inconsistencyAnalysis = this.checkRowConsistency(
                    value,
                    formula,
                    nearbyRows,
                    type,
                    row,
                    col,
                    allFormulas
                );
                
                // Determine severity based on inconsistency and type
                const severity = this.calculateSeverityFromInconsistency(
                    value,
                    type,
                    inconsistencyAnalysis,
                    formula,
                    position
                );
                
                // Only report if significant
                if (severity !== 'Info' || inconsistencyAnalysis.isInconsistent) {
                    hardCodedValues.push({
                        value: value,
                        context: this.extractFormulaContext(formula, position, 50),
                        cellAddress: cellAddress,
                        severity: severity,
                        rationale: inconsistencyAnalysis.rationale || this.getRationaleForType(type, value, inconsistencyAnalysis),
                        suggestedFix: this.getSuggestedFix(type, value, inconsistencyAnalysis),
                        isRepeated: false,
                        repetitionCount: 0,
                        isInconsistent: inconsistencyAnalysis.isInconsistent,
                        inconsistencyType: inconsistencyAnalysis.type,
                        expectedPattern: inconsistencyAnalysis.expectedPattern,
                        nearbyPatterns: inconsistencyAnalysis.nearbyPatterns
                    });
                }
            }
        });
        
        return hardCodedValues;
    }
    
    /**
     * Detect hard-coded values in a formula with context-aware severity scoring
     */
    private detectHardCodedValues(formula: string, cellAddress: string): HardCodedValue[] {
        const hardCodedValues: HardCodedValue[] = [];
        
        // Define patterns for different types of literals
        const literalPatterns = [
            // Numeric literals
            { pattern: /\b(\d+\.\d+)\b/g, type: 'decimal' },
            { pattern: /\b(\d+)\b/g, type: 'integer' },
            // String literals
            { pattern: /"([^"]+)"/g, type: 'string' },
            // Date literals
            { pattern: /\b(\d{1,2}\/\d{1,2}\/\d{4})\b/g, type: 'date' },
            // Array literals
            { pattern: /\{([^}]+)\}/g, type: 'array' },
            // Fixed range endpoints
            { pattern: /:\$([A-Z]+\$\d+)\b/g, type: 'fixed_range' }
        ];
        
        literalPatterns.forEach(({ pattern, type }) => {
            let match;
            while ((match = pattern.exec(formula)) !== null) {
                const value = match[0];
                const position = match.index;
                
                // Analyze context and determine severity
                const analysis = this.analyzeLiteralContext(value, formula, position, type);
                
                if (analysis.severity !== 'Info' || analysis.isSignificant) {
                    hardCodedValues.push({
                        value: value,
                        context: this.extractFormulaContext(formula, position, 50),
                        cellAddress: cellAddress,
                        severity: analysis.severity,
                        rationale: analysis.rationale,
                        suggestedFix: analysis.suggestedFix,
                        isRepeated: false, // Will be updated later
                        repetitionCount: 0
                    });
                }
            }
        });
        
        return hardCodedValues;
    }
    
    /**
     * Determine if a detected value is likely to be hard-coded
     */
    private isLikelyHardCoded(value: string, formula: string, position: number): boolean {
        // Skip if it's part of a cell reference (like A1, B2)
        if (/^[A-Z]+\d+$/.test(value)) {
            return false;
        }
        
        // Skip if it's part of a function name
        const functionNames = ['INDEX', 'MATCH', 'VLOOKUP', 'HLOOKUP', 'SUM', 'AVERAGE', 'COUNT', 'IF', 'AND', 'OR'];
        const beforeValue = formula.substring(Math.max(0, position - 10), position);
        if (functionNames.some(func => beforeValue.toUpperCase().includes(func))) {
            return false;
        }
        
        // Skip if it's clearly a column number in array functions
        if (/\b(COLUMN|ROW|INDEX|MATCH)\s*\([^)]*,\s*$/.test(formula.substring(0, position + value.length))) {
            return false;
        }
        
        // Skip if it's a date component
        if (/\b(19|20)\d{2}\b/.test(value) && value.length === 4) {
            return false;
        }
        
        // Skip very small numbers that are likely function parameters
        const numValue = parseFloat(value);
        if (numValue >= 0 && numValue <= 10) {
            // Additional context check for small numbers
            const context = formula.substring(Math.max(0, position - 20), position + value.length + 20);
            if (/\b(COLUMN|ROW|INDEX|MATCH|VLOOKUP|HLOOKUP)\b/i.test(context)) {
                return false;
            }
        }
        
        return true;
    }
    
    /**
     * Analyze the context of a literal to determine severity and rationale
     */
    private analyzeLiteralContext(value: string, formula: string, position: number, type: string): {
        severity: 'High' | 'Medium' | 'Low' | 'Info';
        rationale: string;
        suggestedFix: string;
        isSignificant: boolean;
    } {
        const context = this.getLiteralContext(formula, position);
        const numValue = parseFloat(value);
        
        // Check if it's a benign function parameter
        if (this.isBenignFunctionParameter(value, context)) {
            return {
                severity: 'Info',
                rationale: 'Benign function parameter',
                suggestedFix: 'No action needed',
                isSignificant: false
            };
        }
        
        // High severity: Financial rates, thresholds, offsets
        if (this.isFinancialRateOrThreshold(value, context)) {
            return {
                severity: 'High',
                rationale: 'Financial rate or threshold driving business logic',
                suggestedFix: 'Move to parameter cell or named range',
                isSignificant: true
            };
        }
        
        // High severity: Fixed range endpoints
        if (type === 'fixed_range') {
            return {
                severity: 'High',
                rationale: 'Fixed range endpoint that may need to be dynamic',
                suggestedFix: 'Consider using dynamic range or named range',
                isSignificant: true
            };
        }
        
        // High severity: Large offsets or scaling factors
        if (this.isLargeOffsetOrScaling(value, context)) {
            return {
                severity: 'High',
                rationale: 'Large offset or scaling factor affecting calculations',
                suggestedFix: 'Parameterize and document',
                isSignificant: true
            };
        }
        
        // Medium severity: Small step constants, unit scalers
        if (this.isStepConstantOrScaler(value, context)) {
            return {
                severity: 'Medium',
                rationale: 'Step constant or unit scaler',
                suggestedFix: 'Consider parameterizing if used frequently',
                isSignificant: true
            };
        }
        
        // Medium severity: Array literals
        if (type === 'array') {
            return {
                severity: 'Medium',
                rationale: 'Inline array that could be externalized',
                suggestedFix: 'Move to separate range or named array',
                isSignificant: true
            };
        }
        
        // Low severity: Small numbers that might be significant
        if (numValue >= 2 && numValue <= 20) {
            return {
                severity: 'Low',
                rationale: 'Small numeric constant',
                suggestedFix: 'Review if this should be parameterized',
                isSignificant: true
            };
        }
        
        // Info: Very small numbers or common values
        return {
            severity: 'Info',
            rationale: 'Common numeric value',
            suggestedFix: 'No action needed unless policy requires',
            isSignificant: false
        };
    }
    
    /**
     * Check if a value is a benign function parameter
     */
    private isBenignFunctionParameter(value: string, context: string): boolean {
        const benignPatterns = [
            // ROUND function digits
            /ROUND\s*\([^,]+,\s*$/,
            // MATCH function mode
            /MATCH\s*\([^,]+,\s*[^,]+,\s*$/,
            // LEFT/RIGHT function width
            /(LEFT|RIGHT)\s*\([^,]+,\s*$/,
            // EOMONTH function offset
            /EOMONTH\s*\([^,]+,\s*$/,
            // VLOOKUP function column index
            /VLOOKUP\s*\([^,]+,\s*[^,]+,\s*$/,
            // HLOOKUP function row index
            /HLOOKUP\s*\([^,]+,\s*[^,]+,\s*$/,
            // INDEX function row/column
            /INDEX\s*\([^,]+,\s*[^,]+,\s*$/,
            // Array function dimensions
            /(ROWS|COLUMNS)\s*\([^)]*,\s*$/
        ];
        
        return benignPatterns.some(pattern => pattern.test(context));
    }
    
    /**
     * Check if a value is a financial rate or threshold
     */
    private isFinancialRateOrThreshold(value: string, context: string): boolean {
        const numValue = parseFloat(value);
        
        // Financial rates (percentages)
        if (numValue > 0 && numValue < 1 && value.includes('.')) {
            return true;
        }
        
        // Common financial thresholds
        const financialThresholds = [0.1, 0.15, 0.2, 0.25, 0.3, 0.5, 1, 2, 5, 10, 12, 24, 36, 60, 120, 240, 360];
        if (financialThresholds.includes(numValue)) {
            return true;
        }
        
        // Context clues for financial values
        const financialKeywords = ['rate', 'discount', 'interest', 'threshold', 'limit', 'cap', 'floor', 'multiplier'];
        return financialKeywords.some(keyword => context.toLowerCase().includes(keyword));
    }
    
    /**
     * Check if a value is a large offset or scaling factor
     */
    private isLargeOffsetOrScaling(value: string, context: string): boolean {
        const numValue = parseFloat(value);
        
        // Large offsets (like -15 in your example)
        if (numValue >= 10 && (context.includes('-') || context.includes('+'))) {
            return true;
        }
        
        // Scaling factors
        const scalingFactors = [12, 24, 52, 365, 1000, 10000];
        if (scalingFactors.includes(numValue)) {
            return true;
        }
        
        return false;
    }
    
    /**
     * Check if a value is a step constant or unit scaler
     */
    private isStepConstantOrScaler(value: string, context: string): boolean {
        const numValue = parseFloat(value);
        
        // Step constants
        if (numValue === -1 || numValue === 1) {
            return true;
        }
        
        // Unit scalers
        const unitScalers = [12, 24, 52, 365];
        if (unitScalers.includes(numValue) && (context.includes('/') || context.includes('*'))) {
            return true;
        }
        
        return false;
    }
    
    /**
     * Get the context around a literal in the formula
     */
    private getLiteralContext(formula: string, position: number): string {
        const start = Math.max(0, position - 30);
        const end = Math.min(formula.length, position + 30);
        return formula.substring(start, end);
    }
    
    /**
     * Extract a clean context snippet around a position
     */
    private extractFormulaContext(formula: string, position: number, maxLength: number): string {
        const start = Math.max(0, position - maxLength / 2);
        const end = Math.min(formula.length, position + maxLength / 2);
        let context = formula.substring(start, end);
        
        // Add ellipsis if truncated
        if (start > 0) context = '...' + context;
        if (end < formula.length) context = context + '...';
        
        return context;
    }
    
    /**
     * Check if a hard-coded value is consistent with nearby rows
     * This is the core of the row-wise comparison logic
     */
    private checkRowConsistency(
        value: string,
        formula: string,
        nearbyRows: Array<{row: number, formula: string, normalized: string}>,
        type: string,
        currentRow: number,
        currentCol: number,
        allFormulas: any[][]
    ): {
        isInconsistent: boolean;
        type?: 'value_mismatch' | 'range_endpoint' | 'pattern_deviation' | 'isolated_hardcode';
        rationale?: string;
        expectedPattern?: string;
        nearbyPatterns?: string[];
    } {
        // If no nearby rows, can't determine inconsistency
        if (nearbyRows.length === 0) {
            return { isInconsistent: false };
        }
        
        // For fixed range endpoints, check if they appear in multiple rows
        if (type === 'fixed_row_ref' || type === 'fixed_col_ref' || type === 'fixed_absolute') {
            const fixedRefCount = nearbyRows.filter(r => r.formula.includes(value)).length;
            if (fixedRefCount > nearbyRows.length * 0.5) {
                return {
                    isInconsistent: true,
                    type: 'range_endpoint',
                    rationale: `Fixed range endpoint "${value}" appears in ${fixedRefCount + 1} consecutive rows - should likely be dynamic`,
                    expectedPattern: 'Dynamic range that grows with data',
                    nearbyPatterns: nearbyRows.slice(0, 3).map(r => r.formula.substring(0, 50))
                };
            }
        }
        
        // For numeric values, check if they differ from similar formulas
        if (type === 'decimal' || type === 'multi_digit' || type === 'single_digit') {
            const numValue = parseFloat(value);
            
            // Skip common benign values
            if (this.isBenignValue(value, formula)) {
                return { isInconsistent: false };
            }
            
            // Check if this numeric value appears at the same position in nearby formulas
            const normalizedCurrent = this.normalizeFormulaForComparison(formula);
            const similarFormulas = nearbyRows.filter(r => {
                const similarity = this.calculateFormulaSimilarity(normalizedCurrent, r.normalized);
                return similarity > 0.7; // 70% similar structure
            });
            
            if (similarFormulas.length > 0) {
                // Extract numeric values from similar formulas at similar positions
                const otherValues = new Set<string>();
                similarFormulas.forEach(sf => {
                    const pattern = type === 'decimal' ? /\d+\.\d+/g : 
                                   type === 'multi_digit' ? /\d{2,}/g : /[2-9]/g;
                    let match;
                    while ((match = pattern.exec(sf.formula)) !== null) {
                        otherValues.add(match[0]);
                    }
                });
                
                // If this value is unique among similar formulas, it's inconsistent
                if (!otherValues.has(value) && otherValues.size > 0) {
                    const otherValuesList = Array.from(otherValues);
                    return {
                        isInconsistent: true,
                        type: 'value_mismatch',
                        rationale: `Value "${value}" differs from nearby rows which use: ${otherValuesList.join(', ')}`,
                        expectedPattern: `Consistent with nearby values: ${otherValuesList[0]}`,
                        nearbyPatterns: similarFormulas.slice(0, 3).map(r => r.formula.substring(0, 50))
                    };
                }
                
                // Check for pattern deviation (e.g., all others reference cells, this one is hard-coded)
                const hasReferencePattern = similarFormulas.some(sf => {
                    const position = sf.formula.indexOf(value);
                    if (position === -1) return false;
                    const before = sf.formula.substring(Math.max(0, position - 10), position);
                    return /[A-Z]+\d+/.test(before);
                });
                
                if (hasReferencePattern) {
                    return {
                        isInconsistent: true,
                        type: 'pattern_deviation',
                        rationale: `Hard-coded value "${value}" where nearby rows use cell references`,
                        expectedPattern: 'Cell reference instead of hard-coded value',
                        nearbyPatterns: similarFormulas.slice(0, 3).map(r => r.formula.substring(0, 50))
                    };
                }
            }
        }
        
        return { isInconsistent: false };
    }
    
    /**
     * Calculate similarity between two normalized formulas
     */
    private calculateFormulaSimilarity(formula1: string, formula2: string): number {
        if (!formula1 || !formula2) return 0;
        
        // Simple Levenshtein distance-based similarity
        const maxLen = Math.max(formula1.length, formula2.length);
        if (maxLen === 0) return 1;
        
        const distance = this.levenshteinDistance(formula1, formula2);
        return 1 - (distance / maxLen);
    }
    
    /**
     * Calculate Levenshtein distance between two strings
     */
    private levenshteinDistance(str1: string, str2: string): number {
        const matrix: number[][] = [];
        
        for (let i = 0; i <= str2.length; i++) {
            matrix[i] = [i];
        }
        
        for (let j = 0; j <= str1.length; j++) {
            matrix[0][j] = j;
        }
        
        for (let i = 1; i <= str2.length; i++) {
            for (let j = 1; j <= str1.length; j++) {
                if (str2.charAt(i - 1) === str1.charAt(j - 1)) {
                    matrix[i][j] = matrix[i - 1][j - 1];
                } else {
                    matrix[i][j] = Math.min(
                        matrix[i - 1][j - 1] + 1, // substitution
                        matrix[i][j - 1] + 1,     // insertion
                        matrix[i - 1][j] + 1      // deletion
                    );
                }
            }
        }
        
        return matrix[str2.length][str1.length];
    }
    
    /**
     * Check if a value is benign and shouldn't be flagged
     */
    private isBenignValue(value: string, formula: string): boolean {
        const numValue = parseFloat(value);
        
        // Common benign values
        const benignValues = [0, 1, 2, 12, 100, 1000];
        if (benignValues.includes(numValue)) {
            // Check context to see if it's actually benign
            const contextClues = [
                /ROUND\s*\([^,]+,\s*\d+\s*\)/i,  // ROUND function parameter
                /ROW\s*\(\s*\)/i,                 // ROW function
                /COLUMN\s*\(\s*\)/i,               // COLUMN function
                /MONTH|DAY|YEAR/i,                // Date functions
                /\*\s*12(?:\s|,|\))/,             // Months in year
                /\/\s*100(?:\s|,|\))/,            // Percentage conversion
                /\*\s*1000(?:\s|,|\))/            // Unit conversion
            ];
            
            return contextClues.some(pattern => pattern.test(formula));
        }
        
        // Index/Match column numbers (1-20 are often column indices)
        if (numValue >= 1 && numValue <= 20) {
            const indexMatchPattern = /(INDEX|MATCH|VLOOKUP|HLOOKUP)\s*\(/i;
            if (indexMatchPattern.test(formula)) {
                return true;
            }
        }
        
        return false;
    }
    
    /**
     * Calculate severity based on inconsistency analysis
     */
    private calculateSeverityFromInconsistency(
        value: string,
        type: string,
        inconsistencyAnalysis: any,
        formula: string,
        position: number
    ): 'High' | 'Medium' | 'Low' | 'Info' {
        // If inconsistent with nearby rows, always at least Medium
        if (inconsistencyAnalysis.isInconsistent) {
            if (inconsistencyAnalysis.type === 'range_endpoint') {
                return 'High'; // Fixed range endpoints are critical
            }
            if (inconsistencyAnalysis.type === 'value_mismatch') {
                // Large value mismatches are High severity
                const numValue = parseFloat(value);
                if (!isNaN(numValue) && Math.abs(numValue) > 100) {
                    return 'High';
                }
                return 'Medium';
            }
            if (inconsistencyAnalysis.type === 'pattern_deviation') {
                return 'Medium';
            }
        }
        
        // For non-inconsistent values, use type-based heuristics
        if (type === 'fixed_row_ref' || type === 'fixed_col_ref' || type === 'fixed_absolute') {
            return 'Medium'; // Fixed references are usually problematic
        }
        
        // Check for financial/business logic indicators
        const numValue = parseFloat(value);
        if (!isNaN(numValue)) {
            // Percentages and rates
            if (numValue > 0 && numValue < 1 && value.includes('.')) {
                return 'High';
            }
            // Large numbers that might be thresholds
            if (numValue > 1000) {
                return 'Medium';
            }
            // Common multipliers
            if ([365, 52, 24, 12].includes(numValue)) {
                return 'Low';
            }
        }
        
        // String literals
        if (type === 'string' && value.length > 3) {
            return 'Medium';
        }
        
        // Default to Info for small values
        return 'Info';
    }
    
    /**
     * Get rationale for a specific type of hard-coded value
     */
    private getRationaleForType(type: string, value: string, inconsistencyAnalysis: any): string {
        if (inconsistencyAnalysis.rationale) {
            return inconsistencyAnalysis.rationale;
        }
        
        switch (type) {
            case 'fixed_row_ref':
            case 'fixed_col_ref':
            case 'fixed_absolute':
                return `Fixed reference "${value}" may not adjust when data grows`;
            case 'decimal':
                return `Decimal value "${value}" may be a rate or percentage that should be parameterized`;
            case 'multi_digit':
                return `Numeric constant "${value}" should potentially be in a parameter cell`;
            case 'single_digit':
                return `Small constant "${value}" - verify if this should be parameterized`;
            case 'string':
                return `String literal "${value}" might need to be externalized`;
            case 'array':
                return `Inline array "${value}" could be moved to a range`;
            default:
                return `Hard-coded value "${value}" detected`;
        }
    }
    
    /**
     * Get suggested fix for a hard-coded value
     */
    private getSuggestedFix(type: string, value: string, inconsistencyAnalysis: any): string {
        if (inconsistencyAnalysis.expectedPattern) {
            return `Change to: ${inconsistencyAnalysis.expectedPattern}`;
        }
        
        switch (type) {
            case 'fixed_row_ref':
            case 'fixed_col_ref':
            case 'fixed_absolute':
                return 'Use dynamic ranges or remove absolute references where appropriate';
            case 'decimal':
            case 'multi_digit':
                return 'Move to a clearly labeled input/parameter cell and reference it';
            case 'string':
                return 'Consider using a lookup table or named range';
            case 'array':
                return 'Move array to a separate range and reference it';
            default:
                return 'Review and parameterize if used across multiple cells';
        }
    }
    
    /**
     * Post-process to analyze pattern consistency across all detected values
     */
    private analyzePatternConsistency(hardCodedValues: HardCodedValue[], formulas: any[][]): void {
        // Group values by column
        const columnGroups = new Map<number, HardCodedValue[]>();
        
        hardCodedValues.forEach(hcv => {
            // Extract column from cell address (e.g., "B5" -> column 1)
            const match = hcv.cellAddress.match(/([A-Z]+)(\d+)/);
            if (match) {
                const colLetters = match[1];
                const col = this.columnLetterToNumber(colLetters) - 1;
                
                if (!columnGroups.has(col)) {
                    columnGroups.set(col, []);
                }
                columnGroups.get(col)!.push(hcv);
            }
        });
        
        // Analyze each column group for patterns
        columnGroups.forEach((values, col) => {
            if (values.length > 2) {
                // Check if all values in the column are the same
                const uniqueValues = new Set(values.map(v => v.value));
                
                if (uniqueValues.size === 1) {
                    // All same value - likely a parameter that should be externalized
                    values.forEach(v => {
                        v.severity = 'High';
                        v.rationale = `Value "${v.value}" repeated ${values.length} times in column - should be parameterized`;
                        v.isRepeated = true;
                        v.repetitionCount = values.length;
                    });
                } else if (uniqueValues.size < values.length * 0.5) {
                    // Few unique values relative to occurrences
                    values.forEach(v => {
                        const count = values.filter(v2 => v2.value === v.value).length;
                        if (count > 1) {
                            v.isRepeated = true;
                            v.repetitionCount = count;
                            if (v.severity === 'Info' || v.severity === 'Low') {
                                v.severity = 'Medium';
                            }
                        }
                    });
                }
            }
        });
    }
    
    /**
     * Convert column letter(s) to column number
     */
    private columnLetterToNumber(letters: string): number {
        let result = 0;
        for (let i = 0; i < letters.length; i++) {
            result = result * 26 + (letters.charCodeAt(i) - 64);
        }
        return result;
    }
    
    /**
     * Analyze large worksheets by processing data in chunks to prevent stack overflow
     */
    private async analyzeLargeWorksheet(
        context: Excel.RequestContext,
        worksheet: Excel.Worksheet,
        formulas: any[][],
        values: any[][],
        totalCells: number,
        options: AnalysisOptions
    ): Promise<WorksheetAnalysisResult> {
        
        const CHUNK_SIZE = 50; // Much smaller chunks for massive models
        const totalRows = formulas.length;
        const totalCols = formulas[0]?.length || 0;
        
        // Initialize accumulators
        let totalFormulas = 0;
        const formulaMap = new Map<string, FormulaInfo>();
        const hardCodedValues: HardCodedValue[] = [];
        let cellsWithFormulas = 0;
        let cellsWithValues = 0;
        let emptyCells = 0;
        
        // Build column formula map for row-wise comparison
        const columnFormulas = new Map<number, Array<{row: number, formula: string, normalized: string}>>();
        
        // First pass: Build formula map for pattern analysis
        for (let row = 0; row < totalRows; row++) {
            for (let col = 0; col < totalCols; col++) {
                const formula = formulas[row][col] as string;
                
                if (this.isFormula(formula)) {
                    if (!columnFormulas.has(col)) {
                        columnFormulas.set(col, []);
                    }
                    
                    const normalized = this.normalizeFormulaForComparison(formula);
                    columnFormulas.get(col)!.push({ row, formula, normalized });
                }
            }
        }
        
        // Process data in chunks with context awareness
        for (let startRow = 0; startRow < totalRows; startRow += CHUNK_SIZE) {
            const endRow = Math.min(startRow + CHUNK_SIZE, totalRows);
            
            // Process this chunk
            for (let row = startRow; row < endRow; row++) {
                for (let col = 0; col < totalCols; col++) {
                    const formula = formulas[row][col] as string;
                    const value = values[row][col];
                    
                    // Cell count analysis
                    if (this.isFormula(formula)) {
                        cellsWithFormulas++;
                        
                        // Formula analysis (only if not empty or if including empty cells)
                        if (options.includeEmptyCells || !this.isEmpty(value)) {
                            totalFormulas++;
                            
                            let normalizedFormula = formula;
                            if (options.groupSimilarFormulas) {
                                normalizedFormula = this.normalizeFormula(formula);
                            }
                            
                            const cellAddress = this.getCellAddress(row, col);
                            
                            if (formulaMap.has(normalizedFormula)) {
                                const existing = formulaMap.get(normalizedFormula)!;
                                existing.count++;
                                existing.cells.push(cellAddress);
                            } else {
                                formulaMap.set(normalizedFormula, {
                                    formula: normalizedFormula,
                                    count: 1,
                                    cells: [cellAddress]
                                });
                            }
                            
                            // Hard-coded value analysis with row-wise context
                            if (totalCells < 50000) { // Only run on smaller models
                                const columnContext = columnFormulas.get(col) || [];
                                const detectedValues = this.detectHardCodedValuesWithContext(
                                    formula,
                                    cellAddress,
                                    row,
                                    col,
                                    columnContext,
                                    formulas
                                );
                                hardCodedValues.push(...detectedValues);
                            }
                        }
                    } else if (!this.isEmpty(value)) {
                        cellsWithValues++;
                    } else {
                        emptyCells++;
                    }
                }
            }
            
            // Yield control back to the browser every chunk
            await new Promise(resolve => setTimeout(resolve, 0));
        }
        
        // Calculate percentages
        const formulaPercentage = totalCells > 0 ? (cellsWithFormulas / totalCells) * 100 : 0;
        const valuePercentage = totalCells > 0 ? (cellsWithValues / totalCells) * 100 : 0;
        
        // Create unique formulas list
        const uniqueFormulasList = Array.from(formulaMap.values())
            .sort((a, b) => b.count - a.count);
        
        // Post-process to identify patterns and adjust severity
        this.analyzePatternConsistency(hardCodedValues, formulas);
        
        // Categorize hard-coded values by severity
        const highSeverity = hardCodedValues.filter(v => v.severity === 'High');
        const mediumSeverity = hardCodedValues.filter(v => v.severity === 'Medium');
        const lowSeverity = hardCodedValues.filter(v => v.severity === 'Low');
        const infoSeverity = hardCodedValues.filter(v => v.severity === 'Info');
        
        // Find repeated values
        const valueCounts = new Map<string, number>();
        hardCodedValues.forEach(v => {
            const key = v.value;
            valueCounts.set(key, (valueCounts.get(key) || 0) + 1);
        });
        
        const repeatedValues = hardCodedValues.filter(v => {
            const count = valueCounts.get(v.value) || 0;
            v.repetitionCount = count;
            v.isRepeated = count > 1;
            return count > 1;
        });
        
        // Find undocumented parameters (high/medium severity, repeated)
        const undocumentedParameters = hardCodedValues.filter(v => 
            (v.severity === 'High' || v.severity === 'Medium') && v.isRepeated
        );
        
        return {
            name: worksheet.name,
            totalCells,
            totalFormulas,
            uniqueFormulas: uniqueFormulasList.length,
            uniqueFormulasList,
            formulaComplexity: this.assessComplexityOptimized(uniqueFormulasList, totalFormulas),
            cellCountAnalysis: {
                totalCells,
                cellsWithFormulas,
                cellsWithValues,
                emptyCells,
                formulaPercentage: Math.round(formulaPercentage * 100) / 100,
                valuePercentage: Math.round(valuePercentage * 100) / 100
            },
            hardCodedValueAnalysis: {
                totalHardCodedValues: hardCodedValues.length,
                highSeverityValues: highSeverity,
                mediumSeverityValues: mediumSeverity,
                lowSeverityValues: lowSeverity,
                infoSeverityValues: infoSeverity,
                repeatedValues: repeatedValues,
                undocumentedParameters: undocumentedParameters
            }
        };
    }
    
    /**
     * Optimized hard-coded value detection with improved patterns
     */
    private detectHardCodedValuesOptimized(formula: string, cellAddress: string): HardCodedValue[] {
        const hardCodedValues: HardCodedValue[] = [];
        
        // Improved patterns to catch more hard-coded values
        const patterns = [
            // High confidence patterns
            {
                pattern: /\b(0\.\d{2,})\b/g,
                confidence: 'High' as const,
                reason: 'Decimal percentage value'
            },
            {
                pattern: /\b(1[0-9]{3,})\b/g, // 1000+
                confidence: 'High' as const,
                reason: 'Large numeric constant'
            },
            {
                pattern: /\b([2-9]\d+)\b/g, // 20-99, 200-999, etc.
                confidence: 'High' as const,
                reason: 'Numeric constant'
            },
            {
                pattern: /"[^"]{3,}"/g, // String literals 3+ chars
                confidence: 'High' as const,
                reason: 'String literal'
            },
            // Medium confidence patterns
            {
                pattern: /\b([2-9])\b/g, // Single digits 2-9
                confidence: 'Medium' as const,
                reason: 'Single digit constant'
            },
            {
                pattern: /\b(1[0-9])\b/g, // Numbers 10-19
                confidence: 'Medium' as const,
                reason: 'Two-digit constant'
            }
        ];
        
        patterns.forEach(({ pattern, confidence, reason }) => {
            let match;
            while ((match = pattern.exec(formula)) !== null) {
                const value = match[0];
                const position = match.index;
                
                // Enhanced filtering to avoid false positives
                if (this.isLikelyHardCodedOptimized(value, formula, position)) {
                    hardCodedValues.push({
                        value: value,
                        context: formula,
                        cellAddress: cellAddress,
                        confidence: confidence,
                        reason: reason
                    });
                }
            }
        });
        
        return hardCodedValues;
    }
    
    /**
     * Enhanced filtering for hard-coded values to reduce false positives
     */
    private isLikelyHardCodedOptimized(value: string, formula: string, position: number): boolean {
        // Skip if it's a cell reference (A1, B2, etc.)
        if (/^[A-Z]+\d+$/.test(value)) {
            return false;
        }
        
        // Skip if it's part of a function name
        const functionNames = ['INDEX', 'MATCH', 'VLOOKUP', 'HLOOKUP', 'SUM', 'AVERAGE', 'COUNT', 'IF', 'AND', 'OR', 'INDIRECT'];
        const beforeValue = formula.substring(Math.max(0, position - 15), position);
        if (functionNames.some(func => beforeValue.toUpperCase().includes(func))) {
            return false;
        }
        
        // Skip if it's clearly a column number in array functions
        const context = formula.substring(Math.max(0, position - 20), position + value.length + 20);
        if (/\b(COLUMN|ROW|INDEX|MATCH|VLOOKUP|HLOOKUP)\s*\([^)]*,\s*$/.test(context)) {
            return false;
        }
        
        // Skip if it's a date component (year)
        if (/\b(19|20)\d{2}\b/.test(value) && value.length === 4) {
            return false;
        }
        
        // Skip very small numbers that are likely function parameters
        const numValue = parseFloat(value);
        if (numValue >= 0 && numValue <= 10) {
            // Additional context check for small numbers
            if (/\b(COLUMN|ROW|INDEX|MATCH|VLOOKUP|HLOOKUP)\b/i.test(context)) {
                return false;
            }
        }
        
        return true;
    }
    
    /**
     * Optimized complexity assessment for large worksheets
     */
    private assessComplexityOptimized(formulas: FormulaInfo[], totalFormulas: number): 'Low' | 'Medium' | 'High' {
        if (formulas.length === 0) return 'Low';
        
        const uniqueFormulas = formulas.length;
        const complexFormulas = formulas.filter(f => this.isComplexFormulaOptimized(f.formula)).length;
        
        // Simplified scoring
        let score = 0;
        
        if (totalFormulas > 1000) score += 2;
        else if (totalFormulas > 100) score += 1;
        
        const uniquenessRatio = uniqueFormulas / totalFormulas;
        if (uniquenessRatio > 0.5) score += 2;
        else if (uniquenessRatio > 0.3) score += 1;
        
        const complexityRatio = complexFormulas / uniqueFormulas;
        if (complexityRatio > 0.2) score += 2;
        else if (complexityRatio > 0.1) score += 1;
        
        if (score >= 4) return 'High';
        if (score >= 2) return 'Medium';
        return 'Low';
    }
    
    /**
     * Optimized complex formula detection with safer patterns
     */
    private isComplexFormulaOptimized(formula: string): boolean {
        // Simplified patterns that are safe for large formulas
        const complexPatterns = [
            /INDEX\s*\(/i,
            /MATCH\s*\(/i,
            /VLOOKUP\s*\(/i,
            /HLOOKUP\s*\(/i,
            /XLOOKUP\s*\(/i,
            /SUMIFS\s*\(/i,
            /COUNTIFS\s*\(/i,
            /AVERAGEIFS\s*\(/i,
            /INDIRECT\s*\(/i,
            /OFFSET\s*\(/i,
            /\{[^}]*\}/ // Array constants (bounded)
        ];
        
        return complexPatterns.some(pattern => pattern.test(formula));
    }
    
    /**
     * Minimal analysis for massive models - only basic counting, no complex operations
     */
    private async analyzeMassiveWorksheet(
        context: Excel.RequestContext,
        worksheet: Excel.Worksheet,
        formulas: any[][],
        values: any[][],
        totalCells: number,
        options: AnalysisOptions
    ): Promise<WorksheetAnalysisResult> {
        
        const totalRows = formulas.length;
        const totalCols = formulas[0]?.length || 0;
        
        // Initialize counters - no complex data structures
        let totalFormulas = 0;
        let cellsWithFormulas = 0;
        let cellsWithValues = 0;
        let emptyCells = 0;
        
        // Process in very small chunks to prevent any stack issues
        const MICRO_CHUNK_SIZE = 10; // Process only 10 rows at a time
        
        for (let startRow = 0; startRow < totalRows; startRow += MICRO_CHUNK_SIZE) {
            const endRow = Math.min(startRow + MICRO_CHUNK_SIZE, totalRows);
            
            // Simple counting only - no regex, no complex operations
            for (let row = startRow; row < endRow; row++) {
                for (let col = 0; col < totalCols; col++) {
                    const formula = formulas[row][col] as string;
                    const value = values[row][col];
                    
                    if (this.isFormula(formula)) {
                        cellsWithFormulas++;
                        if (options.includeEmptyCells || !this.isEmpty(value)) {
                            totalFormulas++;
                        }
                    } else if (!this.isEmpty(value)) {
                        cellsWithValues++;
                    } else {
                        emptyCells++;
                    }
                }
            }
            
            // Yield control after every micro-chunk
            await new Promise(resolve => setTimeout(resolve, 0));
        }
        
        // Calculate percentages
        const formulaPercentage = totalCells > 0 ? (cellsWithFormulas / totalCells) * 100 : 0;
        const valuePercentage = totalCells > 0 ? (cellsWithValues / totalCells) * 100 : 0;
        
        // Return minimal results - no formula analysis, no hard-coded values
        return {
            name: worksheet.name,
            totalCells,
            totalFormulas,
            uniqueFormulas: 0, // Skipped for massive models
            uniqueFormulasList: [], // Skipped for massive models
            formulaComplexity: 'High', // Assume high for massive models
            cellCountAnalysis: {
                totalCells,
                cellsWithFormulas,
                cellsWithValues,
                emptyCells,
                formulaPercentage: Math.round(formulaPercentage * 100) / 100,
                valuePercentage: Math.round(valuePercentage * 100) / 100
            },
            hardCodedValueAnalysis: {
                totalHardCodedValues: 0, // Skipped for massive models
                highConfidenceValues: [],
                mediumConfidenceValues: [],
                lowConfidenceValues: []
            }
        };
    }
} 