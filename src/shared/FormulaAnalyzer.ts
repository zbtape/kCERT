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
 * Interface for worksheet analysis results
 */
export interface WorksheetAnalysisResult {
    name: string;
    totalCells: number;
    totalFormulas: number;
    uniqueFormulas: number;
    uniqueFormulasList: FormulaInfo[];
    formulaComplexity: 'Low' | 'Medium' | 'High';
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
}

/**
 * Main class for analyzing Excel formulas
 */
export class FormulaAnalyzer {
    
    /**
     * Analyze all worksheets in the workbook
     */
    async analyzeWorkbook(context: Excel.RequestContext, options: AnalysisOptions): Promise<AnalysisResult> {
        const workbook = context.workbook;
        const worksheets = workbook.worksheets;
        
        worksheets.load('items/name');
        await context.sync();
        
        const worksheetResults: WorksheetAnalysisResult[] = [];
        let totalFormulas = 0;
        const allUniqueFormulas = new Set<string>();
        
        // Analyze each worksheet
        for (const worksheet of worksheets.items) {
            const worksheetResult = await this.analyzeWorksheet(context, worksheet, options);
            worksheetResults.push(worksheetResult);
            totalFormulas += worksheetResult.totalFormulas;
            
            // Add unique formulas to global set
            worksheetResult.uniqueFormulasList.forEach(formulaInfo => {
                allUniqueFormulas.add(formulaInfo.formula);
            });
        }
        
        return {
            totalWorksheets: worksheets.items.length,
            totalFormulas,
            uniqueFormulas: allUniqueFormulas.size,
            worksheets: worksheetResults,
            analysisTimestamp: new Date().toISOString()
        };
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
                formulaComplexity: this.assessComplexity(uniqueFormulasList)
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
                formulaComplexity: 'Low'
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
            /\{.*\}/,                // Array constants
            /.*\(.*\(.*\(.*\)/,      // Nested functions (3+ levels)
        ];
        
        return complexPatterns.some(pattern => pattern.test(formula));
    }
} 