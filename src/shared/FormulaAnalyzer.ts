/**
 * Streaming-first formula analyzer
 * --------------------------------
 * Designed to process massive workbooks by loading fixed-size blocks of cells,
 * analysing each block immediately, and discarding it before moving on.
 */

import { getCellAddress, isEmpty, isFormula, normalizeFormula, twoDecimals } from './FormulaUtils';

export interface AnalysisOptions {
    includeEmptyCells: boolean;
    groupSimilarFormulas: boolean;
}

export interface FormulaInfo {
    formula: string;
    count: number;
    cells: string[];
}

export interface CellCountAnalysis {
    totalCells: number;
    cellsWithFormulas: number;
    cellsWithValues: number;
    emptyCells: number;
    formulaPercentage: number;
    valuePercentage: number;
}

export interface HardCodedValue {
    value: string;
    context: string;
    cellAddress: string;
    severity: 'High' | 'Medium' | 'Low' | 'Info';
    rationale: string;
    suggestedFix: string;
    isRepeated: boolean;
    repetitionCount: number;
}

export interface HardCodedValueAnalysis {
    totalHardCodedValues: number;
    highSeverityValues: HardCodedValue[];
    mediumSeverityValues: HardCodedValue[];
    lowSeverityValues: HardCodedValue[];
    infoSeverityValues: HardCodedValue[];
    repeatedValues: HardCodedValue[];
    undocumentedParameters: HardCodedValue[];
}

export interface WorksheetAnalysisResult {
    name: string;
    totalCells: number;
    totalFormulas: number;
    uniqueFormulas: number;
    uniqueFormulasList: FormulaInfo[];
    formulaComplexity: 'Low' | 'Medium' | 'High';
    cellCountAnalysis: CellCountAnalysis;
    hardCodedValueAnalysis: HardCodedValueAnalysis;
    analysisMode?: 'streaming' | 'massive-skim' | 'skipped';
    fallbackReason?: string;
}

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

export type ProgressReporter = (message: string) => void;

export class FormulaAnalyzer {
    private static readonly ROW_BLOCK = 200;
    private static readonly COL_BLOCK = 120;
    private static readonly MASSIVE_THRESHOLD = 150_000; // total cells
    private static readonly MAX_FORMULA_SAMPLE_CELLS = 200;
    private static readonly MAX_HARDCODED_PER_SHEET = 400;

    async analyzeWorkbook(
        context: Excel.RequestContext,
        options: AnalysisOptions,
        progress?: ProgressReporter
    ): Promise<AnalysisResult> {
        const worksheets = context.workbook.worksheets;
        worksheets.load('items/name');
        await context.sync();

        const skipSheets = new Set(['kCERT_Analysis_Report']);
        const targets = worksheets.items.filter(ws => !skipSheets.has(ws.name));

        const worksheetResults: WorksheetAnalysisResult[] = [];
        let totalFormulas = 0;
        let totalCells = 0;
        let totalCellsWithFormulas = 0;
        let totalCellsWithValues = 0;
        let totalHardCodes = 0;
        const uniqueSet = new Set<string>();

        for (const ws of targets) {
            progress?.(`Starting worksheet "${ws.name}"`);
            const result = await this.analyzeWorksheetStreaming(context, ws, options, progress);
            worksheetResults.push(result);

            totalFormulas += result.totalFormulas;
            totalCells += result.cellCountAnalysis.totalCells;
            totalCellsWithFormulas += result.cellCountAnalysis.cellsWithFormulas;
            totalCellsWithValues += result.cellCountAnalysis.cellsWithValues;
            totalHardCodes += result.hardCodedValueAnalysis.totalHardCodedValues;
            result.uniqueFormulasList.forEach(info => uniqueSet.add(info.formula));

            await new Promise(resolve => setTimeout(resolve, 0));
        }

        return {
            totalWorksheets: worksheetResults.length,
            totalFormulas,
            uniqueFormulas: uniqueSet.size,
            worksheets: worksheetResults,
            analysisTimestamp: new Date().toISOString(),
            totalCells,
            totalCellsWithFormulas,
            totalCellsWithValues,
            totalHardCodedValues: totalHardCodes
        };
    }

    private async analyzeWorksheetStreaming(
        context: Excel.RequestContext,
        worksheet: Excel.Worksheet,
        options: AnalysisOptions,
        progress?: ProgressReporter
    ): Promise<WorksheetAnalysisResult> {
        worksheet.load('name');
        await context.sync();

        const usedRange = worksheet.getUsedRange(true);
        usedRange.load(['rowCount', 'columnCount', 'rowIndex', 'columnIndex']);
        await context.sync();

        if (usedRange.isNullObject || !usedRange.rowCount || !usedRange.columnCount) {
            return this.emptyResult(worksheet.name, 'skipped', 'no_used_range');
        }

        const totalCells = usedRange.rowCount * usedRange.columnCount;
        if (totalCells >= FormulaAnalyzer.MASSIVE_THRESHOLD) {
            progress?.(`Worksheet "${worksheet.name}" exceeds ${FormulaAnalyzer.MASSIVE_THRESHOLD} cells, using skim mode`);
            return await this.analyzeWorksheetSkim(context, worksheet, usedRange);
        }

        const formulaMap = new Map<string, FormulaInfo>();
        const hardCodes: HardCodedValue[] = [];
        let totalFormulas = 0;
        let cellsWithFormulas = 0;
        let cellsWithValues = 0;

        for (let rowStart = 0; rowStart < usedRange.rowCount; rowStart += FormulaAnalyzer.ROW_BLOCK) {
            const rowHeight = Math.min(FormulaAnalyzer.ROW_BLOCK, usedRange.rowCount - rowStart);
            for (let colStart = 0; colStart < usedRange.columnCount; colStart += FormulaAnalyzer.COL_BLOCK) {
                const colWidth = Math.min(FormulaAnalyzer.COL_BLOCK, usedRange.columnCount - colStart);
                const block = this.getBlock(usedRange, rowStart, colStart, rowHeight, colWidth);
                block.load(['formulas', 'values']);
                await context.sync();

                const formulas = block.formulas as any[][];
                const values = block.values as any[][];
                progress?.(
                    `Streaming ${worksheet.name}: rows ${usedRange.rowIndex + rowStart + 1}-${usedRange.rowIndex + rowStart + rowHeight}, ` +
                    `cols ${usedRange.columnIndex + colStart + 1}-${usedRange.columnIndex + colStart + colWidth}`
                );

                for (let r = 0; r < formulas.length; r++) {
                    for (let c = 0; c < formulas[r].length; c++) {
                        const formula = formulas[r][c] as string;
                        const value = values[r][c];
                        const absRow = usedRange.rowIndex + rowStart + r;
                        const absCol = usedRange.columnIndex + colStart + c;
                        const address = getCellAddress(absRow, absCol);

                        if (isFormula(formula)) {
                            cellsWithFormulas++;
                            totalFormulas++;

                            const key = options.groupSimilarFormulas
                                ? normalizeFormula(formula)
                                : formula;

                            let info = formulaMap.get(key);
                            if (!info) {
                                info = { formula: key, count: 0, cells: [] };
                                formulaMap.set(key, info);
                            }
                            info.count++;
                            if (info.cells.length < FormulaAnalyzer.MAX_FORMULA_SAMPLE_CELLS) {
                                info.cells.push(address);
                            }

                            if (hardCodes.length < FormulaAnalyzer.MAX_HARDCODED_PER_SHEET) {
                                hardCodes.push(...this.detectHardCodedValues(formula, address));
                            }
                        } else if (!isEmpty(value)) {
                            cellsWithValues++;
                        }
                    }
                }
            }
        }

        const emptyCells = Math.max(totalCells - cellsWithFormulas - cellsWithValues, 0);

        return {
            name: worksheet.name,
            totalCells,
            totalFormulas,
            uniqueFormulas: formulaMap.size,
            uniqueFormulasList: Array.from(formulaMap.values()).sort((a, b) => b.count - a.count),
            formulaComplexity: this.assessComplexity(totalFormulas, formulaMap),
            cellCountAnalysis: {
                totalCells,
                cellsWithFormulas,
                cellsWithValues,
                emptyCells,
                formulaPercentage: twoDecimals(totalCells > 0 ? (cellsWithFormulas / totalCells) * 100 : 0),
                valuePercentage: twoDecimals(totalCells > 0 ? (cellsWithValues / totalCells) * 100 : 0)
            },
            hardCodedValueAnalysis: this.finaliseHardCodes(hardCodes),
            analysisMode: 'streaming'
        };
    }

    private async analyzeWorksheetSkim(
        context: Excel.RequestContext,
        worksheet: Excel.Worksheet,
        usedRange: Excel.Range
    ): Promise<WorksheetAnalysisResult> {
        const totalCells = usedRange.rowCount * usedRange.columnCount;
        let cellsWithFormulas = 0;
        let cellsWithValues = 0;

        const formulasRange = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
        const constantsRange = usedRange.getSpecialCellsOrNullObject(Excel.SpecialCellType.constants);
        formulasRange.load('address');
        constantsRange.load('address');

        let formulasCountResult: Excel.ClientResult<number> | null = null;
        let constantsCountResult: Excel.ClientResult<number> | null = null;

        try {
            if (!formulasRange.isNullObject) {
                formulasCountResult = formulasRange.getCellCount();
            }
            if (!constantsRange.isNullObject) {
                constantsCountResult = constantsRange.getCellCount();
            }
            await context.sync();

            cellsWithFormulas = formulasCountResult ? formulasCountResult.value : 0;
            cellsWithValues = constantsCountResult ? constantsCountResult.value : 0;
        } catch (error) {
            console.warn(`Skim mode special cells failed for worksheet "${worksheet.name}"`, error);
        }

        const emptyCells = Math.max(totalCells - cellsWithFormulas - cellsWithValues, 0);

        return {
            name: worksheet.name,
            totalCells,
            totalFormulas: cellsWithFormulas,
            uniqueFormulas: 0,
            uniqueFormulasList: [],
            formulaComplexity: 'High',
            cellCountAnalysis: {
                totalCells,
                cellsWithFormulas,
                cellsWithValues,
                emptyCells,
                formulaPercentage: twoDecimals(totalCells > 0 ? (cellsWithFormulas / totalCells) * 100 : 0),
                valuePercentage: twoDecimals(totalCells > 0 ? (cellsWithValues / totalCells) * 100 : 0)
            },
            hardCodedValueAnalysis: {
                totalHardCodedValues: 0,
                highSeverityValues: [],
                mediumSeverityValues: [],
                lowSeverityValues: [],
                infoSeverityValues: [],
                repeatedValues: [],
                undocumentedParameters: []
            },
            analysisMode: 'massive-skim',
            fallbackReason: 'massive_threshold'
        };
    }

    private detectHardCodedValues(formula: string, address: string): HardCodedValue[] {
        const results: HardCodedValue[] = [];
        if (!isFormula(formula)) {
            return results;
        }

        const push = (value: string, severity: HardCodedValue['severity'], index: number) => {
            results.push({
                value,
                severity,
                context: this.snippet(formula, index, 60),
                cellAddress: address,
                rationale: this.rationaleForSeverity(severity, value),
                suggestedFix: this.fixForSeverity(severity),
                isRepeated: false,
                repetitionCount: 0
            });
        };

        const numericPattern = /(?<![A-Z])(-?\d+(?:\.\d+)?)(?![A-Z0-9])/gi;
        const stringPattern = /"([^"\r\n]*)"/g;
        const arrayPattern = /\{([^}]+)\}/g;

        let match: RegExpExecArray | null;
        while ((match = numericPattern.exec(formula)) !== null) {
            push(match[1], this.severityForNumber(match[1]), match.index);
        }
        while ((match = stringPattern.exec(formula)) !== null) {
            push(`"${match[1]}"`, match[1].length >= 6 ? 'Medium' : match[1].length >= 3 ? 'Low' : 'Info', match.index);
        }
        while ((match = arrayPattern.exec(formula)) !== null) {
            push(`{${match[1]}}`, 'High', match.index);
        }

        return results;
    }

    private finaliseHardCodes(values: HardCodedValue[]): HardCodedValueAnalysis {
        if (values.length === 0) {
            return {
                totalHardCodedValues: 0,
                highSeverityValues: [],
                mediumSeverityValues: [],
                lowSeverityValues: [],
                infoSeverityValues: [],
                repeatedValues: [],
                undocumentedParameters: []
            };
        }

        const valueGroups = new Map<string, HardCodedValue[]>();
        values.forEach(item => {
            const key = item.value;
            if (!valueGroups.has(key)) {
                valueGroups.set(key, []);
            }
            valueGroups.get(key)!.push(item);
        });

        const repeated: HardCodedValue[] = [];
        valueGroups.forEach(group => {
            if (group.length > 1) {
                group.forEach(entry => {
                    entry.isRepeated = true;
                    entry.repetitionCount = group.length;
                    repeated.push(entry);
                });
            }
        });

        return {
            totalHardCodedValues: values.length,
            highSeverityValues: values.filter(v => v.severity === 'High'),
            mediumSeverityValues: values.filter(v => v.severity === 'Medium'),
            lowSeverityValues: values.filter(v => v.severity === 'Low'),
            infoSeverityValues: values.filter(v => v.severity === 'Info'),
            repeatedValues: repeated,
            undocumentedParameters: values.filter(v => (v.severity === 'High' || v.severity === 'Medium') && v.isRepeated)
        };
    }

    private assessComplexity(totalFormulas: number, map: Map<string, FormulaInfo>): 'Low' | 'Medium' | 'High' {
        if (totalFormulas === 0) return 'Low';
        const unique = map.size;
        const uniquenessRatio = unique / totalFormulas;

        let score = 0;
        if (totalFormulas > 5_000) score += 2;
        else if (totalFormulas > 500) score += 1;

        if (uniquenessRatio > 0.5) score += 2;
        else if (uniquenessRatio > 0.3) score += 1;

        if (unique > 200) score += 1;

        if (score >= 4) return 'High';
        if (score >= 2) return 'Medium';
        return 'Low';
    }

    private emptyResult(name: string, mode: WorksheetAnalysisResult['analysisMode'], reason?: string): WorksheetAnalysisResult {
        return {
            name,
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
            },
            analysisMode: mode,
            fallbackReason: reason
        };
    }

    private getBlock(range: Excel.Range, rowStart: number, colStart: number, rowCount: number, colCount: number): Excel.Range {
        return range.getCell(rowStart, colStart).getResizedRange(rowCount - 1, colCount - 1);
    }

    private severityForNumber(value: string): HardCodedValue['severity'] {
        const num = Number(value);
        if (Number.isNaN(num)) {
            return 'Info';
        }
        if (Math.abs(num) >= 1000 || value.includes('.')) {
            return 'High';
        }
        if (Math.abs(num) >= 100) {
            return 'Medium';
        }
        if (Math.abs(num) >= 10) {
            return 'Low';
        }
        return 'Info';
    }

    private rationaleForSeverity(severity: HardCodedValue['severity'], value: string): string {
        switch (severity) {
            case 'High':
                return `Large or precise literal ${value} embedded in formula.`;
            case 'Medium':
                return `Potential driver literal ${value} detected inside formula.`;
            case 'Low':
                return `Literal ${value} may be a configuration constant.`;
            default:
                return `Literal ${value} found in formula.`;
        }
    }

    private fixForSeverity(severity: HardCodedValue['severity']): string {
        switch (severity) {
            case 'High':
                return 'Move value to an inputs sheet and reference it.';
            case 'Medium':
                return 'Consider referencing a named range instead of embedding the literal.';
            case 'Low':
                return 'Review whether this literal should be parameterised.';
            default:
                return 'Confirm with modelling standards whether this literal is acceptable.';
        }
    }

    private snippet(formula: string, index: number, span: number): string {
        const start = Math.max(0, index - Math.floor(span / 2));
        const end = Math.min(formula.length, index + Math.floor(span / 2));
        let snippet = formula.substring(start, end);
        if (start > 0) snippet = '…' + snippet;
        if (end < formula.length) snippet += '…';
        return snippet;
    }
}
