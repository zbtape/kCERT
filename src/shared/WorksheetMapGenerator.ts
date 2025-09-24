import { isEmpty, isFormula } from './FormulaUtils';

export type MapSymbol = 'F' | '<' | '^' | '+' | 'L' | 'N' | 'A' | '';

type MapSymbolKey = Exclude<MapSymbol, ''>;

type MapCounts = Record<MapSymbolKey, number>;

type MapAnomalies = {
    changeOfDirection: number;
    horizontalBreaks: number;
    verticalBreaks: number;
};

export type MapProgressReporter = (message: string) => void;

export interface MapOptions {
    includeHidden: boolean;
    maxCells?: number;
}

export interface WorksheetMapResult {
    worksheetName: string;
    usedRangeAddress: string;
    startRow: number;
    startColumn: number;
    rowCount: number;
    columnCount: number;
    symbols: MapSymbol[][];
    counts: MapCounts;
    arrayAreas: Array<{ top: number; left: number; bottom: number; right: number }>;
    anomalies: MapAnomalies;
    skipped?: boolean;
    skipReason?: string;
}

const DEFAULT_COUNTS = (): MapCounts => ({ F: 0, '<': 0, '^': 0, '+': 0, L: 0, N: 0, A: 0 });

interface DynamicAnchor {
    absoluteRow: number;
    absoluteColumn: number;
}

export class WorksheetMapGenerator {
    private static readonly ROW_BLOCK = 200;
    private static readonly COL_BLOCK = 120;
    private static readonly MAX_DEFAULT_CELLS = 250_000;

    constructor(private readonly options: MapOptions = { includeHidden: true }) {}

    async generate(context: Excel.RequestContext, worksheet: Excel.Worksheet, progress?: MapProgressReporter): Promise<WorksheetMapResult> {
        worksheet.load(['name']);
        await context.sync();

        const usedRange = worksheet.getUsedRange(this.options.includeHidden);
        usedRange.load(['rowCount', 'columnCount', 'address', 'rowIndex', 'columnIndex']);
        await context.sync();

        if (usedRange.isNullObject || !usedRange.rowCount || !usedRange.columnCount) {
            return this.emptyResult(worksheet.name, usedRange.address ?? 'A1', usedRange.rowIndex ?? 0, usedRange.columnIndex ?? 0, 'no_used_range');
        }

        const totalCells = usedRange.rowCount * usedRange.columnCount;
        const maxCells = this.options.maxCells ?? WorksheetMapGenerator.MAX_DEFAULT_CELLS;
        if (totalCells > maxCells) {
            return this.emptyResult(worksheet.name, usedRange.address, usedRange.rowIndex, usedRange.columnIndex, `exceeds_max_cells_${maxCells}`);
        }

        const rowCount = usedRange.rowCount;
        const columnCount = usedRange.columnCount;
        const startRow = usedRange.rowIndex;
        const startColumn = usedRange.columnIndex;

        const symbols: MapSymbol[][] = Array.from({ length: rowCount }, () => new Array<MapSymbol>(columnCount).fill(''));
        const counts = DEFAULT_COUNTS();
        const arrayAreas: WorksheetMapResult['arrayAreas'] = [];
        const dynamicAnchors = new Map<string, DynamicAnchor>();

        const rowFormulaState: (string | null)[] = new Array(rowCount).fill(null);
        const columnFormulaState: (string | null)[] = new Array(columnCount).fill(null);

        for (let rowStart = 0; rowStart < rowCount; rowStart += WorksheetMapGenerator.ROW_BLOCK) {
            const blockRows = Math.min(WorksheetMapGenerator.ROW_BLOCK, rowCount - rowStart);
            for (let colStart = 0; colStart < columnCount; colStart += WorksheetMapGenerator.COL_BLOCK) {
                const blockCols = Math.min(WorksheetMapGenerator.COL_BLOCK, columnCount - colStart);
                const block = usedRange.getCell(rowStart, colStart).getResizedRange(blockRows - 1, blockCols - 1);
                block.load(['rowIndex', 'columnIndex', 'formulasR1C1', 'values', 'text']);
                await context.sync();

                const formulas = block.formulasR1C1 as (string | null)[][];
                const values = block.values as any[][];
                const texts = block.text as (string | null)[][];

                const rowOffset = startRow + rowStart + 1;
                const colOffset = startColumn + colStart + 1;
                progress?.(`Mapping ${worksheet.name}: rows ${rowOffset}-${rowOffset + blockRows - 1}, cols ${colOffset}-${colOffset + blockCols - 1}`);

                for (let localRow = 0; localRow < blockRows; localRow++) {
                    for (let localCol = 0; localCol < blockCols; localCol++) {
                        const relativeRow = rowStart + localRow;
                        const relativeCol = colStart + localCol;
                        const formula = formulas[localRow][localCol];
                        const value = values[localRow][localCol];
                        const text = texts[localRow][localCol];

                        const symbol = this.classifyCell(
                            worksheet,
                            startRow + relativeRow,
                            startColumn + relativeCol,
                            formula,
                            value,
                            text,
                            relativeRow,
                            relativeCol,
                            rowFormulaState,
                            columnFormulaState,
                            counts,
                            dynamicAnchors
                        );
                        symbols[relativeRow][relativeCol] = symbol;
                    }
                }
            }
        }

        await this.promoteDynamicArrays(
            context,
            worksheet,
            usedRange,
            rowCount,
            columnCount,
            dynamicAnchors,
            symbols,
            counts,
            arrayAreas
        );

        const anomalies = this.analyseAnomalies(symbols);

        return {
            worksheetName: worksheet.name,
            usedRangeAddress: usedRange.address,
            startRow,
            startColumn,
            rowCount,
            columnCount,
            symbols,
            counts,
            arrayAreas,
            anomalies,
        };
    }

    private classifyCell(
        worksheet: Excel.Worksheet,
        absoluteRow: number,
        absoluteColumn: number,
        formulaR1C1: string | null,
        value: any,
        text: string | null,
        row: number,
        col: number,
        rowFormulaState: (string | null)[],
        columnFormulaState: (string | null)[],
        counts: MapCounts,
        dynamicAnchors: Map<string, DynamicAnchor>
    ): MapSymbol {
        const trimmedFormula = typeof formulaR1C1 === 'string' ? formulaR1C1.trim() : '';

        if (trimmedFormula.startsWith('{=') && trimmedFormula.endsWith('}')) {
            rowFormulaState[row] = null;
            columnFormulaState[col] = null;
            counts['A']++;
            return 'A';
        }

        if (trimmedFormula && isFormula(trimmedFormula)) {
            if (trimmedFormula.includes('#')) {
                const key = `${absoluteRow}:${absoluteColumn}`;
                if (!dynamicAnchors.has(key)) {
                    dynamicAnchors.set(key, { absoluteRow, absoluteColumn });
                }
            }

            const leftKey = col > 0 ? rowFormulaState[row] : null;
            const upKey = columnFormulaState[col];

            let symbol: MapSymbol;
            if (leftKey === trimmedFormula && upKey === trimmedFormula) {
                symbol = '+';
            } else if (leftKey === trimmedFormula) {
                symbol = '<';
            } else if (upKey === trimmedFormula) {
                symbol = '^';
            } else {
                symbol = 'F';
            }

            rowFormulaState[row] = trimmedFormula;
            columnFormulaState[col] = trimmedFormula;
            counts[symbol as MapSymbolKey]++;
            return symbol;
        }

        rowFormulaState[row] = null;
        columnFormulaState[col] = null;

        if (!isEmpty(value)) {
            if (typeof value === 'number' || typeof value === 'boolean') {
                counts['N']++;
                return 'N';
            }
            counts['L']++;
            return 'L';
        }

        if (text && text.trim().length > 0) {
            counts['L']++;
            return 'L';
        }

        return '';
    }

    private async promoteDynamicArrays(
        context: Excel.RequestContext,
        worksheet: Excel.Worksheet,
        usedRange: Excel.Range,
        rowCount: number,
        columnCount: number,
        dynamicAnchors: Map<string, DynamicAnchor>,
        symbols: MapSymbol[][],
        counts: MapCounts,
        arrayAreas: WorksheetMapResult['arrayAreas']
    ): Promise<void> {
        if (dynamicAnchors.size === 0) {
            return;
        }

        const anchors = Array.from(dynamicAnchors.values());
        const topOffset = usedRange.rowIndex;
        const leftOffset = usedRange.columnIndex;

        for (const anchor of anchors) {
            const cell = worksheet.getCell(anchor.absoluteRow, anchor.absoluteColumn);
            const spillRange = cell.getSpillingToRangeOrNullObject();
            spillRange.load(['isNullObject', 'rowIndex', 'columnIndex', 'rowCount', 'columnCount']);
            await context.sync();

            if (spillRange.isNullObject) {
                continue;
            }

            const top = spillRange.rowIndex - topOffset;
            const left = spillRange.columnIndex - leftOffset;
            const bottom = top + spillRange.rowCount - 1;
            const right = left + spillRange.columnCount - 1;

            if (top < 0 || left < 0 || bottom >= rowCount || right >= columnCount) {
                continue;
            }

            let promotedCells = 0;
            for (let r = top; r <= bottom; r++) {
                for (let c = left; c <= right; c++) {
                    const previous = symbols[r][c];
                    if (previous === 'A') {
                        continue;
                    }
                    if (previous && previous !== '') {
                        counts[previous as MapSymbolKey] = Math.max(0, counts[previous as MapSymbolKey] - 1);
                    }
                    symbols[r][c] = 'A';
                    promotedCells++;
                }
            }

            if (promotedCells > 0) {
                counts['A'] += promotedCells;
                arrayAreas.push({ top, left, bottom, right });
            }
        }
    }

    private analyseAnomalies(symbols: MapSymbol[][]): MapAnomalies {
        let changeOfDirection = 0;
        let horizontalBreaks = 0;
        let verticalBreaks = 0;

        const rows = symbols.length;
        if (rows === 0) {
            return { changeOfDirection, horizontalBreaks, verticalBreaks };
        }
        const cols = symbols[0].length;

        for (const rowSymbols of symbols) {
            let lastDirection: 'left' | 'up' | null = null;
            let inCopySequence = false;
            for (const symbol of rowSymbols) {
                if (symbol === '<' || symbol === '+') {
                    if (lastDirection && lastDirection !== 'left') {
                        changeOfDirection++;
                    }
                    lastDirection = 'left';
                    inCopySequence = true;
                } else if (symbol === '^') {
                    if (lastDirection && lastDirection !== 'up') {
                        changeOfDirection++;
                    }
                    lastDirection = 'up';
                    inCopySequence = true;
                } else if (symbol === 'N' || symbol === 'L' || symbol === '') {
                    if (inCopySequence) {
                        horizontalBreaks++;
                    }
                    inCopySequence = false;
                    lastDirection = null;
                } else {
                    inCopySequence = false;
                    lastDirection = null;
                }
            }
        }

        for (let c = 0; c < cols; c++) {
            let inCopySequence = false;
            for (let r = 0; r < rows; r++) {
                const symbol = symbols[r][c];
                if (symbol === '^' || symbol === '+') {
                    inCopySequence = true;
                } else if (symbol === 'N' || symbol === 'L' || symbol === '') {
                    if (inCopySequence) {
                        verticalBreaks++;
                    }
                    inCopySequence = false;
                } else {
                    inCopySequence = false;
                }
            }
        }

        return { changeOfDirection, horizontalBreaks, verticalBreaks };
    }

    private emptyResult(
        worksheetName: string,
        usedRangeAddress: string,
        startRow: number,
        startColumn: number,
        reason: string
    ): WorksheetMapResult {
        return {
            worksheetName,
            usedRangeAddress,
            startRow,
            startColumn,
            rowCount: 0,
            columnCount: 0,
            symbols: [],
            counts: DEFAULT_COUNTS(),
            arrayAreas: [],
            anomalies: { changeOfDirection: 0, horizontalBreaks: 0, verticalBreaks: 0 },
            skipped: true,
            skipReason: reason,
        };
    }
}
