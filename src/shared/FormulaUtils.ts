export function columnLetter(index: number): string {
    let letter = '';
    let tempIndex = index;
    while (tempIndex > 0) {
        const remainder = (tempIndex - 1) % 26;
        letter = String.fromCharCode(65 + remainder) + letter;
        tempIndex = Math.floor((tempIndex - 1) / 26);
    }
    return letter;
}

export function getCellAddress(row: number, col: number): string {
    return `${columnLetter(col + 1)}${row + 1}`;
}

export function isFormula(value: unknown): value is string {
    return typeof value === 'string' && value.startsWith('=');
}

export function isEmpty(value: unknown): boolean {
    return value === null || value === undefined || value === '';
}

export function normalizeFormula(formula: string): string {
    let normalized = formula;
    normalized = normalized.replace(/\$[A-Z]+\$\d+/g, '<ABS>');
    normalized = normalized.replace(/\$?[A-Z]+\$?\d+/g, '<REL>');
    normalized = normalized.replace(/<REL>:<REL>/g, '<RANGE>');
    normalized = normalized.replace(/<ABS>:<ABS>/g, '<ABS_RANGE>');
    return normalized;
}

export function twoDecimals(value: number): number {
    return Math.round(value * 100) / 100;
}
