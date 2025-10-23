export type ComplexityBand = 'Low' | 'Medium' | 'High';

export interface FormulaComplexityBreakdown {
    lengthScore: number;
    operatorScore: number;
    depthScore: number;
    functionScore: number;
    arrayBonus: number;
}

export interface FormulaComplexityResult {
    score: number;
    band: ComplexityBand;
    breakdown: FormulaComplexityBreakdown;
    functionCount: number;
    maxDepth: number;
    operatorCount: number;
}

const OPERATOR_PATTERN = /[+\-*/^&<>]=?|(?<![A-Z])MOD(?![A-Z])/gi;
const FUNCTION_PATTERN = /([A-Z_][A-Z0-9_.]*)\s*\(/gi;

export function computeFormulaComplexity(formula: string, isArray: boolean): FormulaComplexityResult {
    const sanitized = sanitizeFormula(formula);
    const lengthScore = Math.min(10, Math.ceil(sanitized.length / 25));
    const operatorCount = countOperators(sanitized);
    const operatorScore = Math.min(12, Math.ceil(operatorCount / 2));
    const { maxDepth, functionCount } = analyseStructure(sanitized);
    const depthScore = Math.min(12, Math.max(maxDepth - 1, 0) * 2);
    const functionScore = Math.min(12, Math.ceil(functionCount / 2));
    const arrayBonus = isArray ? 6 : 0;

    const score = lengthScore + operatorScore + depthScore + functionScore + arrayBonus;
    const band = classify(score, isArray, maxDepth, operatorCount);

    return {
        score,
        band,
        breakdown: { lengthScore, operatorScore, depthScore, functionScore, arrayBonus },
        functionCount,
        maxDepth,
        operatorCount
    };
}

function sanitizeFormula(formula: string): string {
    if (!formula) {
        return '';
    }

    let trimmed = formula.trim();
    if (trimmed.startsWith('=')) {
        trimmed = trimmed.substring(1);
    }
    if (trimmed.startsWith('{') && trimmed.endsWith('}')) {
        trimmed = trimmed.substring(1, trimmed.length - 1);
    }
    return trimmed;
}

function countOperators(expression: string): number {
    if (!expression) {
        return 0;
    }
    const matches = expression.match(OPERATOR_PATTERN);
    return matches ? matches.length : 0;
}

function analyseStructure(expression: string): { maxDepth: number; functionCount: number } {
    let depth = 0;
    let maxDepth = 0;
    let functionCount = 0;

    for (let i = 0; i < expression.length; i++) {
        const char = expression[i];
        if (char === '(') {
            depth++;
            maxDepth = Math.max(maxDepth, depth);
        } else if (char === ')') {
            depth = Math.max(0, depth - 1);
        }
    }

    const functions = expression.match(FUNCTION_PATTERN);
    if (functions) {
        functionCount = functions.length;
    }

    return { maxDepth: Math.max(maxDepth, 1), functionCount };
}

function classify(score: number, isArray: boolean, maxDepth: number, operatorCount: number): ComplexityBand {
    if (isArray || score >= 24 || maxDepth >= 5 || operatorCount >= 12) {
        return 'High';
    }
    if (score >= 14 || maxDepth >= 3 || operatorCount >= 6) {
        return 'Medium';
    }
    return 'Low';
}

