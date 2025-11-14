import { ComplexityBand, FUNCTION_WEIGHTS, getComplexityBandFromFScore } from './ScoringConfig';

// Re-export for backward compatibility
export type { ComplexityBand };

export interface FormulaComplexityBreakdown {
    functionWeightedScore: number;
    depthAdj: number;
    operatorAdj: number;
    arrayAdj: number;
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
    const { maxDepth, functionCount, functionNames } = analyseStructure(sanitized);
    const operatorCount = countOperators(sanitized);
    
    // Calculate function-weighted score
    let functionWeightedScore = 0;
    for (const funcName of functionNames) {
        const weight = FUNCTION_WEIGHTS[funcName] || 2; // Default weight of 2 (moderate complexity) for unknown functions
        functionWeightedScore += weight;
    }
    functionWeightedScore = Math.min(60, functionWeightedScore);
    
    // Structural adjustments
    const depthAdj = Math.min(16, Math.max(maxDepth - 1, 0) * 2);
    const operatorAdj = Math.min(8, Math.ceil(operatorCount / 4));
    const arrayAdj = isArray ? 5 : 0;
    
    // Base score (before usage adjustment)
    const score = functionWeightedScore + depthAdj + operatorAdj + arrayAdj;
    
    // Band is computed from base score for compatibility (will be overridden by analyzer using final F-Score)
    const band = getComplexityBandFromFScore(score);

    return {
        score,
        band,
        breakdown: { functionWeightedScore, depthAdj, operatorAdj, arrayAdj },
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

function analyseStructure(expression: string): { maxDepth: number; functionCount: number; functionNames: string[] } {
    let depth = 0;
    let maxDepth = 0;
    const functionNames: string[] = [];

    for (let i = 0; i < expression.length; i++) {
        const char = expression[i];
        if (char === '(') {
            depth++;
            maxDepth = Math.max(maxDepth, depth);
        } else if (char === ')') {
            depth = Math.max(0, depth - 1);
        }
    }

    const functionMatches = expression.matchAll(FUNCTION_PATTERN);
    for (const match of functionMatches) {
        const funcName = match[1].toUpperCase();
        functionNames.push(funcName);
    }

    return { 
        maxDepth: Math.max(maxDepth, 1), 
        functionCount: functionNames.length,
        functionNames
    };
}


