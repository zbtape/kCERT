import { tokenize } from 'excel-formula-tokenizer';
import type { Token } from 'excel-formula-tokenizer';

export type LiteralKind =
    | 'numeric'
    | 'percentage'
    | 'string'
    | 'boolean'
    | 'date'
    | 'time'
    | 'array'
    | 'named-literal'
    | 'external'
    | 'hexadecimal'
    | 'unknown';

export interface HardCodedLiteral {
    kind: LiteralKind;
    value: string;
    display: string;
    index: number;
    length: number;
    context?: string;
    rationale?: string;
    suggestedFix?: string;
    absoluteValue?: number;
    containsMixedTypes?: boolean;
    isFlag?: boolean;
    isLikelyInput?: boolean;
    isLinkedWorkbook?: boolean;
    isRecognizedConstant?: boolean;
    parentFunction?: string;
}

const DATE_FUNCTIONS = new Set([
    'DATE',
    'DATEVALUE',
    'EDATE',
    'EOMONTH',
    'TODAY',
    'NOW',
    'WORKDAY',
    'WORKDAY.INTL',
    'WEEKDAY',
    'WEEKNUM',
    'YEAR',
    'MONTH',
    'DAY'
]);

const TIME_FUNCTIONS = new Set([
    'TIME',
    'TIMEVALUE',
    'HOUR',
    'MINUTE',
    'SECOND'
]);

const SAFE_STRING_CONSTANTS = new Set(['Y', 'N', 'A', 'B', 'X']);
const KNOWN_CONSTANTS = new Set(['PI', 'E', 'PHI', 'TRUE', 'FALSE']);

const HEX_PATTERN = /^0x[0-9a-f]+$/i;
const TIME_PATTERN = /^([01]?\d|2[0-3]):[0-5]\d(?::[0-5]\d(?:\.\d{1,3})?)?$/;
const DATE_LITERAL_PATTERN = /^(\d{1,2}[/-]){2}\d{2,4}$/;

interface TokenWithMeta {
    value: string;
    type: string;
    subtype: string;
    index: number;
    length: number;
}

interface TokenMeta extends TokenWithMeta {
    originalIndex: number;
}

interface ParentContext {
    functionName: string | null;
    argumentIndex: number;
}

export function detectHardCodedLiterals(formula: string): HardCodedLiteral[] {
    if (!formula.startsWith('=')) {
        return [];
    }

    const expression = formula.trimStart().replace(/^=/, '');
    let tokens: TokenMeta[];
    try {
        tokens = tokenizeWithOffsets(expression);
    } catch (error) {
        console.warn('HardCodeDetector: tokenizer failed; falling back to regex', error);
        return fallbackDetection(expression);
    }

    const literals: HardCodedLiteral[] = [];
    const stack: ParentContext[] = [];

    tokens.forEach((token, idx) => {
        if (token.type === 'function') {
            if (token.subtype === 'start') {
                stack.push({ functionName: token.value || null, argumentIndex: 0 });
            } else if (token.subtype === 'stop') {
                stack.pop();
            }
            return;
        }

        if (token.type === 'argument') {
            if (stack.length) {
                stack[stack.length - 1].argumentIndex += 1;
            }
            return;
        }

        const parent = stack.length ? stack[stack.length - 1] : null;

        if (token.type === 'operand') {
            switch (token.subtype) {
                case 'number':
                    literals.push(handleNumberToken(token, parent, tokens, idx));
                    break;
                case 'logical':
                    literals.push(handleBooleanToken(token));
                    break;
                case 'text':
                    literals.push(handleTextToken(token, parent));
                    break;
                case 'range':
                    if (isExternalReference(token.value)) {
                        literals.push(handleExternalReference(token));
                    }
                    break;
                default:
                    break;
            }
        }
    });

    annotateArrayContext(tokens, literals);
    return literals.filter(literal => !isBenignLiteral(literal));
}

function tokenizeWithOffsets(expression: string): TokenMeta[] {
    const rawTokens = tokenize(expression);
    const result: TokenMeta[] = [];
    let offset = 0;

    (rawTokens as Token[]).forEach((token, idx) => {
        const value = typeof token.value === 'string' ? token.value : '';
        if (!value) {
            result.push({ value: '', type: token.type, subtype: token.subtype, index: offset, length: 0, originalIndex: idx });
            return;
        }

        const start = expression.indexOf(value, offset);
        const index = start >= 0 ? start : offset;
        result.push({ value, type: token.type, subtype: token.subtype, index, length: value.length, originalIndex: idx });
        offset = index + value.length;
    });

    return result;
}

function handleNumberToken(token: TokenMeta, parent: ParentContext | null, tokens: TokenMeta[], tokenIndex: number): HardCodedLiteral {
    const numericInfo = classifyNumeric(token.value, parent?.functionName ?? null, parent?.argumentIndex ?? 0, tokens, tokenIndex);
    return {
        kind: numericInfo.kind,
        value: token.value,
        display: numericInfo.display,
        index: token.index,
        length: token.length,
        absoluteValue: numericInfo.absoluteValue,
        isLikelyInput: numericInfo.isLikelyInput,
        containsMixedTypes: numericInfo.containsMixedTypes,
        rationale: numericInfo.rationale,
        suggestedFix: numericInfo.suggestedFix
    };
}

function handleBooleanToken(token: TokenMeta): HardCodedLiteral {
    const valueUpper = token.value.toUpperCase();
    return {
        kind: 'boolean',
        value: valueUpper,
        display: valueUpper,
        index: token.index,
        length: token.length,
        isFlag: valueUpper === 'TRUE' || valueUpper === 'FALSE'
    };
}

function handleTextToken(token: TokenMeta, parent: ParentContext | null): HardCodedLiteral {
    const value = token.value;
    const parentName = parent?.functionName?.toUpperCase() ?? null;
    const display = `"${value}"`;

    if (SAFE_STRING_CONSTANTS.has(value)) {
        return {
            kind: 'string',
            value,
            display,
            index: token.index,
            length: token.length + 2,
            isFlag: true
        };
    }

    if (KNOWN_CONSTANTS.has(value.toUpperCase())) {
        return {
            kind: 'named-literal',
            value,
            display,
            index: token.index,
            length: token.length + 2,
            isRecognizedConstant: true
        };
    }

    if ((parentName && DATE_FUNCTIONS.has(parentName)) || DATE_LITERAL_PATTERN.test(value)) {
        return {
            kind: 'date',
            value,
            display,
            index: token.index,
            length: token.length + 2,
            isLikelyInput: true,
            rationale: 'Date literal embedded directly in formula.'
        };
    }

    if ((parentName && TIME_FUNCTIONS.has(parentName)) || TIME_PATTERN.test(value)) {
        return {
            kind: 'time',
            value,
            display,
            index: token.index,
            length: token.length + 2,
            isLikelyInput: true
        };
    }

    if (HEX_PATTERN.test(value)) {
        return {
            kind: 'hexadecimal',
            value,
            display,
            index: token.index,
            length: token.length + 2
        };
    }

    return {
        kind: 'string',
        value,
        display,
        index: token.index,
        length: token.length + 2
    };
}

function handleExternalReference(token: TokenMeta): HardCodedLiteral {
    return {
        kind: 'external',
        value: token.value,
        display: token.value,
        index: token.index,
        length: token.length,
        isLinkedWorkbook: token.value.includes('[')
    };
}

function annotateArrayContext(tokens: TokenMeta[], literals: HardCodedLiteral[]): void {
    const arrayRanges: { start: number; end: number }[] = [];
    const stack: number[] = [];

    tokens.forEach((token, idx) => {
        if (token.type === 'function' && token.value === 'ARRAY' && token.subtype === 'start') {
            stack.push(idx);
        } else if (token.type === 'function' && token.subtype === 'stop' && stack.length) {
            const startIdx = stack.pop();
            if (startIdx !== undefined) {
                arrayRanges.push({ start: startIdx, end: idx });
            }
        }
    });

    if (arrayRanges.length === 0) {
        return;
    }

    literals.forEach(literal => {
        const tokenIndex = findTokenIndex(tokens, literal.index);
        if (tokenIndex === -1) {
            return;
        }
        const inArray = arrayRanges.some(range => tokenIndex >= range.start && tokenIndex <= range.end);
        if (inArray) {
            literal.kind = literal.kind === 'numeric' ? 'array' : literal.kind;
            literal.containsMixedTypes = true;
        }
    });
}

function findTokenIndex(tokens: TokenMeta[], literalIndex: number): number {
    for (let i = 0; i < tokens.length; i++) {
        const token = tokens[i];
        if (literalIndex >= token.index && literalIndex < token.index + Math.max(token.length, 1)) {
            return i;
        }
    }
    return -1;
}

function classifyNumeric(
    value: string,
    parentFunction: string | null,
    parentArgIndex: number,
    tokens: TokenMeta[],
    tokenIndex: number
): {
    kind: LiteralKind;
    display: string;
    absoluteValue: number;
    isLikelyInput: boolean;
    containsMixedTypes: boolean;
    rationale?: string;
    suggestedFix?: string;
} {
    const trimmed = value.trim();
    const nextToken = tokens[tokenIndex + 1];

    if (nextToken && nextToken.type === 'operator-postfix' && nextToken.value === '%') {
        const numericValue = Number(trimmed);
        return {
            kind: 'percentage',
            display: `${trimmed}%`,
            absoluteValue: Math.abs(numericValue),
            isLikelyInput: true,
            containsMixedTypes: false,
            rationale: 'Percentage literal embedded in formula.',
            suggestedFix: 'Consider referencing a named percentage input instead of hard-coding.'
        };
    }

    const numberValue = Number(trimmed);
    const isInteger = Number.isInteger(numberValue);
    const absoluteValue = Math.abs(numberValue);

    let isLikelyInput = absoluteValue >= 1;
    let rationale: string | undefined;
    let suggestedFix: string | undefined;

    if (parentFunction) {
        const upper = parentFunction.toUpperCase();
        if (DATE_FUNCTIONS.has(upper)) {
            if (parentArgIndex === 0 && trimmed.length === 4) {
                isLikelyInput = true;
                rationale = 'Year literal supplied to date function.';
            } else if (parentArgIndex <= 2 && isInteger) {
                isLikelyInput = true;
                rationale = 'Day or month literal supplied to date function.';
            }
        }
        if (TIME_FUNCTIONS.has(upper) && isInteger && parentArgIndex <= 2) {
            isLikelyInput = true;
            rationale = 'Time component literal supplied to time function.';
        }
    }

    return {
        kind: 'numeric',
        display: trimmed,
        absoluteValue,
        isLikelyInput,
        containsMixedTypes: false,
        rationale,
        suggestedFix
    };
}

function isBenignLiteral(literal: HardCodedLiteral): boolean {
    if (literal.kind === 'boolean' && literal.isFlag) {
        return true;
    }
    if (literal.kind === 'named-literal' && literal.isRecognizedConstant) {
        return true;
    }
    if (literal.kind === 'string' && literal.isFlag) {
        return true;
    }
    return false;
}

function isExternalReference(value: string): boolean {
    return /!/.test(value);
}

function fallbackDetection(expression: string): HardCodedLiteral[] {
    const literals: HardCodedLiteral[] = [];
    const numberPattern = /-?\d+(?:\.\d+)?/g;
    const stringPattern = /"([^"\r\n]*)"/g;

    let match: RegExpExecArray | null;
    while ((match = numberPattern.exec(expression)) !== null) {
        literals.push({
            kind: 'numeric',
            value: match[0],
            display: match[0],
            index: match.index,
            length: match[0].length
        });
    }
    while ((match = stringPattern.exec(expression)) !== null) {
        literals.push({
            kind: 'string',
            value: match[1],
            display: `"${match[1]}"`,
            index: match.index,
            length: match[0].length
        });
    }

    return literals;
}

