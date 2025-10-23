import { detectHardCodedLiterals } from '../src/shared/HardCodeDetector';

const scenarios: Record<string, string[]> = {
    numericConstants: [
        '=SUM(A1,5,ROUND(B1,2))',
        '=A1*1.05',
        '=IF(B1>10,1000,0)',
        '=DATE(2024,10,1)'
    ],
    arrayConstants: [
        '={1,2,3}',
        '={"A","B","C"}',
        '=MMULT({1,2;3,4},A1:B2)'
    ],
    stringsAndBooleans: [
        '=IF(C1="USD",Rate,1)',
        '=IF(C1="Y",1,0)',
        '=AND(TRUE,A1>0)'
    ],
    percentagesAndTime: [
        '=A1*10%',
        '=TIME(12,30,0)',
        '=IF(B1>0.25,25%,B1)'
    ]
};

function main(): void {
    Object.entries(scenarios).forEach(([category, formulas]) => {
        console.log(`\n=== ${category} ===`);
        formulas.forEach(formula => {
            const literals = detectHardCodedLiterals(formula);
            console.log(`Formula: ${formula}`);
            literals.forEach(literal => {
                console.log('  ->', {
                    kind: literal.kind,
                    value: literal.display,
                    severityHint: literal.rationale
                });
            });
            if (literals.length === 0) {
                console.log('  (no literals detected)');
            }
        });
    });
}

main();
