import { postProcessDescription } from './bomEnricher';

const testCases = [
    {
        input: "ATTEN CHIP DC-18GHz 3 DB",
        partNumber: "TS0503W3",
        expected: "ATTEN DC-18GHz 3DB"
    }
];

testCases.forEach(({ input, partNumber, expected }) => {
    const result = postProcessDescription(input, partNumber);
    console.log(`Input: "${input}"`);
    console.log(`Part Number: "${partNumber}"`);
    console.log(`Result: "${result}"`);
    console.log(`Expected: "${expected}"`);
    console.log(`Test ${result === expected ? 'PASSED' : 'FAILED'}\n`);
}); 