import { postProcessDescription } from './bomEnricher';

const testCases = [
    {
        input: "ATTEN CHIP DC-18GHz 3 DB",
        partNumber: "TS0503W3",
        expected: "ATTEN DC-18GHz 3DB"
    },
    {
        input: "CAP 100NF 50V",
        partNumber: "C1234",
        expected: "CAP+CRM 100NF 50V"
    },
    {
        input: "CAP+CRM 100NF 50V",
        partNumber: "C1234",
        expected: "CAP+CRM 100NF 50V"
    },
    {
        input: "CAP CRM 100NF 50V",
        partNumber: "C1234",
        expected: "CAP CRM 100NF 50V"
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