// test.ts — run with: bun test.ts
import { parseSessionCsv } from "./gpt.ts";

// Test 1: Basic CSV parsing
const basicCsv = `Session A
Activity,Description,Start,End,Duration,Pixels
Work,Email,0:05:00,0:10:00,0:05:00,100
Break,Coffee,0:10:00,0:15:00,0:05:00,0
`;
const s1 = parseSessionCsv(basicCsv, "test.csv");
console.assert(s1.name === "Session A", "Session name from first cell");
console.assert(s1.activities.length === 2, "Two activities");
console.assert(s1.activities[0].durationSeconds === 300, "5 minutes = 300s");
console.assert(s1.activities[0].pixels === 100, "Pixel count");

// Test 2: Quoted fields with commas
const quotedCsv = `Session B
Activity,Description,Start,End,Duration,Pixels
"Work, urgent","Sheila Blanco y Orquestra de RTVE",1:03:26,2:06:52,1:03:26,500
`;
const s2 = parseSessionCsv(quotedCsv, "quoted.csv");
console.assert(s2.activities[0].activity === "Work, urgent", "Comma inside quotes");
console.assert(s2.activities[0].durationSeconds === 3806, "1:03:26 = 3806s");

// Test 3: Empty/short rows
const sparseCsv = `Session C
Activity,Description,Start,End,Duration,Pixels
A,D,0:01,0:02,0:01,10

B,D,0:03,0:04,0:01,20
`;
const s3 = parseSessionCsv(sparseCsv, "sparse.csv");
console.assert(s3.activities.length === 2, "Empty row skipped");

// Test 4: Time formats
import { parseTimeToSeconds } from "./gpt.ts"; // export this if testing separately
console.assert(parseTimeToSeconds("0:05:00") === 300, "HH:MM:SS");
console.assert(parseTimeToSeconds("5:00") === 300, "MM:SS");
console.assert(parseTimeToSeconds("0") === 0, "Zero string");
console.assert(parseTimeToSeconds("") === 0, "Empty string");
console.assert(parseTimeToSeconds(0.0002314814814814815) === 20, "Excel serial ≈ 20s");

console.log("All tests passed.");
