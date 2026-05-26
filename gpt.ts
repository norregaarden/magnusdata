// gpt.ts
import path from "node:path";
import { readdir, readFile, writeFile, stat } from "node:fs/promises";
import { join, basename, extname } from "node:path";
import { createReadStream } from "node:fs";
import { Readable } from "node:stream";
import { unzipSync } from "node:zlib"; // built-in, for .zip support

//
// Types
//

export type TimeString = `${number}:${number}` | `${number}:${number}:${number}`;

export interface ActivityRow {
  activity: string;
  description: string;
  start: TimeString;
  end: TimeString;
  duration: TimeString;
  pixels: number;

  startSeconds: number;
  endSeconds: number;
  durationSeconds: number;
}

export interface Session {
  name: string;
  file: string;
  activities: ActivityRow[];
}

//
// Robust CSV tokenizer
// Handles: quoted fields, commas inside quotes, escaped quotes ("")
//

function parseCsvLine(line: string): string[] {
  const fields: string[] = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];
    const next = line[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        // Escaped quote inside quoted field
        current += '"';
        i++; // skip next
      } else {
        inQuotes = !inQuotes;
      }
    } else if (char === ',' && !inQuotes) {
      fields.push(current.trim());
      current = "";
    } else {
      current += char;
    }
  }

  fields.push(current.trim());
  return fields;
}

//
// Time parser: MM:SS or HH:MM:SS
//

function parseTimeToSeconds(time: string): number {
  const t = time.trim();
  if (!t) return 0;

  const parts = t.split(":").map(Number);

  if (parts.some(Number.isNaN)) {
    throw new Error(`Invalid time format (non-numeric): "${time}"`);
  }

  if (parts.length === 2) {
    const [m, s] = parts;
    return m * 60 + s;
  }

  if (parts.length === 3) {
    const [h, m, s] = parts;
    return h * 3600 + m * 60 + s;
  }

  throw new Error(`Invalid time format: "${time}" (expected MM:SS or HH:MM:SS)`);
}

//
// Parse one CSV text into a Session
//

export function parseSessionCsv(csvText: string, fileName: string): Session {
  const lines = csvText
    .split(/\r?\n/)
    .map(x => x.trim())
    .filter(Boolean);

  if (lines.length < 3) {
    throw new Error(`Too few lines in ${fileName}: ${lines.length} lines (need at least 3)`);
  }

  // Row 0: session name
  const firstRow = parseCsvLine(lines[0]);
  const sessionName = firstRow[0] || basename(fileName, extname(fileName));

  // Row 1: headers (skip)
  // Row 2+: data
  const rows: ActivityRow[] = [];

  for (let i = 2; i < lines.length; i++) {
    const line = lines[i];
    const cols = parseCsvLine(line);

    // Defensive: skip completely empty rows
    if (cols.every(c => !c)) continue;

    // Pad to 6 columns if short
    while (cols.length < 6) cols.push("");

    const [
      activity = "",
      description = "",
      start = "00:00",
      end = "00:00",
      duration = "00:00",
      pixels = "0",
    ] = cols;

    rows.push({
      activity,
      description,
      start: start as TimeString,
      end: end as TimeString,
      duration: duration as TimeString,
      pixels: Number(pixels) || 0,

      startSeconds: parseTimeToSeconds(start),
      endSeconds: parseTimeToSeconds(end),
      durationSeconds: parseTimeToSeconds(duration),
    });
  }

  return {
    name: sessionName,
    file: fileName,
    activities: rows,
  };
}

//
// File discovery: recursive, with .zip extraction
//

async function findCsvFiles(inputPath: string): Promise<{ path: string; text: string }[]> {
  const s = await stat(inputPath);

  // Single CSV file
  if (s.isFile() && inputPath.endsWith(".csv")) {
    const text = await readFile(inputPath, "utf8");
    return [{ path: inputPath, text }];
  }

  // Zip file: extract in-memory and find CSVs inside
  if (s.isFile() && inputPath.endsWith(".zip")) {
    const zipBuffer = await readFile(inputPath);
    const extracted = unzipSync(zipBuffer);
    const results: { path: string; text: string }[] = [];

    for (const [name, buffer] of Object.entries(extracted)) {
      if (name.endsWith(".csv")) {
        results.push({ path: `${inputPath}#${name}`, text: buffer.toString("utf8") });
      }
    }

    if (results.length === 0) {
      throw new Error(`No CSV files found inside zip: ${inputPath}`);
    }

    return results;
  }

  // Directory: recursive scan
  if (s.isDirectory()) {
    const entries = await readdir(inputPath, { withFileTypes: true });
    const nested = await Promise.all(
      entries.map(async entry => {
        const full = join(inputPath, entry.name);
        if (entry.isDirectory()) {
          return findCsvFiles(full);
        }
        if (entry.isFile() && full.endsWith(".csv")) {
          const text = await readFile(full, "utf8");
          return [{ path: full, text }];
        }
        if (entry.isFile() && full.endsWith(".zip")) {
          return findCsvFiles(full);
        }
        return [];
      })
    );
    return nested.flat();
  }

  throw new Error(`Not a file, directory, or zip: ${inputPath}`);
}

//
// Path normalization: Windows C:\ -> /mnt/c/
//

function normalizeInputPath(input: string): string {
  const winMatch = input.match(/^([A-Za-z]):\\(.*)$/);
  if (winMatch) {
    const [, drive, rest] = winMatch;
    return `/mnt/${drive.toLowerCase()}/${rest.replaceAll("\\", "/")}`;
  }
  return input;
}

//
// Main
//

async function main() {
  const rawInput = process.argv[2] ?? ".";
  const root = normalizeInputPath(rawInput);

  console.error(`Scanning: ${root}`);

  const sources = await findCsvFiles(root);
  console.error(`Found ${sources.length} CSV source(s)`);

  const sessions: Session[] = [];

  for (const { path: filePath, text } of sources) {
    try {
      const session = parseSessionCsv(text, filePath);
      sessions.push(session);
      console.error(`  ✓ ${filePath} → ${session.activities.length} activities`);
    } catch (e) {
      console.error(`  ✗ ${filePath}: ${e instanceof Error ? e.message : String(e)}`);
      // Continue processing other files; don't crash the whole batch
    }
  }

  const outPath = "sessions.json";
  await writeFile(
    outPath,
    JSON.stringify(sessions, null, 2),
  );

  console.error(`\nWrote ${sessions.length} session(s) to ${outPath}`);
}

main().catch(error => {
  console.error("Fatal error:", error);
  process.exit(1);
});
