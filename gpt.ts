// gpt.ts
import path from "node:path";
import { readdir, readFile, writeFile, stat } from "node:fs/promises";
import { join, basename, extname } from "node:path";
import { unzipSync } from "node:zlib";
import * as XLSX from "xlsx";

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

export interface Partition {
  by: string;
  key: string;
  totalSeconds: number;
  totalPixels: number;
  count: number;
  proportionSeconds: number;
  proportionPixels: number;
}

export interface SessionAnalysis {
  session: Session;
  totalSeconds: number;
  totalPixels: number;
  partitions: Map<string, Partition[]>;
}

export interface GlobalAnalysis {
  sessions: SessionAnalysis[];
  grandTotalSeconds: number;
  grandTotalPixels: number;
  globalPartitions: Map<string, Partition[]>;
}

//
// Time parser: MM:SS, HH:MM:SS, Excel serial numbers, "0", ""
//

export function parseTimeToSeconds(time: string | number): number {
  if (time === "" || time === null || time === undefined) return 0;

  if (typeof time === "number") {
    return Math.round(time * 24 * 60 * 60);
  }

  const t = String(time).trim();
  if (!t || t === "0") return 0;

  const asNum = Number(t);
  if (!Number.isNaN(asNum) && asNum > 0 && asNum < 1 && !t.includes(":")) {
    return Math.round(asNum * 24 * 60 * 60);
  }

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
// Robust CSV tokenizer (RFC-4180-ish)
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
        current += '"';
        i++;
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
// Parse CSV text into Session — THE ONE AND ONLY PARSER
//

export function parseSessionCsv(csvText: string, fileName: string): Session {
  const lines = csvText
    .split(/\r?\n/)
    .map(x => x.trim())
    .filter(Boolean);

  if (lines.length < 3) {
    throw new Error(`Too few lines in ${fileName}: ${lines.length} lines (need at least 3)`);
  }

  const firstRow = parseCsvLine(lines[0]);
  const sessionName = firstRow[0] || basename(fileName, extname(fileName));
  const rows: ActivityRow[] = [];

  for (let i = 2; i < lines.length; i++) {
    const cols = parseCsvLine(lines[i]);
    if (cols.every(c => !c)) continue;
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

  return { name: sessionName, file: fileName, activities: rows };
}

//
// XLSX → CSV text conversion
//

function xlsxToCsvText(buffer: Buffer): string {
  const workbook = XLSX.read(buffer, { type: "buffer" });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  return XLSX.utils.sheet_to_csv(sheet);
}

//
// Path normalization: Windows C:\ → /mnt/c/
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
// File discovery: recursive, case-insensitive extensions
//

async function findSources(inputPath: string): Promise<{ path: string; text: string }[]> {
  const s = await stat(inputPath);

  if (s.isFile()) {
    const ext = extname(inputPath).toLowerCase();
    if (ext === ".csv") {
      return [{ path: inputPath, text: await readFile(inputPath, "utf8") }];
    }
    if (ext === ".xlsx" || ext === ".xls") {
      const buf = await readFile(inputPath);
      return [{ path: inputPath, text: xlsxToCsvText(buf) }];
    }
    if (ext === ".zip") {
      const zipBuffer = await readFile(inputPath);
      const extracted = unzipSync(zipBuffer);
      const results: { path: string; text: string }[] = [];
      for (const [name, buffer] of Object.entries(extracted)) {
        const innerExt = extname(name).toLowerCase();
        if (innerExt === ".csv") {
          results.push({ path: `${inputPath}#${name}`, text: buffer.toString("utf8") });
        }
        if (innerExt === ".xlsx" || innerExt === ".xls") {
          results.push({ path: `${inputPath}#${name}`, text: xlsxToCsvText(buffer) });
        }
      }
      if (results.length === 0) throw new Error(`No CSV/XLSX files found inside zip: ${inputPath}`);
      return results;
    }
    throw new Error(`Unsupported file type: ${inputPath}`);
  }

  if (s.isDirectory()) {
    const entries = await readdir(inputPath, { withFileTypes: true });
    const nested = await Promise.all(
      entries.map(async entry => {
        const full = join(inputPath, entry.name);
        if (entry.isDirectory()) return findSources(full);
        const ext = extname(full).toLowerCase();
        if (entry.isFile() && (ext === ".csv" || ext === ".xlsx" || ext === ".xls" || ext === ".zip")) {
          return findSources(full);
        }
        return [];
      })
    );
    return nested.flat();
  }

  throw new Error(`Not a file or directory: ${inputPath}`);
}

//
// Auto-discovery: scan current directory (non-recursive)
//

async function autoFindSources(): Promise<{ path: string; text: string }[]> {
  const cwd = process.cwd();
  const entries = await readdir(cwd, { withFileTypes: true });

  const pending: Promise<{ path: string; text: string }[]>[] = [];

  for (const entry of entries) {
    if (!entry.isFile()) continue;
    const full = join(cwd, entry.name);
    const ext = extname(full).toLowerCase();
    if (ext === ".csv" || ext === ".xlsx" || ext === ".xls" || ext === ".zip") {
      pending.push(findSources(full));
    }
  }

  const results = await Promise.all(pending);
  return results.flat();
}

//
// Dependent partition engine
//

function getPartitionableFields(row: ActivityRow): string[] {
  return Object.entries(row)
    .filter(([_, v]) => typeof v === "string" && v.length > 0)
    .map(([k, _]) => k);
}

function makePartitions(
  rows: ActivityRow[],
  fieldName: string,
  totalSeconds: number,
  totalPixels: number,
): Partition[] {
  const groups = new Map<string, { seconds: number; pixels: number; count: number }>();

  for (const row of rows) {
    const key = String((row as any)[fieldName] ?? "");
    if (!key) continue;
    const g = groups.get(key) ?? { seconds: 0, pixels: 0, count: 0 };
    g.seconds += row.durationSeconds;
    g.pixels += row.pixels;
    g.count += 1;
    groups.set(key, g);
  }

  const parts: Partition[] = [];
  for (const [key, g] of groups) {
    parts.push({
      by: fieldName,
      key,
      totalSeconds: g.seconds,
      totalPixels: g.pixels,
      count: g.count,
      proportionSeconds: totalSeconds > 0 ? g.seconds / totalSeconds : 0,
      proportionPixels: totalPixels > 0 ? g.pixels / totalPixels : 0,
    });
  }

  return parts.sort((a, b) => b.proportionSeconds - a.proportionSeconds);
}

function analyzeSession(session: Session): SessionAnalysis {
  const totalSeconds = session.activities.reduce((s, r) => s + r.durationSeconds, 0);
  const totalPixels = session.activities.reduce((s, r) => s + r.pixels, 0);

  const fields = session.activities.length > 0
    ? getPartitionableFields(session.activities[0])
    : [];

  const partitions = new Map<string, Partition[]>();
  for (const field of fields) {
    partitions.set(field, makePartitions(session.activities, field, totalSeconds, totalPixels));
  }

  return { session, totalSeconds, totalPixels, partitions };
}

function analyzeGlobal(sessionAnalyses: SessionAnalysis[]): GlobalAnalysis {
  const grandTotalSeconds = sessionAnalyses.reduce((s, a) => s + a.totalSeconds, 0);
  const grandTotalPixels = sessionAnalyses.reduce((s, a) => s + a.totalPixels, 0);
  const allRows = sessionAnalyses.flatMap(a => a.session.activities);

  const allFields = new Set<string>();
  for (const sa of sessionAnalyses) {
    for (const f of sa.partitions.keys()) allFields.add(f);
  }

  const globalPartitions = new Map<string, Partition[]>();
  for (const field of allFields) {
    globalPartitions.set(field, makePartitions(allRows, field, grandTotalSeconds, grandTotalPixels));
  }

  return { sessions: sessionAnalyses, grandTotalSeconds, grandTotalPixels, globalPartitions };
}

//
// Formatters
//

function fmtTime(sec: number): string {
  const h = Math.floor(sec / 3600);
  const m = Math.floor((sec % 3600) / 60);
  const s = sec % 60;
  if (h > 0) return `${h}:${String(m).padStart(2, "0")}:${String(s).padStart(2, "0")}`;
  return `${m}:${String(s).padStart(2, "0")}`;
}

function fmtPct(n: number): string {
  return `${(n * 100).toFixed(1)}%`;
}

function renderPartitionTable(parts: Partition[]): string {
  if (parts.length === 0) return "  (no data)";
  const maxKey = Math.max(8, ...parts.map(p => p.key.length));
  const header = `${"Key".padEnd(maxKey)} | Count |   Time   | Time% | Pixels | Pix%`;
  const lines = parts.map(p =>
    `${p.key.padEnd(maxKey)} | ${String(p.count).padStart(5)} | ${fmtTime(p.totalSeconds).padStart(8)} | ${fmtPct(p.proportionSeconds).padStart(5)} | ${String(p.totalPixels).padStart(6)} | ${fmtPct(p.proportionPixels).padStart(4)}`
  );
  return [header, "-".repeat(header.length), ...lines].join("\n");
}

function renderAnalysis(a: SessionAnalysis): string {
  const parts: string[] = [
    `\nSession: ${a.session.name} (${a.session.file})`,
    `Total: ${fmtTime(a.totalSeconds)} | ${a.totalPixels} pixels | ${a.session.activities.length} activities`,
  ];
  for (const [field, partitions] of a.partitions) {
    parts.push(`\nBy ${field}:`);
    parts.push(renderPartitionTable(partitions));
  }
  return parts.join("\n");
}

function renderGlobal(a: GlobalAnalysis): string {
  const parts: string[] = [
    `\n═══════════════════════════════════════════════════════════════`,
    `GLOBAL ANALYSIS`,
    `${a.sessions.length} sessions | ${fmtTime(a.grandTotalSeconds)} total | ${a.grandTotalPixels} pixels total`,
  ];
  for (const [field, partitions] of a.globalPartitions) {
    parts.push(`\nGlobal by ${field}:`);
    parts.push(renderPartitionTable(partitions));
  }
  parts.push(`═══════════════════════════════════════════════════════════════`);
  return parts.join("\n");
}

//
// Main
//

async function main() {
  const rawInput = process.argv[2];

  let sources: { path: string; text: string }[];

  if (rawInput) {
    const root = normalizeInputPath(rawInput);
    console.error(`Scanning: ${root}`);
    sources = await findSources(root);
  } else {
    console.error(`Auto-scanning: ${process.cwd()}`);
    sources = await autoFindSources();
  }

  console.error(`Found ${sources.length} source(s)`);

  const sessions: Session[] = [];

  for (const source of sources) {
    try {
      const session = parseSessionCsv(source.text, source.path);
      sessions.push(session);
      console.error(`  ✓ ${source.path} → ${session.activities.length} activities`);
    } catch (e) {
      console.error(`  ✗ ${source.path}: ${e instanceof Error ? e.message : String(e)}`);
    }
  }

  if (sessions.length === 0) {
    console.error("No sessions parsed successfully.");
    process.exit(1);
  }

  const sessionAnalyses = sessions.map(analyzeSession);
  const globalAnalysis = analyzeGlobal(sessionAnalyses);

  for (const sa of sessionAnalyses) {
    console.log(renderAnalysis(sa));
  }
  console.log(renderGlobal(globalAnalysis));

  const outPath = "sessions.json";
  await writeFile(
    outPath,
    JSON.stringify(
      { sessions, analysis: globalAnalysis },
      (key, value) => value instanceof Map ? Object.fromEntries(value) : value,
      2,
    ),
  );

  console.error(`\nWrote ${sessions.length} session(s) + analysis to ${outPath}`);
}

main().catch(error => {
  console.error("Fatal error:", error);
  process.exit(1);
});
