import path from "node:path";
import { readdir, readFile, writeFile } from "node:fs/promises";
import { join } from "node:path";

//
// Types
//

export type TimeString = `${number}:${number}`;

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
// Helpers
//

function parseTimeToSeconds(time: string): number {
  const parts = time.trim().split(":").map(Number);

  if (parts.length === 2) {
    const [minutes, seconds] = parts;
    return minutes * 60 + seconds;
  }

  if (parts.length === 3) {
    const [hours, minutes, seconds] = parts;
    return hours * 3600 + minutes * 60 + seconds;
  }

  throw new Error(`Invalid time format: ${time}`);
}

function splitCsvLine(line: string): string[] {
  return line.split(",").map(x => x.trim());
}

//
// Parse one CSV
//

export function parseSessionCsv(
  csvText: string,
  fileName: string,
): Session {

  const lines = csvText
    .split(/\r?\n/)
    .map(x => x.trim())
    .filter(Boolean);

  if (lines.length < 3) {
    throw new Error(`Too few lines in ${fileName}`);
  }

  const firstRow = splitCsvLine(lines[0]);

  const sessionName =
    firstRow[0] ||
    fileName.replace(/\.csv$/i, "");

  const rows: ActivityRow[] = [];

  for (const line of lines.slice(2)) {

    const [
      activity = "",
      description = "",
      start = "00:00",
      end = "00:00",
      duration = "00:00",
      pixels = "0",
    ] = splitCsvLine(line);

    rows.push({
      activity,
      description,
      start: start as TimeString,
      end: end as TimeString,
      duration: duration as TimeString,
      pixels: Number(pixels),

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
// Recursive file discovery
//

async function findCsvFiles(dir: string): Promise<string[]> {

  const entries = await readdir(dir, {
    withFileTypes: true,
  });

  const files = await Promise.all(
    entries.map(async entry => {

      const full = join(dir, entry.name);

      if (entry.isDirectory()) {
        return findCsvFiles(full);
      }

      if (entry.isFile() && full.endsWith(".csv")) {
        return [full];
      }

      return [];
    }),
  );

  return files.flat();
}

//
// Main
//

async function main() {


  function normalizeInputPath(input: string): string {
    // C:\Users\foo\bar
    const winMatch = input.match(/^([A-Za-z]):\\(.*)$/);

    if (winMatch) {
      const [, drive, rest] = winMatch;

      return `/mnt/${drive.toLowerCase()}/${rest.replaceAll("\\", "/")}`;
    }

    return input;
  }


  const _root =
    process.argv[2] ??
    ".";
  const root = normalizeInputPath(
    process.argv[2] ?? ".",
  );

  const csvFiles = await findCsvFiles(root);

  const sessions: Session[] = [];

  for (const file of csvFiles) {

    const text = await readFile(file, "utf8");

    sessions.push(
      parseSessionCsv(text, file),
    );
  }

  await writeFile(
    "sessions.json",
    JSON.stringify(sessions, null, 2),
  );

  console.log(
    `Parsed ${sessions.length} session(s)`
  );
}

main().catch(error => {
  console.error(error);
  process.exit(1);
});
