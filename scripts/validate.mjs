// Validates the extension layout the Camillia hub expects from a connector:
// extension.json metadata, connector.json manifest rules (same as the server-side
// scan), and the included files. Run with `node scripts/validate.mjs`.
import { existsSync, readFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const failures = [];
const fail = (message) => failures.push(message);

const KNOWN_KINDS = ["camflow", "skill", "rules-pack", "conventions-pack", "plugin", "workspace-template", "connector"];
const KNOWN_GRANTS = ["edit-task", "comment-task", "create-task"];
const KNOWN_TRIGGERS = ["webhook", "poll"];

function readJson(file) {
  try {
    return JSON.parse(readFileSync(join(root, file), "utf8"));
  } catch (error) {
    fail(`${file} is not readable JSON: ${error.message}`);
    return null;
  }
}

function isValidCronField(field, min, max) {
  return field.split(",").every(part => {
    const [base, step] = part.split("/");
    if (step !== undefined && (!/^\d+$/.test(step) || Number(step) < 1)) return false;
    if (base === "*") return true;
    const range = base.split("-");
    return range.every(bound => /^\d+$/.test(bound) && Number(bound) >= min && Number(bound) <= max);
  });
}

function isValidCron(expression) {
  const fields = (expression ?? "").trim().split(/\s+/);
  if (fields.length !== 5) return false;
  const ranges = [[0, 59], [0, 23], [1, 31], [1, 12], [0, 7]];
  return fields.every((field, i) => isValidCronField(field, ranges[i][0], ranges[i][1]));
}

// --- extension.json ---------------------------------------------------------------
const hub = readJson("extension.json");
if (hub) {
  if (!hub.name || !hub.slug) fail("extension.json needs name and slug.");
  if (!KNOWN_KINDS.includes(hub.kind)) fail(`extension.json kind '${hub.kind}' is unknown (known: ${KNOWN_KINDS.join(", ")}).`);
  if (hub.kind !== "connector") fail(`this repository template is a connector; extension.json says '${hub.kind}'.`);
  if (!Array.isArray(hub.include) || hub.include.length === 0) fail("extension.json include lists the published files.");
  for (const file of hub.include ?? []) {
    if (!existsSync(join(root, file))) fail(`extension.json includes '${file}' but the file does not exist.`);
  }
}

// --- connector.json ---------------------------------------------------------
const manifest = readJson("connector.json");
if (manifest) {
  const entry = (manifest.entry ?? "index.js").trim();
  if (entry.includes("..") || entry.startsWith("/")) fail(`entry '${entry}' is not a safe relative path.`);
  if (!existsSync(join(root, entry))) fail(`entry '${entry}' does not exist.`);
  if (!hub?.include?.includes(entry)) fail(`entry '${entry}' must be listed in extension.json include.`);

  const triggers = Array.isArray(manifest.triggers) ? manifest.triggers : [];
  if (triggers.length === 0) fail("a connector declares at least one trigger.");
  const seen = new Set();
  for (const trigger of triggers) {
    if (!KNOWN_TRIGGERS.includes(trigger?.type)) fail(`trigger type '${trigger?.type}' is unknown (known: ${KNOWN_TRIGGERS.join(", ")}).`);
    if (seen.has(trigger?.type)) fail(`trigger '${trigger?.type}' is declared twice.`);
    seen.add(trigger?.type);
    if (trigger?.type === "poll" && !isValidCron(trigger?.schedule)) fail(`poll trigger needs a valid 5-field cron schedule, got '${trigger?.schedule}'.`);
    if (trigger?.type === "webhook" && trigger?.schedule) fail("a webhook trigger takes no schedule.");
  }

  for (const grant of manifest.grants ?? []) {
    if (!KNOWN_GRANTS.includes(grant)) fail(`grant '${grant}' is unknown (known: ${KNOWN_GRANTS.join(", ")}).`);
  }

  const keys = new Set();
  for (const secret of manifest.secrets ?? []) {
    if (!/^[A-Z][A-Z0-9_]*$/.test(secret?.key ?? "")) fail(`secret key '${secret?.key}' is not UPPER_SNAKE_CASE.`);
    if (keys.has(secret?.key)) fail(`secret '${secret?.key}' is declared twice.`);
    keys.add(secret?.key);
  }
}

if (failures.length > 0) {
  console.error("Validation failed:");
  for (const failure of failures) console.error(`  - ${failure}`);
  process.exit(1);
}
console.log("connector-teams: manifest and layout are valid.");
