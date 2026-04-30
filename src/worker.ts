import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

type ParsedRow = {
  number: string;
  name: string;
  cue: string;
};

type SyncResult = {
  added: string[];
  updated: string[];
  removed: string[];
};

type ServiceAccount = {
  client_email: string;
  private_key: string;
  token_uri?: string;
};

const COOKIE_NAME = "poc_session";
const HEADERS = ["", "Name", "Promo Number", "Promo Name", "Cue", "Notes", "Character"];
const SCOPES = "https://www.googleapis.com/auth/spreadsheets";

const HEADER_RE = /^\s*#?\s*(\d+)\s*[–-]\s*(.+)$/;
const VARIANT_RE = /^\s*(ALT\s+READ|Prelim\s+read|Main\s+Card\s+read)\s*$/i;
const PHONETIC_RE = /^\s*PHONETIC\s*[-–]/i;
const NOTE_RE = /^\s*SPOT\s*\(NOT\s+VIZ/i;

export default {
  async fetch(request: Request, env: Env, ctx: ExecutionContext): Promise<Response> {
    try {
      return await handleRequest(request, env, ctx);
    } catch (error) {
      if (error instanceof HttpError) {
        return json({ detail: error.message }, error.status);
      }
      console.error(JSON.stringify({
        event: "request_error",
        message: error instanceof Error ? error.message : String(error),
      }));
      return json({ detail: "Internal server error." }, 500);
    }
  },
};

async function handleRequest(request: Request, env: Env, _ctx: ExecutionContext): Promise<Response> {
  const url = new URL(request.url);
  const path = url.pathname;

  if (path.startsWith("/static/")) {
    return serveAsset(env, request, path.replace(/^\/static/, "") || "/index.html");
  }

  if (path === "/login" && request.method === "GET") {
    return serveAsset(env, request, "/login");
  }

  if (path === "/auth/login" && request.method === "POST") {
    return login(request, env);
  }

  if (path === "/auth/logout" && request.method === "GET") {
    return new Response(null, {
      status: 303,
      headers: {
        Location: "/login",
        "Set-Cookie": `${COOKIE_NAME}=; HttpOnly; Path=/; SameSite=Lax; Max-Age=0`,
      },
    });
  }

  if (!(await verifyToken(getCookie(request, COOKIE_NAME), env))) {
    return new Response(null, { status: 303, headers: { Location: "/login" } });
  }

  if (path === "/" && request.method === "GET") {
    return serveAsset(env, request, "/");
  }

  if (path === "/view" && request.method === "GET") {
    return serveAsset(env, request, "/view");
  }

  if (path === "/config-status" && request.method === "GET") {
    return json({ warnings: validateConfig(env) });
  }

  if (path === "/parse" && request.method === "POST") {
    const file = await getDocxFile(request);
    const rows = await parseDocx(await file.arrayBuffer());
    return json({ rows, count: rows.length });
  }

  if (path === "/upload" && request.method === "POST") {
    const file = await getDocxFile(request);
    const rows = await parseDocx(await file.arrayBuffer());
    const bytes = await writeXlsx(rows);
    const baseName = stripDocxExtension(file.name || "output");
    return xlsxResponse(bytes, `${baseName}_output.xlsx`);
  }

  if (path === "/download-rows" && request.method === "POST") {
    const payload = await request.json() as unknown;
    const rows = validateRows((payload as { rows?: unknown }).rows);
    const filename = sanitizeFilename(String((payload as { filename?: unknown }).filename || "output"));
    return xlsxResponse(await writeXlsx(rows), `${filename}_output.xlsx`);
  }

  if (path === "/sync" && request.method === "POST") {
    const form = await request.formData();
    const rowsRaw = String(form.get("rows") || "");
    const showType = String(form.get("show_type") || "").toUpperCase();
    const spreadsheetId = extractSpreadsheetId(String(form.get("spreadsheet_id") || ""));
    if (showType !== "PPV" && showType !== "FN") {
      return json({ detail: "show_type must be 'PPV' or 'FN'." }, 400);
    }
    if (!spreadsheetId) {
      return json({ detail: "Google Sheet target is required." }, 400);
    }

    let rows: ParsedRow[];
    try {
      rows = validateRows(JSON.parse(rowsRaw));
    } catch {
      return json({ detail: "Invalid rows JSON." }, 422);
    }

    const result = await syncRows(rows, spreadsheetId, env);
    return json({
      show_type: showType,
      added: result.added,
      updated: result.updated,
      removed: result.removed,
      total_rows: rows.length,
    });
  }

  return new Response("Not found", { status: 404 });
}

async function serveAsset(env: Env, request: Request, assetPath: string): Promise<Response> {
  const url = new URL(request.url);
  url.pathname = assetPath;
  url.search = "";
  return env.ASSETS.fetch(new Request(url.toString(), request));
}

async function login(request: Request, env: Env): Promise<Response> {
  const form = await request.formData();
  const username = String(form.get("username") || "");
  const password = String(form.get("password") || "");

  if (!(await verifyCredentials(username, password, env))) {
    return new Response(null, { status: 303, headers: { Location: "/login?error=1" } });
  }

  const token = await makeToken(username, env);
  return new Response(null, {
    status: 303,
    headers: {
      Location: "/",
      "Set-Cookie": `${COOKIE_NAME}=${token}; HttpOnly; Path=/; SameSite=Lax`,
    },
  });
}

async function verifyCredentials(username: string, password: string, env: Env): Promise<boolean> {
  const users = parseUsers(env);
  const stored = users[username];
  if (stored === undefined) {
    return false;
  }
  return timingSafeEqual(stored, password);
}

function parseUsers(env: Env): Record<string, string> {
  const raw = env.APP_USERS_JSON || "";
  if (!raw) {
    return {};
  }
  try {
    const parsed = JSON.parse(raw) as Record<string, unknown>;
    return Object.fromEntries(
      Object.entries(parsed).map(([key, value]) => [key, String(value)]),
    );
  } catch {
    return {};
  }
}

async function makeToken(username: string, env: Env): Promise<string> {
  const sig = await hmacHex(username, sessionSecret(env));
  return `${sig}.${username}`;
}

async function verifyToken(token: string | null, env: Env): Promise<string | null> {
  if (!token || !token.includes(".")) {
    return null;
  }
  const [sig, ...rest] = token.split(".");
  const username = rest.join(".");
  const expected = await hmacHex(username, sessionSecret(env));
  return (await timingSafeEqual(sig, expected)) ? username : null;
}

function sessionSecret(env: Env): string {
  return env.SESSION_SECRET || "local-development-secret-change-before-deploy";
}

async function hmacHex(message: string, secret: string): Promise<string> {
  const key = await crypto.subtle.importKey(
    "raw",
    utf8(secret),
    { name: "HMAC", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const sig = await crypto.subtle.sign("HMAC", key, utf8(message));
  return hex(sig);
}

async function timingSafeEqual(a: string, b: string): Promise<boolean> {
  const da = new Uint8Array(await crypto.subtle.digest("SHA-256", utf8(a)));
  const db = new Uint8Array(await crypto.subtle.digest("SHA-256", utf8(b)));
  let diff = da.length ^ db.length;
  for (let i = 0; i < Math.max(da.length, db.length); i += 1) {
    diff |= (da[i] || 0) ^ (db[i] || 0);
  }
  return diff === 0;
}

function getCookie(request: Request, name: string): string | null {
  const cookie = request.headers.get("Cookie") || "";
  for (const part of cookie.split(";")) {
    const [key, ...valueParts] = part.trim().split("=");
    if (key === name) {
      return valueParts.join("=");
    }
  }
  return null;
}

async function getDocxFile(request: Request): Promise<File> {
  const form = await request.formData();
  const value = form.get("file");
  if (!(value instanceof File) || !value.name.toLowerCase().endsWith(".docx")) {
    throw new HttpError("Only .docx files are supported.", 400);
  }
  return value;
}

async function parseDocx(buffer: ArrayBuffer): Promise<ParsedRow[]> {
  const zip = await JSZip.loadAsync(buffer);
  const documentXml = await zip.file("word/document.xml")?.async("text");
  if (!documentXml) {
    throw new HttpError("Failed to parse document: word/document.xml missing.", 422);
  }

  const parser = new XMLParser({
    ignoreAttributes: false,
    removeNSPrefix: true,
    trimValues: false,
  });
  const parsed = parser.parse(documentXml) as unknown;
  const paragraphs = asArray(getPath(parsed, ["document", "body", "p"]));
  return parseParagraphs(paragraphs.map(paragraphText));
}

function parseParagraphs(paragraphs: string[]): ParsedRow[] {
  const rows: ParsedRow[] = [];
  let currentNumber: string | null = null;
  let currentName: string | null = null;
  let cueLines: string[] = [];
  let variantCount = 0;
  let baseSlotTaken = false;
  let lastBaseNumber = 0;
  let blankRun = 0;
  let justSetHeader = false;

  const flush = (number: string, name: string, lines: string[]) => {
    const body: string[] = [];
    const phonetic: string[] = [];
    for (const raw of lines) {
      const line = raw.replace(/\u00a0/g, " ").replace(/^:\s+/, "");
      if (!line) {
        continue;
      }
      if (PHONETIC_RE.test(line)) {
        phonetic.push(line);
      } else {
        body.push(line);
      }
    }
    let cue = body.join(" ").toUpperCase();
    if (phonetic.length) {
      cue = `${cue}\n\n${phonetic.join(" ").toUpperCase()}`;
    }
    rows.push({ number, name, cue });
  };

  const startSection = (number: string, name: string) => {
    currentNumber = number;
    currentName = name.trim();
    cueLines = [];
    variantCount = 0;
    baseSlotTaken = false;
    justSetHeader = true;
    const parsedBase = Number.parseInt(number.split("-")[0] || "", 10);
    if (Number.isFinite(parsedBase)) {
      lastBaseNumber = parsedBase;
    }
  };

  for (const raw of paragraphs) {
    const text = raw.trim();
    if (!text) {
      blankRun += 1;
      justSetHeader = false;
      continue;
    }

    const previousBlankRun = blankRun;
    blankRun = 0;
    const header = HEADER_RE.exec(text);
    if (header) {
      if (currentNumber !== null && currentName !== null) {
        flush(currentNumber, currentName, cueLines);
      }
      startSection(header[1], header[2]);
      continue;
    }

    if (currentNumber === null || currentName === null) {
      continue;
    }

    if (previousBlankRun >= 5 && cueLines.length > 0 && text.length < 80 && !text.includes(".")) {
      flush(currentNumber, currentName, cueLines);
      startSection(String(lastBaseNumber + 1), text);
      continue;
    }

    if (justSetHeader && previousBlankRun === 0 && cueLines.length === 0 && NOTE_RE.test(text)) {
      currentName = `${currentName} ${text}`;
      continue;
    }

    const variant = VARIANT_RE.exec(text);
    if (variant) {
      const label = variant[1];
      const hasBaseCue = cueLines.length > 0;
      const base: string = currentNumber.split("-")[0] || "";
      if (hasBaseCue) {
        flush(currentNumber, currentName, cueLines);
        cueLines = [];
        variantCount += 1;
        const suffix = baseSlotTaken ? variantCount - 1 : variantCount;
        currentNumber = `${base}-${suffix}`;
      } else {
        variantCount += 1;
        if (variantCount === 1) {
          currentNumber = base;
          baseSlotTaken = true;
        } else {
          currentNumber = `${base}-${variantCount - 1}`;
        }
      }

      const baseTitle: string = currentName.split(/ - (ALT READ|Prelim|Main)$/)[0] || "";
      if (/prelim/i.test(label)) {
        currentName = `${baseTitle} - Prelim`;
      } else if (/main/i.test(label)) {
        currentName = `${baseTitle} - Main`;
      } else {
        currentName = `${baseTitle} - ALT READ`;
      }
      justSetHeader = false;
      continue;
    }

    cueLines.push(text);
    justSetHeader = false;
  }

  if (currentNumber !== null && currentName !== null) {
    flush(currentNumber, currentName, cueLines);
  }

  return rows;
}

function paragraphText(paragraph: unknown): string {
  const parts: string[] = [];
  collectText(paragraph, parts);
  return parts.join("");
}

function collectText(value: unknown, parts: string[]): void {
  if (value === null || value === undefined) {
    return;
  }
  if (typeof value === "string" || typeof value === "number") {
    return;
  }
  if (Array.isArray(value)) {
    for (const item of value) {
      collectText(item, parts);
    }
    return;
  }
  if (typeof value !== "object") {
    return;
  }

  const record = value as Record<string, unknown>;
  if (typeof record.t === "string" || typeof record.t === "number") {
    parts.push(String(record.t));
  } else if (record.t && typeof record.t === "object") {
    const textRecord = record.t as Record<string, unknown>;
    if (typeof textRecord["#text"] === "string" || typeof textRecord["#text"] === "number") {
      parts.push(String(textRecord["#text"]));
    }
  }
  if (record.tab !== undefined) {
    parts.push("\t");
  }
  for (const [key, child] of Object.entries(record)) {
    if (key.startsWith("@_") || key === "t" || key === "tab") {
      continue;
    }
    collectText(child, parts);
  }
}

async function writeXlsx(rows: ParsedRow[]): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file("[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`);
  zip.folder("_rels")?.file(".rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`);
  zip.folder("docProps")?.file("app.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>POC Sheet Generator</Application>
</Properties>`);
  zip.folder("docProps")?.file("core.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>POC Sheet Generator</dc:creator>
  <cp:lastModifiedBy>POC Sheet Generator</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>
</cp:coreProperties>`);

  const xl = zip.folder("xl");
  xl?.file("workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="Sheet1" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>`);
  xl?.file("styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>
  <borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="3">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`);
  xl?.folder("_rels")?.file("workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`);
  xl?.folder("worksheets")?.file("sheet1.xml", worksheetXml(rows));
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

function worksheetXml(rows: ParsedRow[]): string {
  const allRows = [
    HEADERS,
    ...rows.map((row) => [row.number, row.name, row.number, row.name, row.cue, "", ""]),
  ];
  const rowXml = allRows.map((values, index) => {
    const rowNumber = index + 1;
    const cells = values.map((value, colIndex) => {
      const ref = `${columnName(colIndex + 1)}${rowNumber}`;
      const style = rowNumber === 1 ? ` s="1"` : colIndex === 4 ? ` s="2"` : "";
      return inlineStringCell(ref, String(value), style);
    });
    if (rowNumber > 1) {
      cells[6] = `<c r="G${rowNumber}"><f>LEN(E${rowNumber})</f></c>`;
    }
    return `<row r="${rowNumber}">${cells.join("")}</row>`;
  }).join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cols>
    <col min="1" max="1" width="12" customWidth="1"/>
    <col min="2" max="2" width="40" customWidth="1"/>
    <col min="3" max="3" width="14" customWidth="1"/>
    <col min="4" max="4" width="40" customWidth="1"/>
    <col min="5" max="5" width="80" customWidth="1"/>
    <col min="6" max="6" width="20" customWidth="1"/>
    <col min="7" max="7" width="12" customWidth="1"/>
  </cols>
  <sheetData>${rowXml}</sheetData>
</worksheet>`;
}

function inlineStringCell(ref: string, value: string, style: string): string {
  return `<c r="${ref}" t="inlineStr"${style}><is><t xml:space="preserve">${xmlEscape(value)}</t></is></c>`;
}

function columnName(index: number): string {
  let n = index;
  let name = "";
  while (n > 0) {
    const remainder = (n - 1) % 26;
    name = String.fromCharCode(65 + remainder) + name;
    n = Math.floor((n - 1) / 26);
  }
  return name;
}

function xmlEscape(value: string): string {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

async function syncRows(rows: ParsedRow[], spreadsheetId: string, env: Env): Promise<SyncResult> {
  const token = await getGoogleAccessToken(env);
  const client = new SheetsClient(spreadsheetId, token);
  const sheetId = await client.getSheetId("Sheet1");

  let allValues = await client.getValues("Sheet1!A:G");
  if (!allValues.length) {
    await client.appendValues("Sheet1!A:G", [HEADERS]);
    allValues = [HEADERS];
  }

  const dataRows = allValues.slice(1);
  const existing = new Map<string, number>();
  dataRows.forEach((row, index) => {
    const num = normalizeNum(row[0] || "");
    if (num) {
      existing.set(num, index + 2);
    }
  });

  const target = new Map<string, ParsedRow>();
  rows.forEach((row) => target.set(normalizeNum(row.number), row));

  const added: string[] = [];
  const updated: string[] = [];
  const removed: string[] = [];

  const updateData: Array<{ range: string; values: string[][] }> = [];
  for (const [num, row] of target.entries()) {
    const sheetRow = existing.get(num);
    if (sheetRow !== undefined) {
      updateData.push({ range: `Sheet1!A${sheetRow}:G${sheetRow}`, values: [rowToValues(row, sheetRow)] });
      updated.push(row.name);
    }
  }
  if (updateData.length) {
    await client.batchUpdateValues(updateData);
  }

  const toDelete = Array.from(existing.entries())
    .filter(([num]) => !target.has(num))
    .map(([, rowIndex]) => rowIndex)
    .sort((a, b) => b - a);
  for (const rowIndex of toDelete) {
    const cached = dataRows[rowIndex - 2] || [];
    removed.push(cached[1] || "?");
  }
  if (toDelete.length) {
    await client.deleteRows(sheetId, toDelete);
  }

  const rowsToAppend = Array.from(target.entries())
    .filter(([num]) => !existing.has(num))
    .map(([, row]) => row);
  if (rowsToAppend.length) {
    const currentCount = allValues.length - toDelete.length;
    const appendValues = rowsToAppend.map((row, index) => {
      added.push(row.name);
      return rowToValues(row, currentCount + index + 1);
    });
    await client.appendValues("Sheet1!A:G", appendValues);
  }

  const totalRows = allValues.length - toDelete.length + rowsToAppend.length;
  if (totalRows >= 2) {
    await client.formatSheet(sheetId, totalRows);
  }

  return { added, updated, removed };
}

class SheetsClient {
  constructor(
    private readonly spreadsheetId: string,
    private readonly token: string,
  ) {}

  async getSheetId(title: string): Promise<number> {
    const data = await this.request<{ sheets?: Array<{ properties?: { sheetId?: number; title?: string } }> }>(
      `?fields=sheets(properties(sheetId,title))`,
    );
    const sheet = data.sheets?.find((item) => item.properties?.title === title);
    if (sheet?.properties?.sheetId === undefined) {
      throw new HttpError(`Worksheet '${title}' not found.`, 503);
    }
    return sheet.properties.sheetId;
  }

  async getValues(range: string): Promise<string[][]> {
    const data = await this.request<{ values?: string[][] }>(
      `/values/${encodeURIComponent(range)}?valueRenderOption=FORMATTED_VALUE`,
    );
    return data.values || [];
  }

  async appendValues(range: string, values: string[][]): Promise<void> {
    await this.request(
      `/values/${encodeURIComponent(range)}:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS`,
      { method: "POST", body: JSON.stringify({ values }) },
    );
  }

  async batchUpdateValues(data: Array<{ range: string; values: string[][] }>): Promise<void> {
    await this.request("/values:batchUpdate", {
      method: "POST",
      body: JSON.stringify({ valueInputOption: "USER_ENTERED", data }),
    });
  }

  async deleteRows(sheetId: number, rowIndices: number[]): Promise<void> {
    await this.batchUpdate(rowIndices.map((rowIndex) => ({
      deleteDimension: {
        range: {
          sheetId,
          dimension: "ROWS",
          startIndex: rowIndex - 1,
          endIndex: rowIndex,
        },
      },
    })));
  }

  async formatSheet(sheetId: number, totalRows: number): Promise<void> {
    await this.batchUpdate([
      {
        repeatCell: {
          range: { sheetId, startRowIndex: 1, endRowIndex: totalRows, startColumnIndex: 0, endColumnIndex: 7 },
          cell: { userEnteredFormat: { textFormat: { fontFamily: "Arial", fontSize: 10, bold: false } } },
          fields: "userEnteredFormat.textFormat",
        },
      },
      ...["A", "C"].map((col) => columnFormat(sheetId, col, totalRows, {
        horizontalAlignment: "CENTER",
        verticalAlignment: "MIDDLE",
        textFormat: { fontFamily: "Georgia", fontSize: 10, bold: false },
      })),
      columnFormat(sheetId, "F", totalRows, {
        horizontalAlignment: "CENTER",
        verticalAlignment: "MIDDLE",
      }),
      ...["B", "D", "G"].map((col) => columnFormat(sheetId, col, totalRows, {
        horizontalAlignment: "CENTER",
        verticalAlignment: "MIDDLE",
        textFormat: { fontFamily: "Arial", fontSize: 10, bold: true },
      })),
      columnFormat(sheetId, "E", totalRows, {
        horizontalAlignment: "LEFT",
        verticalAlignment: "TOP",
        wrapStrategy: "WRAP",
      }),
      {
        autoResizeDimensions: {
          dimensions: { sheetId, dimension: "ROWS", startIndex: 1, endIndex: totalRows },
        },
      },
      {
        autoResizeDimensions: {
          dimensions: { sheetId, dimension: "COLUMNS", startIndex: 1, endIndex: 2 },
        },
      },
      {
        autoResizeDimensions: {
          dimensions: { sheetId, dimension: "COLUMNS", startIndex: 3, endIndex: 4 },
        },
      },
    ]);
  }

  private async batchUpdate(requests: unknown[]): Promise<void> {
    await this.request(":batchUpdate", { method: "POST", body: JSON.stringify({ requests }) });
  }

  private async request<T = unknown>(path: string, init: RequestInit = {}): Promise<T> {
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${this.spreadsheetId}${path}`;
    const response = await fetch(url, {
      ...init,
      headers: {
        Authorization: `Bearer ${this.token}`,
        "Content-Type": "application/json",
        ...(init.headers || {}),
      },
    });
    if (!response.ok) {
      const detail = await response.text();
      throw new HttpError(`Google Sheets error: ${response.status} ${detail}`, 500);
    }
    return response.json() as Promise<T>;
  }
}

function columnFormat(sheetId: number, column: string, totalRows: number, format: Record<string, unknown>): unknown {
  const index = column.charCodeAt(0) - "A".charCodeAt(0);
  return {
    repeatCell: {
      range: { sheetId, startRowIndex: 1, endRowIndex: totalRows, startColumnIndex: index, endColumnIndex: index + 1 },
      cell: { userEnteredFormat: format },
      fields: Object.keys(format).map((key) => `userEnteredFormat.${key}`).join(","),
    },
  };
}

async function getGoogleAccessToken(env: Env): Promise<string> {
  if (!env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    throw new HttpError("Google service account secret is not configured.", 503);
  }

  let account: ServiceAccount;
  try {
    account = JSON.parse(env.GOOGLE_SERVICE_ACCOUNT_JSON) as ServiceAccount;
  } catch {
    throw new HttpError("Google service account secret is invalid JSON.", 503);
  }

  const now = Math.floor(Date.now() / 1000);
  const header = { alg: "RS256", typ: "JWT" };
  const payload = {
    iss: account.client_email,
    scope: SCOPES,
    aud: account.token_uri || "https://oauth2.googleapis.com/token",
    exp: now + 3600,
    iat: now,
  };
  const signingInput = `${base64urlJson(header)}.${base64urlJson(payload)}`;
  const signature = await signRs256(signingInput, account.private_key);
  const assertion = `${signingInput}.${signature}`;

  const response = await fetch(payload.aud, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion,
    }).toString(),
  });
  if (!response.ok) {
    throw new HttpError(`Google auth error: ${response.status} ${await response.text()}`, 503);
  }
  const data = await response.json() as { access_token?: string };
  if (!data.access_token) {
    throw new HttpError("Google auth response did not include an access token.", 503);
  }
  return data.access_token;
}

async function signRs256(message: string, pem: string): Promise<string> {
  const keyData = pemToArrayBuffer(pem);
  const key = await crypto.subtle.importKey(
    "pkcs8",
    keyData,
    { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" },
    false,
    ["sign"],
  );
  const signature = await crypto.subtle.sign("RSASSA-PKCS1-v1_5", key, utf8(message));
  return base64urlBytes(new Uint8Array(signature));
}

function pemToArrayBuffer(pem: string): ArrayBuffer {
  const base64 = pem
    .replace(/-----BEGIN PRIVATE KEY-----/g, "")
    .replace(/-----END PRIVATE KEY-----/g, "")
    .replace(/\s/g, "");
  const binary = atob(base64);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes.buffer;
}

function rowToValues(row: ParsedRow, rowIndex: number): string[] {
  return [row.number, row.name, row.number, row.name, row.cue, "", `=LEN(E${rowIndex})`];
}

function normalizeNum(value: unknown): string {
  const text = String(value).trim();
  const number = Number(text);
  if (Number.isFinite(number) && number === Math.trunc(number)) {
    return String(number);
  }
  return text;
}

function validateRows(value: unknown): ParsedRow[] {
  if (!Array.isArray(value)) {
    throw new HttpError("Invalid rows payload.", 422);
  }
  return value.map((row) => {
    if (!row || typeof row !== "object") {
      throw new HttpError("Invalid rows payload.", 422);
    }
    const record = row as Record<string, unknown>;
    return {
      number: String(record.number || ""),
      name: String(record.name || ""),
      cue: String(record.cue || ""),
    };
  });
}

function validateConfig(env: Env): string[] {
  const warnings: string[] = [];
  if (!env.GOOGLE_SERVICE_ACCOUNT_JSON) {
    warnings.push("Google Sheets: GOOGLE_SERVICE_ACCOUNT_JSON secret not configured");
  }
  if (!env.APP_USERS_JSON) {
    warnings.push("Auth: APP_USERS_JSON secret not configured");
  }
  if (!env.SESSION_SECRET) {
    warnings.push("Auth: SESSION_SECRET secret not configured");
  }
  return warnings;
}

function extractSpreadsheetId(value: string): string {
  const text = value.trim();
  if (!text) {
    return "";
  }
  const match = text.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
  const candidate = match?.[1] || text;
  return /^[a-zA-Z0-9_-]{20,}$/.test(candidate) ? candidate : "";
}

function xlsxResponse(bytes: Uint8Array, filename: string): Response {
  return new Response(toArrayBuffer(bytes), {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename="${sanitizeFilename(filename)}"`,
    },
  });
}

function json(data: unknown, status = 200): Response {
  return new Response(JSON.stringify(data), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function asArray(value: unknown): unknown[] {
  if (Array.isArray(value)) {
    return value;
  }
  return value === undefined || value === null ? [] : [value];
}

function getPath(value: unknown, path: string[]): unknown {
  let current = value;
  for (const key of path) {
    if (!current || typeof current !== "object") {
      return undefined;
    }
    current = (current as Record<string, unknown>)[key];
  }
  return current;
}

function stripDocxExtension(filename: string): string {
  return sanitizeFilename(filename.replace(/\.docx$/i, ""));
}

function sanitizeFilename(filename: string): string {
  return filename.replace(/[^\w .-]/g, "_").trim() || "output";
}

function base64urlJson(value: unknown): string {
  return base64urlBytes(new Uint8Array(utf8(JSON.stringify(value))));
}

function base64urlBytes(bytes: ArrayLike<number>): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 1) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/g, "");
}

function utf8(value: string): ArrayBuffer {
  return toArrayBuffer(new TextEncoder().encode(value));
}

function toArrayBuffer(bytes: Uint8Array): ArrayBuffer {
  const copy = new Uint8Array(bytes.byteLength);
  copy.set(bytes);
  return copy.buffer;
}

function hex(buffer: ArrayBuffer): string {
  return Array.from(new Uint8Array(buffer), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

class HttpError extends Error {
  constructor(message: string, readonly status: number) {
    super(message);
  }
}
