import { NextRequest, NextResponse } from "next/server";
import JSZip from "jszip";

export const runtime = "nodejs";

const NS = {
  spreadsheet:
    "http://schemas.openxmlformats.org/spreadsheetml/2006/main",

  relationships:
    "http://schemas.openxmlformats.org/package/2006/relationships",

  officeRelationships:
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
};

interface SheetInfo {
  name: string;
  sheetId: string;
  relationshipId: string;
  relationshipTarget: string;
}

interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: string;
}

/**
 * POST /api/split
 *
 * Receives:
 *   multipart/form-data
 *   file = .xlsx / .xlsm
 *
 * Returns:
 *   ZIP containing one workbook per worksheet.
 */
export async function POST(request: NextRequest) {
  try {
    const formData = await request.formData();

    const uploaded = formData.get("file");

    if (!(uploaded instanceof File)) {
      return jsonError("No Excel file was uploaded.", 400);
    }

    const fileName = uploaded.name;

    const extension = getExtension(fileName);

    if (extension !== ".xlsx" && extension !== ".xlsm") {
      return jsonError(
        "Only .xlsx and .xlsm files are supported.",
        400
      );
    }

    const inputBuffer = Buffer.from(await uploaded.arrayBuffer());

    const sourceZip = await JSZip.loadAsync(inputBuffer);

    const workbookXmlFile = sourceZip.file("xl/workbook.xml");

    if (!workbookXmlFile) {
      return jsonError(
        "Invalid Excel file: xl/workbook.xml was not found.",
        400
      );
    }

    const workbookXml = await workbookXmlFile.async("string");

    const workbookRelsFile = sourceZip.file(
      "xl/_rels/workbook.xml.rels"
    );

    if (!workbookRelsFile) {
      return jsonError(
        "Invalid Excel file: workbook relationships were not found.",
        400
      );
    }

    const workbookRelsXml =
      await workbookRelsFile.async("string");

    const sheets = parseSheets(workbookXml);

    if (!sheets.length) {
      return jsonError(
        "The workbook does not contain worksheets.",
        400
      );
    }

    const workbookRelationships =
      parseRelationships(workbookRelsXml);

    /*
     * Resolve every worksheet relationship.
     */
    for (const sheet of sheets) {
      const relationship = workbookRelationships.find(
        (rel) => rel.id === sheet.relationshipId
      );

      if (!relationship) {
        throw new Error(
          `Could not resolve relationship ${sheet.relationshipId} for sheet "${sheet.name}".`
        );
      }

      sheet.relationshipTarget = normalizePackagePath(
        "xl/workbook.xml",
        relationship.target
      );
    }

    const outputZip = new JSZip();

    /*
     * Generate one independent workbook for every worksheet.
     */
    for (const selectedSheet of sheets) {
      const result = await createSingleSheetWorkbook(
        sourceZip,
        workbookXml,
        workbookRelsXml,
        workbookRelationships,
        sheets,
        selectedSheet
      );

      const outputName =
        sanitizeFilename(selectedSheet.name) + extension;

      outputZip.file(outputName, result);
    }

    const zipBuffer = await outputZip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: {
        level: 6,
      },
    });

    return new NextResponse(new Uint8Array(zipBuffer), {
      status: 200,
      headers: {
        "Content-Type": "application/zip",
        "Content-Disposition": 'attachment; filename="split-excel-files.zip"',
        "Cache-Control": "no-store",
      },
    });
  } catch (error) {
    console.error("Excel splitting failed:", error);

    return jsonError(
      error instanceof Error
        ? error.message
        : "Failed to split Excel file.",
      500
    );
  }
}

/* ============================================================
 * CREATE SINGLE WORKBOOK
 * ============================================================
 */

async function createSingleSheetWorkbook(
  sourceZip: JSZip,
  originalWorkbookXml: string,
  originalWorkbookRelsXml: string,
  workbookRelationships: Relationship[],
  sheets: SheetInfo[],
  selectedSheet: SheetInfo
): Promise<Buffer> {
  /*
   * Clone all package parts first.
   *
   * This is intentional.
   *
   * We do NOT reconstruct:
   *
   * - cells
   * - styles
   * - images
   * - charts
   * - drawings
   * - tables
   * - comments
   * - VBA
   * - themes
   *
   * We copy the original package parts directly.
   */
  const zip = new JSZip();

  const files = Object.values(sourceZip.files);

  for (const entry of files) {
    if (entry.dir) {
      continue;
    }

    const data = await entry.async("nodebuffer");

    zip.file(entry.name, data);
  }

  /*
   * ----------------------------------------------------------
   * 1. Modify workbook.xml
   * ----------------------------------------------------------
   */

  let workbookXml = originalWorkbookXml;

  workbookXml = keepOnlySelectedSheet(
    workbookXml,
    selectedSheet
  );

  /*
   * Remove workbook defined names that belong to other sheets.
   */
  workbookXml = filterDefinedNames(
    workbookXml,
    selectedSheet
  );

  /*
   * ----------------------------------------------------------
   * 2. Modify workbook.xml.rels
   * ----------------------------------------------------------
   */

  let workbookRelsXml = originalWorkbookRelsXml;

  const selectedRelationshipIds = new Set([
    selectedSheet.relationshipId,
  ]);

  workbookRelsXml = filterWorkbookRelationships(
    workbookRelsXml,
    workbookRelationships,
    selectedRelationshipIds
  );

  /*
   * ----------------------------------------------------------
   * 3. Determine the dependency tree of the selected worksheet
   * ----------------------------------------------------------
   *
   * Example:
   *
   * sheet1.xml
   *    |
   *    +-- drawing1.xml
   *          |
   *          +-- image1.png
   *
   * sheet1.xml
   *    |
   *    +-- table1.xml
   *
   * sheet1.xml
   *    |
   *    +-- comments1.xml
   *
   * All these parts must remain.
   */

  const requiredParts = new Set<string>();

  await collectPartDependencies(
    sourceZip,
    selectedSheet.relationshipTarget,
    requiredParts
  );

  /*
   * ----------------------------------------------------------
   * 4. Remove worksheet parts belonging to other sheets
   * ----------------------------------------------------------
   */

  const otherWorksheetParts = new Set(
    sheets
      .filter(
        (sheet) => sheet.name !== selectedSheet.name
      )
      .map((sheet) => sheet.relationshipTarget)
  );

  /*
   * Remove unrelated worksheet dependencies.
   *
   * Global parts remain:
   *
   * - styles.xml
   * - sharedStrings.xml
   * - theme
   * - fonts
   * - VBA project
   * - workbook properties
   *
   * This is important for fidelity.
   */

  for (const entry of Object.values(zip.files)) {
    if (entry.dir) continue;

    const path = normalizeZipPath(entry.name);

    if (otherWorksheetParts.has(path)) {
      zip.remove(entry.name);
    }
  }

  /*
   * ----------------------------------------------------------
   * 5. Remove relationship files associated with removed
   *    worksheets
   * ----------------------------------------------------------
   */

  for (const sheet of sheets) {
    if (sheet.name === selectedSheet.name) {
      continue;
    }

    const relPath = relationshipPartPath(
      sheet.relationshipTarget
    );

    if (zip.file(relPath)) {
      zip.remove(relPath);
    }
  }

  /*
   * ----------------------------------------------------------
   * 6. Remove unrelated worksheet-specific parts
   * ----------------------------------------------------------
   *
   * We determine which parts are reachable from the selected
   * worksheet. Worksheet-specific objects that aren't reachable
   * are removed.
   */

  const protectedGlobalParts =
    getGlobalParts(sourceZip);

  for (const entry of Object.values(zip.files)) {
    if (entry.dir) continue;

    const path = normalizeZipPath(entry.name);

    if (
      isWorksheetSpecific(path) &&
      !requiredParts.has(path)
    ) {
      zip.remove(entry.name);
    }
  }

  /*
   * ----------------------------------------------------------
   * 7. Remove calculation chain
   * ----------------------------------------------------------
   *
   * The calculation chain can contain references to cells on
   * worksheets that no longer exist.
   *
   * Excel can rebuild this.
   */

  await removeCalculationChain(zip);

  /*
   * ----------------------------------------------------------
   * 8. Fix [Content_Types].xml
   * ----------------------------------------------------------
   */

  const contentTypesFile =
    zip.file("[Content_Types].xml");

  if (contentTypesFile) {
    let contentTypes =
      await contentTypesFile.async("string");

    contentTypes = removeContentTypesForMissingParts(
      contentTypes,
      zip
    );

    zip.file("[Content_Types].xml", contentTypes);
  }

  /*
   * ----------------------------------------------------------
   * 9. Write modified workbook
   * ----------------------------------------------------------
   */

  zip.file(
    "xl/workbook.xml",
    workbookXml
  );

  zip.file(
    "xl/_rels/workbook.xml.rels",
    workbookRelsXml
  );

  return await zip.generateAsync({
    type: "nodebuffer",
    compression: "DEFLATE",
    compressionOptions: {
      level: 6,
    },
  });
}

/* ============================================================
 * SHEET PARSING
 * ============================================================
 */

function parseSheets(
  workbookXml: string
): SheetInfo[] {
  const sheets: SheetInfo[] = [];

  const sheetsMatch = workbookXml.match(
    /<sheets\b[^>]*>([\s\S]*?)<\/sheets>/i
  );

  if (!sheetsMatch) {
    return sheets;
  }

  const sheetsXml = sheetsMatch[1];

  const sheetRegex =
    /<sheet\b([^>]*?)(?:\/>|>[\s\S]*?<\/sheet>)/gi;

  let match: RegExpExecArray | null;

  while ((match = sheetRegex.exec(sheetsXml))) {
    const attributes = parseAttributes(match[1]);

    const name = attributes.name;
    const sheetId = attributes.sheetId;
    const relationshipId =
      attributes["r:id"] ??
      attributes["id"];

    if (!name || !relationshipId) {
      continue;
    }

    sheets.push({
      name: decodeXml(name),
      sheetId: sheetId ?? "",
      relationshipId,
      relationshipTarget: "",
    });
  }

  return sheets;
}

/* ============================================================
 * RELATIONSHIPS
 * ============================================================
 */

function parseRelationships(
  xml: string
): Relationship[] {
  const relationships: Relationship[] = [];

  const regex =
    /<Relationship\b([^>]*?)(?:\/>|>[\s\S]*?<\/Relationship>)/gi;

  let match: RegExpExecArray | null;

  while ((match = regex.exec(xml))) {
    const attributes = parseAttributes(match[1]);

    if (!attributes.Id || !attributes.Type) {
      continue;
    }

    relationships.push({
      id: attributes.Id,
      type: attributes.Type,
      target: decodeXml(attributes.Target ?? ""),
      targetMode: attributes.TargetMode,
    });
  }

  return relationships;
}

/* ============================================================
 * WORKBOOK SHEET FILTER
 * ============================================================
 */

function keepOnlySelectedSheet(
  xml: string,
  selected: SheetInfo
): string {
  return xml.replace(
    /(<sheets\b[^>]*>)([\s\S]*?)(<\/sheets>)/i,
    (_, open, content, close) => {
      const sheetRegex =
        /<sheet\b([^>]*?)(?:\/>|>[\s\S]*?<\/sheet>)/gi;

      let result = "";

      let match: RegExpExecArray | null;

      while ((match = sheetRegex.exec(content))) {
        const attributes = parseAttributes(
          match[1]
        );

        const relationshipId =
          attributes["r:id"] ??
          attributes["id"];

        if (
          relationshipId === selected.relationshipId
        ) {
          result += match[0];
        }
      }

      return `${open}${result}${close}`;
    }
  );
}

/* ============================================================
 * DEFINED NAMES
 * ============================================================
 */

function filterDefinedNames(
  xml: string,
  selected: SheetInfo
): string {
  /*
   * We preserve workbook-level defined names.
   *
   * Sheet-local defined names are retained only when their
   * localSheetId matches the selected sheet.
   */

  const sheetIndex =
    getSheetIndexFromWorkbook(
      xml,
      selected.relationshipId
    );

  return xml.replace(
    /<definedNames\b[^>]*>([\s\S]*?)<\/definedNames>/i,
    (full, content) => {
      const nameRegex =
        /<definedName\b([^>]*?)>([\s\S]*?)<\/definedName>/gi;

      let names = "";

      let match: RegExpExecArray | null;

      while ((match = nameRegex.exec(content))) {
        const attrs = parseAttributes(match[1]);

        if (
          attrs.localSheetId === undefined ||
          Number(attrs.localSheetId) === sheetIndex
        ) {
          names += match[0];
        }
      }

      if (!names) {
        return "";
      }

      return `<definedNames>${names}</definedNames>`;
    }
  );
}

function getSheetIndexFromWorkbook(
  xml: string,
  relationshipId: string
): number {
  const sheetsMatch = xml.match(
    /<sheets\b[^>]*>([\s\S]*?)<\/sheets>/i
  );

  if (!sheetsMatch) {
    return 0;
  }

  const content = sheetsMatch[1];

  const sheetRegex =
    /<sheet\b([^>]*?)(?:\/>|>[\s\S]*?<\/sheet>)/gi;

  let index = 0;

  let match: RegExpExecArray | null;

  while ((match = sheetRegex.exec(content))) {
    const attrs = parseAttributes(match[1]);

    const id =
      attrs["r:id"] ??
      attrs["id"];

    if (id === relationshipId) {
      return index;
    }

    index++;
  }

  return 0;
}

/* ============================================================
 * WORKBOOK RELATIONSHIPS
 * ============================================================
 */

function filterWorkbookRelationships(
  xml: string,
  relationships: Relationship[],
  allowedIds: Set<string>
): string {
  const allowedRelationships =
    relationships.filter((rel) => {
      /*
       * Keep:
       *
       * - selected worksheet
       * - styles
       * - shared strings
       * - theme
       * - VBA
       * - workbook-level objects
       *
       * Remove only relationships that directly point to
       * other worksheets.
       */

      if (
        rel.type.endsWith("/worksheet")
      ) {
        return allowedIds.has(rel.id);
      }

      if (
        rel.type.endsWith("/chartsheet")
      ) {
        return false;
      }

      return true;
    });

  const allowedSet = new Set(
    allowedRelationships.map((r) => r.id)
  );

  return xml.replace(
    /<Relationship\b([^>]*?)(?:\/>|>[\s\S]*?<\/Relationship>)/gi,
    (full, attrsText) => {
      const attrs = parseAttributes(attrsText);

      if (!attrs.Id) {
        return full;
      }

      return allowedSet.has(attrs.Id)
        ? full
        : "";
    }
  );
}

/* ============================================================
 * DEPENDENCY GRAPH
 * ============================================================
 */

async function collectPartDependencies(
  zip: JSZip,
  partPath: string,
  visited: Set<string>
): Promise<void> {
  partPath = normalizeZipPath(partPath);

  if (visited.has(partPath)) {
    return;
  }

  visited.add(partPath);

  const file = zip.file(partPath);

  if (!file) {
    return;
  }

  /*
   * Only XML parts can contain relationship references.
   */

  if (!partPath.endsWith(".xml")) {
    return;
  }

  const xml = await file.async("string");

  const relPath =
    relationshipPartPath(partPath);

  const relFile = zip.file(relPath);

  if (!relFile) {
    return;
  }

  const relXml =
    await relFile.async("string");

  const relationships =
    parseRelationships(relXml);

  for (const relationship of relationships) {
    /*
     * External links don't belong inside the ZIP.
     */
    if (
      relationship.targetMode?.toLowerCase() ===
      "external"
    ) {
      continue;
    }

    /*
     * Relationship parts themselves are required.
     */
    visited.add(relPath);

    const target =
      normalizePackagePath(
        partPath,
        relationship.target
      );

    if (!target) {
      continue;
    }

    if (!zip.file(target)) {
      continue;
    }

    await collectPartDependencies(
      zip,
      target,
      visited
    );
  }
}

/* ============================================================
 * CONTENT TYPES
 * ============================================================
 */

function removeContentTypesForMissingParts(
  xml: string,
  zip: JSZip
): string {
  /*
   * Remove <Override PartName="..."/> entries whose target
   * no longer exists.
   */

  return xml.replace(
    /<Override\b([^>]*?)(?:\/>|>[\s\S]*?<\/Override>)/gi,
    (full, attrsText) => {
      const attrs = parseAttributes(attrsText);

      if (!attrs.PartName) {
        return full;
      }

      const part =
        normalizeZipPath(
          attrs.PartName.replace(/^\//, "")
        );

      if (!zip.file(part)) {
        return "";
      }

      return full;
    }
  );
}

/* ============================================================
 * CALCULATION CHAIN
 * ============================================================
 */

async function removeCalculationChain(zip: JSZip) {
  const candidates = ["xl/calcChain.xml", "xl/calcChain.xml.rels"];

  for (const path of candidates) {
    if (zip.file(path)) {
      zip.remove(path);
    }
  }

  const relPath = "xl/_rels/workbook.xml.rels";

  const relFile = zip.file(relPath);

  if (!relFile) {
    return;
  }

  const xml = await relFile.async("string");

  const updated = xml.replace(
    /<Relationship\b([^>]*?)(?:\/>|>[\s\S]*?<\/Relationship>)/gi,
    (full, attrsText) => {
      const attrs = parseAttributes(attrsText);

      if (attrs.Type?.endsWith("/calcChain")) {
        return "";
      }

      return full;
    },
  );

  zip.file(relPath, updated);
}

/* ============================================================
 * PATH UTILITIES
 * ============================================================
 */

function relationshipPartPath(
  partPath: string
): string {
  const normalized =
    normalizeZipPath(partPath);

  const slash =
    normalized.lastIndexOf("/");

  if (slash === -1) {
    return "_rels/" + normalized + ".rels";
  }

  const directory =
    normalized.substring(0, slash);

  const filename =
    normalized.substring(slash + 1);

  return `${directory}/_rels/${filename}.rels`;
}

function normalizePackagePath(
  sourcePart: string,
  target: string
): string {
  if (!target) {
    return "";
  }

  if (
    target.startsWith("http://") ||
    target.startsWith("https://") ||
    target.startsWith("/")
  ) {
    return normalizeZipPath(
      target.replace(/^\//, "")
    );
  }

  const source =
    normalizeZipPath(sourcePart);

  const sourceDir =
    source.substring(
      0,
      source.lastIndexOf("/") + 1
    );

  return normalizeZipPath(
    sourceDir + target
  );
}

function normalizeZipPath(
  path: string
): string {
  const parts = path
    .replace(/\\/g, "/")
    .split("/");

  const output: string[] = [];

  for (const part of parts) {
    if (!part || part === ".") {
      continue;
    }

    if (part === "..") {
      output.pop();
      continue;
    }

    output.push(part);
  }

  return output.join("/");
}

/* ============================================================
 * WORKSHEET-SPECIFIC PART DETECTION
 * ============================================================
 */

function isWorksheetSpecific(
  path: string
): boolean {
  return (
    path.startsWith("xl/worksheets/") ||
    path.startsWith("xl/drawings/") ||
    path.startsWith("xl/comments") ||
    path.startsWith("xl/threadedComments/") ||
    path.startsWith("xl/persons/") ||
    path.startsWith("xl/tables/") ||
    path.startsWith("xl/pivotTables/") ||
    path.startsWith("xl/pivotCache/") ||
    path.startsWith("xl/queryTables/") ||
    path.startsWith("xl/externalLinks/")
  );
}

function getGlobalParts(
  zip: JSZip
): Set<string> {
  const result = new Set<string>();

  for (const entry of Object.values(zip.files)) {
    const path = normalizeZipPath(
      entry.name
    );

    if (!isWorksheetSpecific(path)) {
      result.add(path);
    }
  }

  return result;
}

/* ============================================================
 * XML ATTRIBUTE PARSER
 * ============================================================
 */

function parseAttributes(
  text: string
): Record<string, string> {
  const result: Record<string, string> = {};

  const regex =
    /([A-Za-z_][\w:.-]*)\s*=\s*("([^"]*)"|'([^']*)')/g;

  let match: RegExpExecArray | null;

  while ((match = regex.exec(text))) {
    result[match[1]] =
      match[3] ?? match[4] ?? "";
  }

  return result;
}

/* ============================================================
 * XML UTILITIES
 * ============================================================
 */

function decodeXml(
  value: string
): string {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");
}

/* ============================================================
 * FILE NAME
 * ============================================================
 */

function sanitizeFilename(
  value: string
): string {
  return (
    value
      .replace(/[<>:"/\\|?*\x00-\x1F]/g, "_")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 150) || "Sheet"
  );
}

function getExtension(
  filename: string
): string {
  const index =
    filename.lastIndexOf(".");

  if (index === -1) {
    return "";
  }

  return filename
    .substring(index)
    .toLowerCase();
}

/* ============================================================
 * ERROR RESPONSE
 * ============================================================
 */

function jsonError(
  message: string,
  status: number
) {
  return NextResponse.json(
    {
      error: message,
    },
    {
      status,
    }
  );
}