import JSZip from "jszip";

/**
 * Decode basic XML entities
 */
const decodeXml = (value) => {
  return value
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">");
};

/**
 * Make a safe filename
 */
const sanitizeFileName = (name) => {
  return name
    .replace(/[<>:"/\\|?*]/g, "_")
    .trim()
    .substring(0, 200);
};

/**
 * Read sheets directly from workbook.xml
 *
 * Example:
 *
 * <sheet
 *   name="Record_1"
 *   sheetId="1"
 *   r:id="rId1"
 * />
 */
const getSheets = async (zip) => {
  const workbookFile = zip.file("xl/workbook.xml");

  if (!workbookFile) {
    throw new Error("Invalid XLSX: workbook.xml not found.");
  }

  const workbookXml = await workbookFile.async("string");

  const sheetsSection = workbookXml.match(
    /<sheets\b[^>]*>([\s\S]*?)<\/sheets>/,
  );

  if (!sheetsSection) {
    throw new Error("No <sheets> section found in workbook.");
  }

  const sheetsXml = sheetsSection[1];

  const sheetMatches = [
    ...sheetsXml.matchAll(/<sheet\b([^>]*?)(?:\/>|>[\s\S]*?<\/sheet>)/),
  ];

  const sheets = [];

  for (const match of sheetMatches) {
    const attributes = match[1];

    const nameMatch = attributes.match(/\bname=(["'])(.*?)\1/);

    const sheetIdMatch = attributes.match(/\bsheetId=(["'])(.*?)\1/);

    const relationshipMatch = attributes.match(/\br:id=(["'])(.*?)\1/);

    if (!nameMatch || !relationshipMatch) {
      continue;
    }

    sheets.push({
      name: decodeXml(nameMatch[2]),
      sheetId: sheetIdMatch ? sheetIdMatch[2] : null,
      rId: relationshipMatch[2],
      xml: match[0],
    });
  }

  return sheets;
};

/**
 * Read workbook relationships.
 *
 * Example:
 *
 * <Relationship
 *   Id="rId1"
 *   Type="..."
 *   Target="worksheets/sheet1.xml"
 * />
 */
const getRelationships = async (zip) => {
  const relFile = zip.file("xl/_rels/workbook.xml.rels");

  if (!relFile) {
    throw new Error("Invalid XLSX: workbook.xml.rels not found.");
  }

  const relXml = await relFile.async("string");

  const relationships = {};

  const relationshipMatches = [
    ...relXml.matchAll(
      /<Relationship\b([^>]*?)(?:\/>|>[\s\S]*?<\/Relationship>)/,
    ),
  ];

  for (const match of relationshipMatches) {
    const attributes = match[1];

    const idMatch = attributes.match(/\bId=(["'])(.*?)\1/);

    const targetMatch = attributes.match(/\bTarget=(["'])(.*?)\1/);

    if (!idMatch || !targetMatch) {
      continue;
    }

    relationships[idMatch[2]] = targetMatch[2];
  }

  return relationships;
};

/**
 * Extract all relationship XML nodes
 */
const getRelationshipNodes = (relXml) => {
  return [
    ...relXml.matchAll(
      /<Relationship\b[^>]*?(?:\/>|>[\s\S]*?<\/Relationship>)/g,
    ),
  ].map((match) => match[0]);
};

/**
 * Get relationship Id from XML
 */
const getRelationshipId = (xml) => {
  const match = xml.match(/\bId=(["'])(.*?)\1/);

  return match ? match[2] : null;
};

/**
 * Split workbook while preserving the original
 * XLSX package.
 */
export const splitExcelWorkbook = async (
  inputFile,
  selectedSheetNames = null,
) => {
  const inputBuffer = await inputFile.arrayBuffer();

  const originalZip = await JSZip.loadAsync(inputBuffer);

  const sheets = await getSheets(originalZip);

  if (!sheets.length) {
    throw new Error("No worksheets found in workbook.");
  }

  const relationships = await getRelationships(originalZip);

  const selectedSheets = selectedSheetNames?.length
    ? sheets.filter((sheet) => selectedSheetNames.includes(sheet.name))
    : sheets;

  if (!selectedSheets.length) {
    throw new Error("No selected worksheets found.");
  }

  const outputFiles = [];

  for (const currentSheet of selectedSheets) {
    console.log("Creating:", currentSheet.name);

    const outputZip = new JSZip();

    /*
     * ------------------------------------------
     * Copy the ENTIRE original XLSX package
     * ------------------------------------------
     *
     * This is what helps preserve:
     *
     * - styles
     * - fonts
     * - fills
     * - borders
     * - number formats
     * - merged cells
     * - row heights
     * - column widths
     * - formulas
     * - shared strings
     * - themes
     * - drawings
     * - images
     * - worksheet relationships
     */
    const files = Object.keys(originalZip.files);

    for (const fileName of files) {
      const file = originalZip.files[fileName];

      if (file.dir) {
        outputZip.folder(fileName);
        continue;
      }

      /*
       * These three files need to be rebuilt.
       */
      if (
        fileName === "xl/workbook.xml" ||
        fileName === "xl/_rels/workbook.xml.rels" ||
        fileName === "[Content_Types].xml"
      ) {
        continue;
      }

      const content = await file.async("uint8array");

      outputZip.file(fileName, content);
    }

    /*
     * ------------------------------------------
     * workbook.xml
     * ------------------------------------------
     */

    const workbookXml = await originalZip
      .file("xl/workbook.xml")
      .async("string");

    /*
     * Find selected sheet XML node.
     */
    const selectedSheetXml = currentSheet.xml;

    /*
     * Replace ALL sheets with ONLY
     * the selected sheet.
     */
    const newWorkbookXml = workbookXml.replace(
      /(<sheets\b[^>]*>)[\s\S]*?(<\/sheets>)/,
      `$1${selectedSheetXml}$2`,
    );

    outputZip.file("xl/workbook.xml", newWorkbookXml);

    /*
     * ------------------------------------------
     * workbook.xml.rels
     * ------------------------------------------
     */

    const relXml = await originalZip
      .file("xl/_rels/workbook.xml.rels")
      .async("string");

    const relationshipNodes = getRelationshipNodes(relXml);

    const selectedRelationship = relationshipNodes.find(
      (relationshipXml) =>
        getRelationshipId(relationshipXml) === currentSheet.rId,
    );

    if (!selectedRelationship) {
      throw new Error(
        `Relationship ${currentSheet.rId} not found for sheet "${currentSheet.name}".`,
      );
    }

    /*
     * Keep only the selected sheet relationship.
     *
     * IMPORTANT:
     * Other relationship types such as
     * theme, styles, shared strings etc.
     * are also required, so we don't remove
     * those.
     */
    const relationshipType = selectedRelationship.match(/\bType=(["'])(.*?)\1/);

    const selectedTarget = selectedRelationship.match(/\bTarget=(["'])(.*?)\1/);

    const newRelationshipNodes = relationshipNodes.filter((relationshipXml) => {
      const typeMatch = relationshipXml.match(/\bType=(["'])(.*?)\1/);

      /*
       * Remove worksheet relationships
       * except selected worksheet.
       */
      if (typeMatch && typeMatch[2].endsWith("/worksheet")) {
        return getRelationshipId(relationshipXml) === currentSheet.rId;
      }

      /*
       * Keep everything else:
       *
       * styles
       * theme
       * shared strings
       * calc chain
       * etc.
       */
      return true;
    });

    const newRelXml = relXml.replace(
      /(<Relationships\b[^>]*>)[\s\S]*?(<\/Relationships>)/,
      `$1${newRelationshipNodes.join("")}$2`,
    );

    outputZip.file("xl/_rels/workbook.xml.rels", newRelXml);

    /*
     * ------------------------------------------
     * Content Types
     * ------------------------------------------
     *
     * We leave the original content types.
     * Extra worksheet overrides are harmless.
     * This avoids accidentally removing required
     * content types for images, drawings, etc.
     */
    const contentTypesXml = await originalZip
      .file("[Content_Types].xml")
      .async("string");

    outputZip.file("[Content_Types].xml", contentTypesXml);

    /*
     * ------------------------------------------
     * Generate XLSX
     * ------------------------------------------
     */

    const blob = await outputZip.generateAsync({
      type: "blob",
      compression: "DEFLATE",
      compressionOptions: {
        level: 6,
      },
    });

    outputFiles.push({
      name: `${sanitizeFileName(currentSheet.name)}.xlsx`,
      blob,
    });
  }

  return outputFiles;
};
