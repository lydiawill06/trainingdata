const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const ExcelJS = require("exceljs");

function normalizeHeader(header) {
  return String(header)
    .replace(/ /g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function buildNormMap(rows) {
  const map = new Map();
  if (rows.length === 0) return map;
  for (const key of Object.keys(rows[0])) {
    map.set(normalizeHeader(key), key);
  }
  return map;
}

function getVal(row, normMap, candidateName) {
  const originalKey = normMap.get(candidateName);
  if (originalKey === undefined) return undefined;
  const val = row[originalKey];
  if (val === undefined || val === null) return "";
  return String(val);
}

function getFirstVal(row, normMap, candidateNames) {
  for (const candidate of candidateNames) {
    const val = getVal(row, normMap, candidate);
    if (val !== undefined) return val;
  }
  return undefined;
}

function rowIsComplete(row, normMap) {
  const isComplete = getVal(row, normMap, "is complete");
  if (isComplete !== undefined) {
    return isComplete.trim().toUpperCase() === "Y";
  }
  const statusId = getVal(row, normMap, "completion status id");
  if (statusId !== undefined) {
    return statusId.trim().toLowerCase() === "online complete";
  }
  const status = getVal(row, normMap, "completion status");
  if (status !== undefined) {
    const v = status.trim().toLowerCase();
    return v.includes("complete") && !v.includes("incomplete") && !v.includes("not complete");
  }
  return false;
}

function rowTrainingName(row, normMap) {
  const programTitle = getVal(row, normMap, "program title");
  if (programTitle && programTitle.trim()) return programTitle.trim();
  const itemTitle = getVal(row, normMap, "item title");
  if (itemTitle && itemTitle.trim()) return itemTitle.trim();
  return null;
}

function collapseWhitespace(str) {
  return String(str).replace(/\s+/g, " ").trim();
}

function setIfBlank(info, key, value) {
  if (value === undefined) return;
  const clean = collapseWhitespace(value);
  if (!clean) return;
  if (!info[key]) info[key] = clean;
}

const ALLOWED_LEGAL_ENTITIES = [
  "ENGIE North America Inc.",
  "MATEP LLC",
  "SoCore Energy LLC",
  "SoCore Installation Services LLC",
];
const ALLOWED_LEGAL_ENTITY_SET = new Set(ALLOWED_LEGAL_ENTITIES.map((name) => name.toLowerCase()));

function cleanLegalEntity(value) {
  return collapseWhitespace(String(value).replace(/\s*\([^)]*\)\s*$/, ""));
}

function readFileAsRows(filePath) {
  const workbook = XLSX.readFile(filePath);
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  return XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
}

function compileFiles(filesData) {
  const people = new Map();
  const trainingSet = new Set();
  const statusMap = new Map();
  const personInfo = new Map();
  let skippedRows = 0;

  for (const file of filesData) {
    const { rows, normMap } = file;
    for (const row of rows) {
      const rawFirst = getVal(row, normMap, "first name");
      const rawLast = getVal(row, normMap, "last name");
      if (!rawFirst || !rawLast || !rawFirst.trim() || !rawLast.trim()) {
        skippedRows++;
        continue;
      }
      const rawUserId = getFirstVal(row, normMap, ["user id"]);
      if (!rawUserId || !rawUserId.trim()) {
        skippedRows++;
        continue;
      }
      const rawLegalEntity = getVal(row, normMap, "legal entity");
      const legalEntity = rawLegalEntity ? cleanLegalEntity(rawLegalEntity) : "";
      if (!legalEntity || !ALLOWED_LEGAL_ENTITY_SET.has(legalEntity.toLowerCase())) {
        skippedRows++;
        continue;
      }
      const trainingName = rowTrainingName(row, normMap);
      if (!trainingName) {
        skippedRows++;
        continue;
      }
      const firstName = collapseWhitespace(rawFirst);
      const lastName = collapseWhitespace(rawLast);
      const userId = collapseWhitespace(rawUserId);
      const personKey = userId;

      if (!people.has(personKey)) {
        people.set(personKey, { userId, firstName, lastName });
      }
      trainingSet.add(trainingName);

      if (!statusMap.has(personKey)) {
        statusMap.set(personKey, new Map());
      }
      const personStatus = statusMap.get(personKey);
      if (!personStatus.has(trainingName)) {
        personStatus.set(trainingName, { allComplete: true });
      }
      const entry = personStatus.get(trainingName);
      entry.allComplete = entry.allComplete && rowIsComplete(row, normMap);

      if (!personInfo.has(personKey)) {
        personInfo.set(personKey, {
          managerId: "",
          managerFirstName: "",
          managerLastName: "",
          organizationId: "",
          legalEntity: "",
        });
      }
      const info = personInfo.get(personKey);
      setIfBlank(info, "managerId", getFirstVal(row, normMap, ["manager id", "manager id (user)"]));
      setIfBlank(info, "managerFirstName", getFirstVal(row, normMap, ["first name manager", "manager first name"]));
      setIfBlank(info, "managerLastName", getFirstVal(row, normMap, ["last name manager", "manager last name"]));
      setIfBlank(info, "organizationId", getFirstVal(row, normMap, ["organization id"]));
      setIfBlank(info, "legalEntity", legalEntity);
    }
  }

  const trainings = Array.from(trainingSet).sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: "base" })
  );

  const peopleList = Array.from(people.entries()).map(([key, val]) => ({ key, ...val }));
  peopleList.sort((a, b) => {
    const lastCmp = a.lastName.localeCompare(b.lastName, undefined, { sensitivity: "base" });
    if (lastCmp !== 0) return lastCmp;
    return a.firstName.localeCompare(b.firstName, undefined, { sensitivity: "base" });
  });

  const matrix = peopleList.map((person) => {
    const personStatus = statusMap.get(person.key) || new Map();
    const cells = trainings.map((training) => {
      const entry = personStatus.get(training);
      if (!entry) return "Unassigned";
      return entry.allComplete ? "Complete" : "Incomplete";
    });
    const completedCount = cells.filter((status) => status === "Complete").length;
    const percentComplete = trainings.length > 0 ? `${completedCount}/${trainings.length}` : "0/0";
    const info = personInfo.get(person.key) || {};
    return {
      userId: person.userId,
      firstName: person.firstName,
      lastName: person.lastName,
      legalEntity: info.legalEntity || "",
      cells,
      percentComplete,
      managerId: info.managerId || "",
      managerFirstName: info.managerFirstName || "",
      managerLastName: info.managerLastName || "",
      organizationId: info.organizationId || "",
    };
  });

  return { trainings, matrix, skippedRows, fileCount: filesData.length };
}

const STATUS_COLORS = {
  Complete: { font: "FF1F9D55", fill: "FFE6F7EC" },
  Incomplete: { font: "FFC0392B", fill: "FFFDECEA" },
  Unassigned: { font: "FF8A94A3", fill: "FFEEF0F3" },
};

async function exportWorkbook(compiled, outPath) {
  const { trainings, matrix } = compiled;
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Training Roster", {
    views: [{ state: "frozen", ySplit: 1 }],
  });

  const headerLabels = [
    "User ID",
    "First Name",
    "Last Name",
    ...trainings,
    "Percent Complete",
    "Manager ID",
    "Manager First Name",
    "Manager Last Name",
    "Organization ID",
    "Legal Entity",
  ];
  sheet.columns = headerLabels.map((label) => ({
    header: label,
    key: label,
    width: Math.min(Math.max(label.length + 4, 14), 40),
  }));

  matrix.forEach((person) => {
    sheet.addRow([
      person.userId,
      person.firstName,
      person.lastName,
      ...person.cells,
      person.percentComplete,
      person.managerId,
      person.managerFirstName,
      person.managerLastName,
      person.organizationId,
      person.legalEntity,
    ]);
  });

  const headerRow = sheet.getRow(1);
  headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
  headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2F6FED" } };
  headerRow.alignment = { vertical: "middle", horizontal: "center" };

  sheet.autoFilter = { from: { row: 1, column: 1 }, to: { row: 1, column: headerLabels.length } };

  const leadingCols = 3; // User ID, First Name, Last Name
  const lastTrainingCol = leadingCols + trainings.length;
  for (let r = 2; r <= sheet.rowCount; r++) {
    const row = sheet.getRow(r);
    for (let c = leadingCols + 1; c <= lastTrainingCol; c++) {
      const cell = row.getCell(c);
      const colors = STATUS_COLORS[cell.value];
      if (colors) {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.fill } };
        cell.font = { color: { argb: colors.font }, bold: true };
      }
      cell.alignment = { horizontal: "center" };
    }
    for (let c = lastTrainingCol + 1; c <= headerLabels.length; c++) {
      row.getCell(c).alignment = { horizontal: "center" };
    }
    for (let c = 1; c <= leadingCols; c++) {
      row.getCell(c).alignment = { horizontal: "left" };
    }
  }

  sheet.eachRow((row) => {
    row.eachCell((cell) => {
      cell.border = {
        top: { style: "thin", color: { argb: "FFDDE1E7" } },
        left: { style: "thin", color: { argb: "FFDDE1E7" } },
        bottom: { style: "thin", color: { argb: "FFDDE1E7" } },
        right: { style: "thin", color: { argb: "FFDDE1E7" } },
      };
    });
  });

  await workbook.xlsx.writeFile(outPath);
}

(async () => {
  const file1 = process.argv[2];
  const file2 = process.argv[3];
  const outPath = process.argv[4];

  const filesData = [file1, file2].map((p) => {
    const rows = readFileAsRows(p);
    return { name: path.basename(p), rows, normMap: buildNormMap(rows) };
  });

  const compiled = compileFiles(filesData);
  console.log(`People: ${compiled.matrix.length}, Trainings: ${compiled.trainings.length}, Skipped rows: ${compiled.skippedRows}`);

  await exportWorkbook(compiled, outPath);
  console.log("Wrote", outPath);
})();
