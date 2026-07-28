(() => {
  "use strict";

  const fileInput = document.getElementById("file-input");
  const fileDrop = document.getElementById("file-drop");
  const fileListEl = document.getElementById("file-list");
  const compileBtn = document.getElementById("compile-btn");
  const clearBtn = document.getElementById("clear-btn");
  const statusEl = document.getElementById("status");
  const previewPanel = document.getElementById("preview-panel");
  const exportPanel = document.getElementById("export-panel");
  const previewTable = document.getElementById("preview-table");
  const summaryHint = document.getElementById("summary-hint");
  const exportBtn = document.getElementById("export-btn");
  const newCompileDateRow = document.getElementById("new-compile-date-row");
  const newCompileDateInput = document.getElementById("new-compile-date-input");

  // Each entry: { name, rows, normMap, kind: "raw"|"compiled", dateGuess: Date|null, dateLabel: string }
  const filesData = [];

  // Result of the last compile, kept for export.
  let compiled = null;

  function setStatus(message, isError) {
    statusEl.textContent = message || "";
    statusEl.classList.toggle("error", Boolean(isError));
  }

  function normalizeHeader(header) {
    return String(header)
      .replace(/ /g, " ")
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

  // --- File kind & date detection --------------------------------------------

  // A file is a raw training-system export if it has a training-name column
  // ("Program Title"/"Item Title") or a completion-signal column ("Is
  // Complete"/"Completion Status ID"/"Completion Status") - the columns this
  // tool's own compiled exports never have (with or without "Percent
  // Complete", which was only added partway through this tool's history).
  const RAW_SHAPE_MARKERS = [
    "program title",
    "item title",
    "is complete",
    "completion status id",
    "completion status",
  ];
  function detectKind(normMap) {
    const isRaw = RAW_SHAPE_MARKERS.some((marker) => normMap.has(marker));
    return isRaw ? "raw" : "compiled";
  }

  function guessDateFromFilename(name) {
    let m = name.match(/(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (m) {
      const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
      if (!isNaN(d)) return d;
    }
    // e.g. "7.1.26" or "7-1-26"
    m = name.match(/\b(\d{1,2})[.\-](\d{1,2})[.\-](\d{2})\b/);
    if (m) {
      const month = Number(m[1]);
      const day = Number(m[2]);
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        const d = new Date(2000 + Number(m[3]), month - 1, day);
        if (!isNaN(d)) return d;
      }
    }
    // e.g. "7.1.2026" or "7-1-2026"
    m = name.match(/\b(\d{1,2})[.\-](\d{1,2})[.\-](\d{4})\b/);
    if (m) {
      const month = Number(m[1]);
      const day = Number(m[2]);
      if (month >= 1 && month <= 12 && day >= 1 && day <= 31) {
        const d = new Date(Number(m[3]), month - 1, day);
        if (!isNaN(d)) return d;
      }
    }
    return null;
  }

  function formatDateLabel(date) {
    if (!date || isNaN(date)) return "";
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, "0");
    const d = String(date.getDate()).padStart(2, "0");
    return `${y}-${m}-${d}`;
  }

  function parseDateLabel(label) {
    if (!label || !label.trim()) return null;
    const iso = label.trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (iso) {
      const d = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
      if (!isNaN(d)) return d;
    }
    const d = new Date(label.trim());
    return isNaN(d) ? null : d;
  }

  function readFileAsRows(file) {
    const isCsv = /\.csv$/i.test(file.name);
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(reader.error || new Error("Failed to read file"));
      reader.onload = () => {
        try {
          const workbook = isCsv
            ? XLSX.read(reader.result, { type: "string" })
            : XLSX.read(reader.result, { type: "array" });
          const firstSheetName = workbook.SheetNames[0];
          const sheet = workbook.Sheets[firstSheetName];
          const rows = XLSX.utils.sheet_to_json(sheet, { defval: "", raw: false });
          resolve(rows);
        } catch (err) {
          reject(err);
        }
      };
      if (isCsv) {
        reader.readAsText(file);
      } else {
        reader.readAsArrayBuffer(file);
      }
    });
  }

  function renderFileList() {
    fileListEl.innerHTML = "";
    filesData.forEach((entry, index) => {
      const li = document.createElement("li");

      const meta = document.createElement("span");
      meta.className = "file-meta";
      meta.textContent = entry.name;
      li.appendChild(meta);

      const badge = document.createElement("span");
      badge.className = `kind-badge ${entry.kind}`;
      badge.textContent = entry.kind === "compiled" ? "Past compilation" : "Raw data";
      li.appendChild(badge);

      if (entry.kind === "compiled") {
        const dateInput = document.createElement("input");
        dateInput.type = "text";
        dateInput.className = "date-input";
        dateInput.value = entry.dateLabel;
        dateInput.title = "Date this past compilation represents";
        dateInput.addEventListener("input", () => {
          entry.dateLabel = dateInput.value;
        });
        li.appendChild(dateInput);
      }

      const count = document.createElement("span");
      count.className = "row-count";
      count.textContent = `${entry.rows.length} rows`;
      li.appendChild(count);

      const removeBtn = document.createElement("button");
      removeBtn.type = "button";
      removeBtn.textContent = "Remove";
      removeBtn.addEventListener("click", () => {
        filesData.splice(index, 1);
        renderFileList();
        updateActionState();
        hideResults();
      });
      li.appendChild(removeBtn);

      fileListEl.appendChild(li);
    });

    updateNewCompileDateVisibility();
  }

  function updateNewCompileDateVisibility() {
    const hasRaw = filesData.some((e) => e.kind === "raw");
    newCompileDateRow.hidden = !hasRaw;
    if (!hasRaw) {
      newCompileDateInput.value = "";
      return;
    }
    if (!newCompileDateInput.value.trim()) {
      const guesses = filesData
        .filter((e) => e.kind === "raw" && e.dateGuess)
        .map((e) => e.dateGuess.getTime());
      const guess = guesses.length ? new Date(Math.min(...guesses)) : new Date();
      newCompileDateInput.value = formatDateLabel(guess);
    }
  }

  function updateActionState() {
    const hasFiles = filesData.length > 0;
    compileBtn.disabled = !hasFiles;
    clearBtn.disabled = !hasFiles;
  }

  function hideResults() {
    previewPanel.hidden = true;
    exportPanel.hidden = true;
    compiled = null;
  }

  async function handleFiles(fileList) {
    const incoming = Array.from(fileList).filter((f) =>
      /\.(csv|xlsx|xls)$/i.test(f.name)
    );
    if (incoming.length === 0) {
      setStatus("Please choose .csv, .xlsx, or .xls files.", true);
      return;
    }

    setStatus(`Reading ${incoming.length} file(s)...`);
    for (const file of incoming) {
      try {
        const rows = await readFileAsRows(file);
        const normMap = buildNormMap(rows);
        const dateGuess =
          guessDateFromFilename(file.name) ||
          (file.lastModified ? new Date(file.lastModified) : null);
        filesData.push({
          name: file.name,
          rows,
          normMap,
          kind: detectKind(normMap),
          dateGuess,
          dateLabel: formatDateLabel(dateGuess),
        });
      } catch (err) {
        setStatus(`Failed to read "${file.name}": ${err.message}`, true);
      }
    }
    renderFileList();
    updateActionState();
    hideResults();
    setStatus(`${filesData.length} file(s) loaded. Click "Compile Files" to continue.`);
  }

  fileDrop.addEventListener("dragover", (e) => {
    e.preventDefault();
    fileDrop.classList.add("dragover");
  });
  fileDrop.addEventListener("dragleave", () => {
    fileDrop.classList.remove("dragover");
  });
  fileDrop.addEventListener("drop", (e) => {
    e.preventDefault();
    fileDrop.classList.remove("dragover");
    if (e.dataTransfer && e.dataTransfer.files.length) {
      handleFiles(e.dataTransfer.files);
    }
  });

  fileInput.addEventListener("change", () => {
    if (fileInput.files.length) {
      handleFiles(fileInput.files);
    }
    fileInput.value = "";
  });

  clearBtn.addEventListener("click", () => {
    filesData.length = 0;
    renderFileList();
    updateActionState();
    hideResults();
    setStatus("");
  });

  // --- Completion detection -------------------------------------------------

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

  // --- Legal entity filtering -------------------------------------------------

  const ALLOWED_LEGAL_ENTITIES = [
    "ENGIE North America Inc.",
    "MATEP LLC",
    "SoCore Energy LLC",
    "SoCore Installation Services LLC",
  ];
  const ALLOWED_LEGAL_ENTITY_SET = new Set(
    ALLOWED_LEGAL_ENTITIES.map((name) => name.toLowerCase())
  );

  // Strips a trailing code suffix like " (42775)" from a raw Legal Entity value.
  function cleanLegalEntity(value) {
    return collapseWhitespace(String(value).replace(/\s*\([^)]*\)\s*$/, ""));
  }

  // --- Compile raw training-system exports ------------------------------------

  function compileRawFiles(entries) {
    const people = new Map(); // personKey (User ID) -> { userId, firstName, lastName }
    const trainingSet = new Set();
    // personKey -> Map(trainingName -> { allComplete: bool })
    const statusMap = new Map();
    // personKey -> { managerId, managerFirstName, managerLastName, organizationId, legalEntity }
    const personInfo = new Map();

    let skippedRows = 0;

    for (const file of entries) {
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

    const peopleList = Array.from(people.entries()).map(([key, val]) => ({
      key,
      ...val,
    }));
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
        key: person.userId,
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

    return { trainings, matrix, skippedRows, fileCount: entries.length };
  }

  // --- Parse a past compilation (this tool's own export format) --------------

  const FIXED_COMPILED_COLUMNS = new Set([
    "user id",
    "first name",
    "last name",
    "percent complete",
    "manager id",
    "manager first name",
    "manager last name",
    "organization id",
    "legal entity",
  ]);

  function parseCompiledFile(entry) {
    const { rows, normMap } = entry;
    const trainings = [];
    for (const [normKey, originalKey] of normMap.entries()) {
      if (!FIXED_COMPILED_COLUMNS.has(normKey)) {
        trainings.push(originalKey);
      }
    }

    const matrix = rows
      .map((row) => {
        const cells = trainings.map((training) => {
          const val = row[training];
          const clean = val === undefined || val === null ? "" : collapseWhitespace(val);
          return clean || "Unassigned";
        });
        const userId = collapseWhitespace(getVal(row, normMap, "user id") || "");
        const firstName = collapseWhitespace(getVal(row, normMap, "first name") || "");
        const lastName = collapseWhitespace(getVal(row, normMap, "last name") || "");
        return {
          // Older exports (from before this tool tracked User ID) only have
          // names, so fall back to a name-based key for matching across
          // snapshots when no User ID is present.
          key: userId || `name:${firstName.toLowerCase()}|${lastName.toLowerCase()}`,
          userId,
          firstName,
          lastName,
          legalEntity: collapseWhitespace(getVal(row, normMap, "legal entity") || ""),
          cells,
          percentComplete: getVal(row, normMap, "percent complete") || "",
          managerId: getVal(row, normMap, "manager id") || "",
          managerFirstName: getVal(row, normMap, "manager first name") || "",
          managerLastName: getVal(row, normMap, "manager last name") || "",
          organizationId: getVal(row, normMap, "organization id") || "",
        };
      })
      .filter((person) => person.firstName && person.lastName);

    return { trainings, matrix };
  }

  // --- Build one snapshot per input dataset (raw compile + each past file) ---

  function buildSnapshots() {
    const rawEntries = filesData.filter((e) => e.kind === "raw");
    const compiledEntries = filesData.filter((e) => e.kind === "compiled");

    const snapshots = [];

    compiledEntries.forEach((entry) => {
      const { trainings, matrix } = parseCompiledFile(entry);
      const date = parseDateLabel(entry.dateLabel) || entry.dateGuess || null;
      snapshots.push({
        label: entry.dateLabel.trim() || entry.name,
        date,
        trainings,
        matrix,
        totalCount: matrix.length,
        sourceName: entry.name,
        isRawCompile: false,
      });
    });

    let rawResult = null;
    if (rawEntries.length > 0) {
      rawResult = compileRawFiles(rawEntries);
      const label = newCompileDateInput.value.trim() || formatDateLabel(new Date());
      snapshots.push({
        label,
        date: parseDateLabel(label),
        trainings: rawResult.trainings,
        matrix: rawResult.matrix,
        totalCount: rawResult.matrix.length,
        sourceName: "New compile",
        isRawCompile: true,
      });
    }

    snapshots.sort((a, b) => {
      if (a.date && b.date) return a.date - b.date;
      if (a.date) return -1;
      if (b.date) return 1;
      return 0;
    });

    const primary = rawResult
      ? snapshots.find((s) => s.isRawCompile)
      : snapshots[snapshots.length - 1];

    return {
      primary,
      snapshots,
      hasOld: compiledEntries.length > 0,
      skippedRows: rawResult ? rawResult.skippedRows : 0,
      fileCount: filesData.length,
    };
  }

  function renderPreview(result) {
    const { primary, snapshots, hasOld, skippedRows, fileCount } = result;
    const { trainings, matrix } = primary;

    previewTable.innerHTML = "";
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    [
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
    ].forEach((label) => {
      const th = document.createElement("th");
      th.textContent = label;
      headRow.appendChild(th);
    });
    thead.appendChild(headRow);
    previewTable.appendChild(thead);

    const tbody = document.createElement("tbody");
    matrix.forEach((person) => {
      const tr = document.createElement("tr");
      const tdUserId = document.createElement("td");
      tdUserId.textContent = person.userId;
      tr.appendChild(tdUserId);
      const tdFirst = document.createElement("td");
      tdFirst.textContent = person.firstName;
      tr.appendChild(tdFirst);
      const tdLast = document.createElement("td");
      tdLast.textContent = person.lastName;
      tr.appendChild(tdLast);
      person.cells.forEach((status) => {
        const td = document.createElement("td");
        const tag = document.createElement("span");
        tag.className = `tag ${status.toLowerCase()}`;
        tag.textContent = status;
        td.appendChild(tag);
        tr.appendChild(td);
      });

      [
        person.percentComplete,
        person.managerId,
        person.managerFirstName,
        person.managerLastName,
        person.organizationId,
        person.legalEntity,
      ].forEach((value) => {
        const td = document.createElement("td");
        td.textContent = value;
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });
    previewTable.appendChild(tbody);

    let hint = primary.isRawCompile
      ? `${matrix.length} people, ${trainings.length} training(s) compiled from ${fileCount} file(s).`
      : `${matrix.length} people, ${trainings.length} training(s) shown from past compilation "${primary.sourceName}".`;
    if (skippedRows > 0) hint += ` (${skippedRows} row(s) skipped for missing name or training data.)`;
    hint += ` ${snapshots.length} snapshot(s) will appear in the Compliance History sheet.`;
    const diffPair = hasOld ? pickDiffPair(snapshots) : null;
    if (diffPair) {
      hint += ` A Roster Changes sheet will compare "${diffPair.oldSnap.label}" against "${diffPair.newSnap.label}".`;
    }
    summaryHint.textContent = hint;

    previewPanel.hidden = false;
    exportPanel.hidden = false;
  }

  compileBtn.addEventListener("click", () => {
    if (filesData.length === 0) return;
    try {
      compiled = buildSnapshots();
      renderPreview(compiled);
      setStatus(`Compiled successfully.`);
    } catch (err) {
      setStatus(`Error compiling files: ${err.message}`, true);
    }
  });

  // --- Export -----------------------------------------------------------------

  const STATUS_COLORS = {
    Complete: { font: "FF1F9D55", fill: "FFE6F7EC" },
    Incomplete: { font: "FFC0392B", fill: "FFFDECEA" },
    Unassigned: { font: "FF8A94A3", fill: "FFEEF0F3" },
  };

  function styleHeaderRow(row, color) {
    row.font = { bold: true, color: { argb: "FFFFFFFF" } };
    row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: color || "FF2F6FED" } };
    row.alignment = { vertical: "middle", horizontal: "center" };
  }

  function applyBorders(sheet) {
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
  }

  function colorTrainingCells(sheet, firstDataRow, lastDataRow, firstTrainingCol, lastTrainingCol) {
    for (let r = firstDataRow; r <= lastDataRow; r++) {
      const row = sheet.getRow(r);
      for (let c = firstTrainingCol; c <= lastTrainingCol; c++) {
        const cell = row.getCell(c);
        const colors = STATUS_COLORS[cell.value];
        if (colors) {
          cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: colors.fill } };
          cell.font = { color: { argb: colors.font }, bold: true };
        }
        cell.alignment = { horizontal: "center" };
      }
    }
  }

  function addRosterSheet(workbook, sheetName, trainings, matrix) {
    const sheet = workbook.addWorksheet(sheetName, {
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

    styleHeaderRow(sheet.getRow(1));
    sheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: headerLabels.length },
    };

    const leadingCols = 3; // User ID, First Name, Last Name
    const lastTrainingCol = leadingCols + trainings.length;
    colorTrainingCells(sheet, 2, sheet.rowCount, leadingCols + 1, lastTrainingCol);
    for (let r = 2; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      for (let c = lastTrainingCol + 1; c <= headerLabels.length; c++) {
        row.getCell(c).alignment = { horizontal: "center" };
      }
      for (let c = 1; c <= leadingCols; c++) {
        row.getCell(c).alignment = { horizontal: "left" };
      }
    }

    applyBorders(sheet);
    return sheet;
  }

  // Union of every training name that appears in any of the given snapshots,
  // in first-seen order.
  function collectTrainingUnion(snapshots) {
    const seen = [];
    const set = new Set();
    snapshots.forEach((snap) => {
      snap.trainings.forEach((t) => {
        if (!set.has(t)) {
          set.add(t);
          seen.push(t);
        }
      });
    });
    return seen;
  }

  // Builds the "Compliance History" sheet: one row per snapshot (oldest to
  // newest), with a Complete count and % Complete for each training.
  function addHistorySheet(workbook, snapshots) {
    const trainings = collectTrainingUnion(snapshots);
    const sheet = workbook.addWorksheet("Compliance History");

    const latest = snapshots[snapshots.length - 1];
    const totalCols = 1 + trainings.length * 2;

    const row1 = sheet.getRow(1);
    row1.getCell(1).value = `${latest.totalCount} Trainees`;
    trainings.forEach((training, i) => {
      const startCol = 2 + i * 2;
      sheet.mergeCells(1, startCol, 1, startCol + 1);
      row1.getCell(startCol).value = training;
    });

    const row2 = sheet.getRow(2);
    trainings.forEach((training, i) => {
      const startCol = 2 + i * 2;
      row2.getCell(startCol).value = "Complete";
      row2.getCell(startCol + 1).value = "% Complete";
    });

    snapshots.forEach((snap) => {
      const rowValues = [snap.label];
      trainings.forEach((training) => {
        const idx = snap.trainings.indexOf(training);
        if (idx === -1) {
          rowValues.push("", "");
          return;
        }
        const completeCount = snap.matrix.filter((p) => p.cells[idx] === "Complete").length;
        const pct = snap.totalCount > 0 ? completeCount / snap.totalCount : 0;
        rowValues.push(completeCount, pct);
      });
      sheet.addRow(rowValues);
    });

    for (let i = 0; i < trainings.length; i++) {
      const pctCol = 2 + i * 2 + 1;
      for (let r = 3; r <= sheet.rowCount; r++) {
        const cell = sheet.getRow(r).getCell(pctCol);
        if (typeof cell.value === "number") cell.numFmt = "0.00%";
      }
    }

    styleHeaderRow(row1);
    styleHeaderRow(row2);
    sheet.getColumn(1).width = 16;
    for (let c = 2; c <= totalCols; c++) sheet.getColumn(c).width = 13;

    for (let r = 1; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      row.getCell(1).alignment = { horizontal: "left", vertical: "middle" };
      for (let c = 2; c <= totalCols; c++) {
        row.getCell(c).alignment = { horizontal: "center", vertical: "middle" };
      }
    }

    applyBorders(sheet);
    return sheet;
  }

  // Picks the two snapshots to diff: a freshly compiled raw snapshot is
  // always "new" (it's the current, live pull) regardless of what date its
  // source filenames happen to carry; "old" is the earliest past
  // compilation. With no raw data, falls back to oldest vs newest compiled
  // snapshot overall.
  function pickDiffPair(snapshots) {
    if (snapshots.length < 2) return null;
    const compiledSnapshots = snapshots.filter((s) => !s.isRawCompile);
    const rawSnapshot = snapshots.find((s) => s.isRawCompile);
    if (rawSnapshot) {
      if (compiledSnapshots.length === 0) return null;
      return { oldSnap: compiledSnapshots[0], newSnap: rawSnapshot };
    }
    return { oldSnap: compiledSnapshots[0], newSnap: compiledSnapshots[compiledSnapshots.length - 1] };
  }

  function buildRosterIndex(snapshot) {
    const map = new Map();
    snapshot.matrix.forEach((person) => {
      if (person.key) map.set(person.key, person);
    });
    return map;
  }

  // Builds the "Roster Changes" sheet: every person present in only the
  // oldest snapshot, or only in the newest snapshot, not both.
  function addDiffSheet(workbook, oldSnap, newSnap) {
    const trainings = collectTrainingUnion([oldSnap, newSnap]);
    const oldIndex = buildRosterIndex(oldSnap);
    const newIndex = buildRosterIndex(newSnap);

    const rows = [];
    oldIndex.forEach((person, key) => {
      if (!newIndex.has(key)) {
        rows.push({ person, snap: oldSnap, status: `Only in ${oldSnap.label}` });
      }
    });
    newIndex.forEach((person, key) => {
      if (!oldIndex.has(key)) {
        rows.push({ person, snap: newSnap, status: `Only in ${newSnap.label}` });
      }
    });

    rows.sort((a, b) => {
      if (a.status !== b.status) return a.status.localeCompare(b.status);
      return a.person.lastName.localeCompare(b.person.lastName, undefined, { sensitivity: "base" });
    });

    const sheet = workbook.addWorksheet("Roster Changes", {
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
      "Snapshot Date",
      "Status",
    ];
    sheet.columns = headerLabels.map((label) => ({
      header: label,
      key: label,
      width: Math.min(Math.max(label.length + 4, 14), 40),
    }));

    rows.forEach(({ person, snap, status }) => {
      const cells = trainings.map((training) => {
        const idx = snap.trainings.indexOf(training);
        return idx === -1 ? "" : person.cells[idx];
      });
      sheet.addRow([
        person.userId,
        person.firstName,
        person.lastName,
        ...cells,
        person.percentComplete,
        person.managerId,
        person.managerFirstName,
        person.managerLastName,
        person.organizationId,
        person.legalEntity,
        snap.label,
        status,
      ]);
    });

    styleHeaderRow(sheet.getRow(1));
    sheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: headerLabels.length },
    };

    const leadingCols = 3;
    const lastTrainingCol = leadingCols + trainings.length;
    colorTrainingCells(sheet, 2, sheet.rowCount, leadingCols + 1, lastTrainingCol);
    for (let r = 2; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      for (let c = lastTrainingCol + 1; c <= headerLabels.length; c++) {
        row.getCell(c).alignment = { horizontal: "center" };
      }
      for (let c = 1; c <= leadingCols; c++) {
        row.getCell(c).alignment = { horizontal: "left" };
      }
    }

    applyBorders(sheet);
    return sheet;
  }

  async function exportWorkbook() {
    if (!compiled) return;
    const { primary, snapshots, hasOld } = compiled;

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Training Compliance Compiler";
    workbook.created = new Date();

    addRosterSheet(workbook, "Training Roster", primary.trainings, primary.matrix);

    if (snapshots.length > 0) {
      addHistorySheet(workbook, snapshots);
    }

    const diffPair = hasOld ? pickDiffPair(snapshots) : null;
    if (diffPair) {
      addDiffSheet(workbook, diffPair.oldSnap, diffPair.newSnap);
    }

    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([buffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    const stamp = new Date().toISOString().slice(0, 10);
    a.href = url;
    a.download = `training_compliance_${stamp}.xlsx`;
    document.body.appendChild(a);
    a.click();
    a.remove();
    URL.revokeObjectURL(url);
  }

  exportBtn.addEventListener("click", () => {
    exportWorkbook().catch((err) => setStatus(`Export failed: ${err.message}`, true));
  });
})();
