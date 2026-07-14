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

  // Each entry: { name, rows: [ {header: value, ...} ], normMap: Map(normalizedHeader -> originalHeader) }
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
        filesData.push({
          name: file.name,
          rows,
          normMap: buildNormMap(rows),
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

  // --- Compile ---------------------------------------------------------------

  function compileFiles() {
    const people = new Map(); // personKey -> { firstName, lastName }
    const trainingSet = new Set();
    // personKey -> Map(trainingName -> { allComplete: bool })
    const statusMap = new Map();
    // personKey -> { managerId, managerFirstName, managerLastName, organizationId }
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
        const trainingName = rowTrainingName(row, normMap);
        if (!trainingName) {
          skippedRows++;
          continue;
        }

        const firstName = collapseWhitespace(rawFirst);
        const lastName = collapseWhitespace(rawLast);
        const personKey = `${firstName.toLowerCase()}|${lastName.toLowerCase()}`;

        if (!people.has(personKey)) {
          people.set(personKey, { firstName, lastName });
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
          });
        }
        const info = personInfo.get(personKey);
        setIfBlank(info, "managerId", getFirstVal(row, normMap, ["manager id", "manager id (user)"]));
        setIfBlank(info, "managerFirstName", getFirstVal(row, normMap, ["first name manager", "manager first name"]));
        setIfBlank(info, "managerLastName", getFirstVal(row, normMap, ["last name manager", "manager last name"]));
        setIfBlank(info, "organizationId", getFirstVal(row, normMap, ["organization id"]));
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
        firstName: person.firstName,
        lastName: person.lastName,
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

  function renderPreview(result) {
    const { trainings, matrix, skippedRows, fileCount } = result;

    previewTable.innerHTML = "";
    const thead = document.createElement("thead");
    const headRow = document.createElement("tr");
    [
      "First Name",
      "Last Name",
      ...trainings,
      "Percent Complete",
      "Manager ID",
      "Manager First Name",
      "Manager Last Name",
      "Organization ID",
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
      ].forEach((value) => {
        const td = document.createElement("td");
        td.textContent = value;
        tr.appendChild(td);
      });

      tbody.appendChild(tr);
    });
    previewTable.appendChild(tbody);

    summaryHint.textContent =
      `${matrix.length} people, ${trainings.length} training(s) compiled from ${fileCount} file(s).` +
      (skippedRows > 0 ? ` (${skippedRows} row(s) skipped for missing name or training data.)` : "");

    previewPanel.hidden = false;
    exportPanel.hidden = false;
  }

  compileBtn.addEventListener("click", () => {
    if (filesData.length === 0) return;
    try {
      compiled = compileFiles();
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

  async function exportWorkbook() {
    if (!compiled) return;
    const { trainings, matrix } = compiled;

    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Training Compliance Compiler";
    workbook.created = new Date();

    const sheet = workbook.addWorksheet("Training Roster", {
      views: [{ state: "frozen", ySplit: 1 }],
    });

    const headerLabels = [
      "First Name",
      "Last Name",
      ...trainings,
      "Percent Complete",
      "Manager ID",
      "Manager First Name",
      "Manager Last Name",
      "Organization ID",
    ];
    sheet.columns = headerLabels.map((label) => ({
      header: label,
      key: label,
      width: Math.min(Math.max(label.length + 4, 14), 40),
    }));

    matrix.forEach((person) => {
      sheet.addRow([
        person.firstName,
        person.lastName,
        ...person.cells,
        person.percentComplete,
        person.managerId,
        person.managerFirstName,
        person.managerLastName,
        person.organizationId,
      ]);
    });

    const headerRow = sheet.getRow(1);
    headerRow.font = { bold: true, color: { argb: "FFFFFFFF" } };
    headerRow.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FF2F6FED" },
    };
    headerRow.alignment = { vertical: "middle", horizontal: "center" };

    sheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: 1, column: headerLabels.length },
    };

    const lastTrainingCol = 2 + trainings.length;
    for (let r = 2; r <= sheet.rowCount; r++) {
      const row = sheet.getRow(r);
      for (let c = 3; c <= lastTrainingCol; c++) {
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
      row.getCell(1).alignment = { horizontal: "left" };
      row.getCell(2).alignment = { horizontal: "left" };
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
