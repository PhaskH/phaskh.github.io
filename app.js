const STORAGE_KEYS = {
  builds: "mhn_builds_v1",
  weapons: "mhn_weapons_v1",
  selectedBuildId: "mhn_selected_build_v1",
  selectedWeaponId: "mhn_selected_weapon_v1",
  uptimes: "mhn_uptimes_v1",
};

const state = {
  data: null,
  builds: [],
  weapons: [],
  selectedBuildId: null,
  selectedWeaponId: null,
  editingBuildId: null,
  editingWeaponId: null,
  buildDraft: null,
  weaponDraft: null,
  matrixComparison: null,
  buildEditorColumnCount: 1,
  uptimeFields: [],
  uptimeValues: {},
  uptimeDraft: null,
};

const els = {
  buildList: document.getElementById("build-list"),
  weaponList: document.getElementById("weapon-list"),
  buildForm: document.getElementById("build-form"),
  weaponForm: document.getElementById("weapon-form"),
  calculatorBuild: document.getElementById("calculator-build"),
  calculatorWeapon: document.getElementById("calculator-weapon"),
  resultGrid: document.getElementById("result-grid"),
  calculatorActions: document.getElementById("calculator-actions"),
  selectionActions: document.getElementById("selection-actions"),
  skillSummaryList: document.getElementById("skill-summary-list"),
  newBuild: document.getElementById("new-build"),
  newWeapon: document.getElementById("new-weapon"),
  exportSelectedBuilds: document.getElementById("export-selected-builds"),
  exportSelectedWeapons: document.getElementById("export-selected-weapons"),
  exportData: document.getElementById("export-data"),
  importData: document.getElementById("import-data"),
  riftModal: document.getElementById("rift-modal"),
  modalTitle: document.getElementById("modal-title"),
  closeRiftModal: document.getElementById("close-rift-modal"),
  riftModalContent: document.getElementById("rift-modal-content"),
};

function scrollEditorIntoView(editorForm) {
  const panel = editorForm?.closest(".panel");
  if (!panel) {
    return;
  }
  requestAnimationFrame(() => {
    panel.scrollIntoView({
      behavior: "smooth",
      block: "start",
    });
  });
}

function syncBuildEditorSeparators() {
  const grid = els.buildForm?.querySelector(".build-editor-grid");
  if (!grid) {
    return;
  }

  grid.querySelectorAll(".editor-column-separator").forEach((separator) => separator.remove());

  const rows = Array.from(grid.querySelectorAll(".editor-row"));
  if (rows.length < 2) {
    return;
  }

  const top = rows.reduce((value, row) => Math.min(value, row.offsetTop), Number.POSITIVE_INFINITY);
  const bottom = rows.reduce(
    (value, row) => Math.max(value, row.offsetTop + row.offsetHeight),
    Number.NEGATIVE_INFINITY,
  );

  const columnStarts = [...new Set(rows.map((row) => row.offsetLeft))].sort((a, b) => a - b);
  if (columnStarts.length < 2) {
    return;
  }

  for (let index = 1; index < columnStarts.length; index += 1) {
    const separator = document.createElement("div");
    separator.className = "editor-column-separator";
    separator.style.left = `${columnStarts[index] - 6}px`;
    separator.style.top = `${top}px`;
    separator.style.height = `${bottom - top}px`;
    grid.append(separator);
  }
}

function getBuildEditorColumnCount() {
  if (window.innerWidth <= 780) {
    return 1;
  }

  const gridWidth = els.buildForm?.clientWidth || els.buildForm?.closest(".panel")?.clientWidth || 0;
  const minColumnWidth = 260;
  const gap = 10;
  if (!gridWidth) {
    return 1;
  }

  return Math.max(1, Math.floor((gridWidth + gap) / (minColumnWidth + gap)));
}

function orderBuildFieldsByVisibleColumn(fields) {
  const columnCount = getBuildEditorColumnCount();
  if (columnCount <= 1) {
    return fields;
  }

  const rowCount = Math.ceil(fields.length / columnCount);
  const ordered = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const sourceIndex = columnIndex * rowCount + rowIndex;
      if (fields[sourceIndex]) {
        ordered.push(fields[sourceIndex]);
      }
    }
  }

  return ordered;
}

function deepClone(value) {
  if (typeof structuredClone === "function") {
    return structuredClone(value);
  }
  return JSON.parse(JSON.stringify(value));
}

function makeId() {
  return `${Date.now()}_${Math.random().toString(36).slice(2, 10)}`;
}

function cellRefToCoords(ref) {
  const match = /^([A-Z]{1,3})(\d+)$/.exec(ref);
  if (!match) {
    throw new Error(`Bad cell ref: ${ref}`);
  }

  const colLetters = match[1];
  const row = Number(match[2]) - 1;
  let col = 0;
  for (const ch of colLetters) {
    col = col * 26 + (ch.charCodeAt(0) - 64);
  }
  return { row, col: col - 1 };
}

function toAddress(sheetId, ref) {
  const { row, col } = cellRefToCoords(ref);
  return { sheet: sheetId, row, col };
}

function createEngine() {
  return HyperFormula.buildFromSheets(deepClone(state.data.sheets), {
    licenseKey: "gpl-v3",
  });
}

function readCell(engine, sheetName, ref) {
  const sheetId = engine.getSheetId(sheetName);
  return engine.getCellValue(toAddress(sheetId, ref));
}

function writeCell(engine, sheetName, ref, value) {
  const sheetId = engine.getSheetId(sheetName);
  engine.setCellContents(toAddress(sheetId, ref), [[value]]);
}

function isErrorValue(value) {
  return value && typeof value === "object" && value.type;
}

function formatResult(value) {
  if (isErrorValue(value)) {
    return value.value || value.type || "Error";
  }
  if (typeof value === "number") {
    return value.toFixed(2);
  }
  return String(value ?? "");
}

function formatSignedPercent(value) {
  if (!Number.isFinite(value)) {
    return "N/A";
  }
  if (value === 0) {
    return "0.00%";
  }
  return `${value > 0 ? "+" : ""}${value.toFixed(2)}%`;
}

function getDeltaBackground(percent) {
  if (!Number.isFinite(percent) || percent === 0) {
    return "rgba(255, 255, 255, 0.03)";
  }
  const intensity = Math.min(Math.abs(percent), 10) / 10;
  if (percent > 0) {
    return `rgba(83, 179, 125, ${0.14 + intensity * 0.34})`;
  }
  return `rgba(210, 100, 100, ${0.14 + intensity * 0.34})`;
}

function loadStoredItems(key) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return [];
    }
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

function saveStoredItems(key, items) {
  localStorage.setItem(key, JSON.stringify(items));
}

function loadStoredValue(key) {
  try {
    return localStorage.getItem(key);
  } catch {
    return null;
  }
}

function saveStoredValue(key, value) {
  localStorage.setItem(key, value);
}

function loadStoredObject(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    if (!raw) {
      return fallback;
    }
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? parsed : fallback;
  } catch {
    return fallback;
  }
}

function exportAppData() {
  return JSON.stringify(
    {
      version: 1,
      builds: state.builds,
      weapons: state.weapons,
    },
    null,
    0,
  );
}

function exportScopedData({ builds = [], weapons = [] }) {
  return JSON.stringify(
    {
      version: 1,
      builds,
      weapons,
    },
    null,
    0,
  );
}

function normalizeImportedBuild(item) {
  const base = buildDefaultBuild();
  return {
    ...base,
    ...item,
    id: makeId(),
    name: String(item?.name || "Imported Build"),
    compareEnabled: typeof item?.compareEnabled === "boolean" ? item.compareEnabled : true,
    values: {
      ...base.values,
      ...(item?.values && typeof item.values === "object" ? item.values : {}),
    },
  };
}

function normalizeImportedWeapon(item) {
  const base = buildDefaultWeapon();
  return {
    ...base,
    ...item,
    id: makeId(),
    name: String(item?.name || "Imported Weapon"),
    compareEnabled: typeof item?.compareEnabled === "boolean" ? item.compareEnabled : true,
    values: {
      ...base.values,
      ...(item?.values && typeof item.values === "object" ? item.values : {}),
    },
    isRift: Boolean(item?.isRift),
  };
}

function stableStringify(value) {
  if (Array.isArray(value)) {
    return `[${value.map(stableStringify).join(",")}]`;
  }
  if (value && typeof value === "object") {
    const entries = Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`);
    return `{${entries.join(",")}}`;
  }
  return JSON.stringify(value);
}

function buildImportFingerprint(item) {
  return stableStringify({
    name: item.name,
    values: item.values,
  });
}

function weaponImportFingerprint(item) {
  return stableStringify({
    name: item.name,
    values: item.values,
    isRift: Boolean(item.isRift),
  });
}

function collectDefaultUptimeFields() {
  const sheet = state.data?.sheets?.Calculator1;
  if (!sheet) {
    return [];
  }
  const engine = createEngine();
  const getDefaultValue = (ref, fallback = 0) => {
    const value = readCell(engine, "Calculator1", ref);
    return typeof value === "number" && Number.isFinite(value) ? value : fallback;
  };

  const fields = [];
  const remainingHealthLabel = sheet[14]?.[3];
  fields.push({
    ref: "E15",
    label: String(remainingHealthLabel || "Remaining Health"),
    defaultValue: 100,
    displayScale: 1,
  });

  for (let rowIndex = 23; rowIndex <= 46; rowIndex += 1) {
    const label = sheet[rowIndex]?.[3];
    const defaultValue = Number(sheet[rowIndex]?.[4]);
    if (!label) {
      continue;
    }
    fields.push({
      ref: `E${rowIndex + 1}`,
      label: String(label),
      defaultValue: Number.isFinite(defaultValue) ? defaultValue : 0,
      displayScale: 100,
    });
  }

  [
    { ref: "H68", label: "Resus BE" },
    { ref: "I68", label: "Coal BE" },
    { ref: "H69", label: "Resus BD" },
    { ref: "I69", label: "Coal BD" },
  ].forEach((field) => {
    fields.push({
      ref: field.ref,
      label: field.label,
      defaultValue: getDefaultValue(field.ref),
      displayScale: 100,
    });
  });

  return fields;
}

function buildDefaultUptimeValues() {
  return Object.fromEntries(
    state.uptimeFields.map((field) => [field.ref, field.defaultValue]),
  );
}

function isUsingDefaultUptime(field, value) {
  return Math.abs((value ?? 0) - (field?.defaultValue ?? 0)) < 0.000001;
}

function persistUptimes() {
  localStorage.setItem(STORAGE_KEYS.uptimes, JSON.stringify(state.uptimeValues));
}

async function copyTextToClipboard(text) {
  if (navigator.clipboard?.writeText) {
    await navigator.clipboard.writeText(text);
    return;
  }

  const input = document.createElement("textarea");
  input.value = text;
  input.setAttribute("readonly", "");
  input.style.position = "absolute";
  input.style.left = "-9999px";
  document.body.append(input);
  input.select();
  document.execCommand("copy");
  input.remove();
}

function valueFromDefault(rawValue, fallback) {
  if (typeof fallback === "number" && typeof rawValue === "string" && /^-?\d+(\.\d+)?$/.test(rawValue)) {
    return Number(rawValue);
  }
  return rawValue;
}

function normalizeWeaponValue(ref, rawValue, fallback) {
  if (ref === "E6") {
    const numeric = Number(rawValue);
    return Number.isFinite(numeric) ? numeric / 100 : fallback;
  }
  return valueFromDefault(rawValue, fallback);
}

function displayWeaponValue(ref, value) {
  if (ref === "E6") {
    return String(Number(value ?? 0) * 100);
  }
  return String(value ?? "");
}

function schemaMap(schema) {
  return Object.fromEntries(schema.map((field) => [field.ref, field]));
}

function buildSchemaMap() {
  return schemaMap(state.data.buildFields);
}

function weaponSchemaMap() {
  return schemaMap(state.data.weaponFields);
}

function buildDefaultBuild() {
  return {
    id: makeId(),
    name: "Default Build",
    compareEnabled: true,
    values: Object.fromEntries(
      state.data.buildFields.map((field) => [field.ref, 0]),
    ),
  };
}

function buildDefaultWeapon() {
  const values = Object.fromEntries(
    state.data.weaponFields.map((field) => [field.ref, 0]),
  );
  values.E5 = "Raw";
  values.E7 = "Lance";
  return {
    id: makeId(),
    name: "Default Weapon",
    values,
    isRift: false,
    compareEnabled: true,
  };
}

function getBuildById(id) {
  return state.builds.find((item) => item.id === id) ?? null;
}

function getWeaponById(id) {
  const item = state.weapons.find((weapon) => weapon.id === id) ?? null;
  if (!item) {
    return null;
  }
  if (typeof item.isRift !== "boolean") {
    item.isRift = false;
  }
  return item;
}

function applyScenario(engine, build, weapon) {
  for (const field of state.data.buildFields) {
    writeCell(engine, "Calculator1", field.ref, build.values[field.ref]);
  }
  for (const field of state.data.weaponFields) {
    writeCell(engine, "Calculator1", field.ref, weapon.values[field.ref]);
  }
  for (const field of state.uptimeFields) {
    writeCell(engine, "Calculator1", field.ref, state.uptimeValues[field.ref]);
  }
}

function getBuildLabels(build, weapon) {
  const engine = createEngine();
  applyScenario(engine, build, weapon);
  const labels = {};
  for (const field of state.data.buildFields) {
    labels[field.ref] = readCell(engine, "Calculator1", field.labelRef);
    if (labels[field.ref] === "Vital Fire") {
      labels[field.ref] = "Vital Element";
    }
  }
  labels.B3 = "Elemental Attack";
  labels.B4 = "Adv. Elemental Attack";
  return labels;
}

function getWeaponLabels(weapon) {
  const engine = createEngine();
  applyScenario(engine, state.buildDraft ?? buildDefaultBuild(), weapon);
  const labels = {};
  for (const field of state.data.weaponFields) {
    labels[field.ref] = readCell(engine, "Calculator1", field.labelRef);
  }
  return labels;
}

function calculateSelectedScenario() {
  const build = getBuildById(state.selectedBuildId);
  const weapon = getWeaponById(state.selectedWeaponId);
  if (!build || !weapon) {
    return null;
  }

  const engine = createEngine();
  applyScenario(engine, build, weapon);
  return {
    h12: readCell(engine, "Calculator1", state.data.resultCell),
  };
}

window.__mhnDebugScenario = function __mhnDebugScenario() {
  const build = getBuildById(state.selectedBuildId);
  const weapon = getWeaponById(state.selectedWeaponId);
  if (!build || !weapon) {
    return null;
  }
  const engine = createEngine();
  applyScenario(engine, build, weapon);
  const refs = [
    "H12",
    "BO71",
    "BO67",
    "BO68",
    "BO69",
    "BC67",
    "BF67",
    "BF68",
    "BF71",
    "BF72",
    "AK89",
    "S67",
    "S68",
    "AO68",
    "AX86",
    "AX87",
    "AX88",
  ];
  return Object.fromEntries(refs.map((ref) => [ref, readCell(engine, "Calculator1", ref)]));
};

function renderCalculatorSelectors() {
  const buildOptions = state.builds
    .map(
      (item) =>
        `<option value="${item.id}" ${item.id === state.selectedBuildId ? "selected" : ""}>${escapeHtml(item.name)}</option>`,
    )
    .join("");
  const weaponOptions = state.weapons
    .map(
      (item) =>
        `<option value="${item.id}" ${item.id === state.selectedWeaponId ? "selected" : ""}>${escapeHtml(item.name)}</option>`,
    )
    .join("");

  els.calculatorBuild.innerHTML = buildOptions;
  els.calculatorWeapon.innerHTML = weaponOptions;
}

function renderResultGrid() {
  const result = calculateSelectedScenario();
  if (!result) {
    els.resultGrid.innerHTML = `<div class="empty-state">Create and select at least one build and one weapon.</div>`;
    return;
  }

  els.resultGrid.innerHTML = `
    <div class="result-card">
      <div class="label">Effective Damage</div>
      <div class="value">${escapeHtml(formatResult(result.h12))}</div>
    </div>
  `;
}

function renderCalculatorActions() {
  const selectedWeaponId = els.calculatorWeapon.value || state.selectedWeaponId;
  const weapon = getWeaponById(selectedWeaponId);
  const actions = [
    `<button id="compare-rift" type="button" class="${weapon?.isRift ? "" : "button-disabled"}">Compare Rift Combinations</button>`,
  ];

  els.calculatorActions.innerHTML = actions.join("");
  els.calculatorActions.querySelector("#compare-rift").addEventListener("click", () => {
    if (weapon?.isRift) {
      openRiftComparison();
      return;
    }
    openRiftUnavailableMessage();
  });
}

function renderSelectionActions() {
  els.selectionActions.innerHTML = `
    <div class="selection-actions-row">
      <button id="compare-build-weapon-matrix" type="button">Compare Selected Builds & Weapons</button>
      <div class="selection-actions-help">Select the builds and weapons you want to compare below, then click the button.</div>
    </div>
    <div class="selection-actions-row">
      <button id="view-edit-uptimes" type="button">View/Edit Uptimes</button>
    </div>
  `;
  els.selectionActions
    .querySelector("#compare-build-weapon-matrix")
    .addEventListener("click", () => {
      openBuildWeaponComparison();
    });
  els.selectionActions
    .querySelector("#view-edit-uptimes")
    .addEventListener("click", () => {
      openUptimesModal();
    });
}

function openRiftUnavailableMessage() {
  els.modalTitle.textContent = "Rift Combinations";
  els.riftModal.querySelector(".modal-window")?.classList.add("modal-window-medium");
  els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-compact");
  els.riftModalContent.innerHTML = `
    <div class="comparison-intro">This option is only available with rift base weapons.</div>
  `;
  els.riftModal.classList.remove("hidden");
}

function openUptimesModal() {
  state.uptimeDraft = { ...state.uptimeValues };
  renderUptimesModal();
}

function renderUptimesModal() {
  if (!state.uptimeDraft) {
    return;
  }

  els.modalTitle.textContent = "View/Edit Uptimes";
  els.riftModal.querySelector(".modal-window")?.classList.add("modal-window-uptime");
  els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-medium");
  els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-compact");
  els.riftModalContent.innerHTML = `
    <div class="editor-actions-row uptime-actions-row">
      <button type="button" id="save-uptimes">Save Values</button>
      <button class="secondary" type="button" id="revert-uptimes">Revert to Default (KreaTV1 sheet)</button>
      <button class="secondary" type="button" id="discard-uptimes">Discard</button>
    </div>
    <div class="uptime-list">
      ${state.uptimeFields
        .map((field) => `
          <label class="uptime-row" for="uptime-${field.ref}">
            <span class="uptime-label">${escapeHtml(field.label)}</span>
            <input id="uptime-${field.ref}" type="number" min="${field.displayScale === 1 ? "1" : "0"}" max="100" step="${field.displayScale === 1 ? "1" : "0.1"}" data-uptime-field="${field.ref}" value="${escapeHtml(
              field.displayScale === 1
                ? String(Math.round(state.uptimeDraft[field.ref] ?? 0))
                : ((state.uptimeDraft[field.ref] ?? 0) * (field.displayScale ?? 100)).toFixed(1),
            )}" />
            <span class="uptime-unit">${field.displayScale === 1 ? "" : "%"}</span>
            <span class="uptime-default-indicator" data-uptime-default-indicator="${field.ref}">${isUsingDefaultUptime(field, state.uptimeDraft[field.ref]) ? "Using Krea default" : ""}</span>
          </label>
        `)
        .join("")}
    </div>
  `;

  els.riftModalContent.querySelectorAll("[data-uptime-field]").forEach((input) => {
    input.addEventListener("input", (event) => {
      const ref = event.target.dataset.uptimeField;
      const field = state.uptimeFields.find((item) => item.ref === ref);
      const numeric = Number(event.target.value);
      const minValue = field?.displayScale === 1 ? 1 : 0;
      const fallbackValue = field?.displayScale === 1 ? 1 : 0;
      const boundedValue = Number.isFinite(numeric) ? Math.min(100, Math.max(minValue, numeric)) : fallbackValue;
      state.uptimeDraft[ref] = boundedValue / (field?.displayScale ?? 100);
      const indicator = els.riftModalContent.querySelector(`[data-uptime-default-indicator="${ref}"]`);
      if (indicator) {
        indicator.textContent = isUsingDefaultUptime(field, state.uptimeDraft[ref]) ? "Using Krea default" : "";
      }
    });
  });

  els.riftModalContent.querySelector("#save-uptimes").addEventListener("click", () => {
    state.uptimeValues = { ...state.uptimeDraft };
    persistUptimes();
    state.uptimeDraft = null;
    closeModal();
    renderAll();
  });

  els.riftModalContent.querySelector("#revert-uptimes").addEventListener("click", () => {
    state.uptimeDraft = buildDefaultUptimeValues();
    renderUptimesModal();
  });

  els.riftModalContent.querySelector("#discard-uptimes").addEventListener("click", () => {
    closeModal();
  });

  els.riftModal.classList.remove("hidden");
}

function renderSkillSummary() {
  const build = getBuildById(state.selectedBuildId);
  const weapon = getWeaponById(state.selectedWeaponId);
  if (!build || !weapon) {
    els.skillSummaryList.innerHTML = `<div class="empty-state">No build selected.</div>`;
    return;
  }

  const labels = getBuildLabels(build, weapon);

  const active = state.data.buildFields
    .map((field) => ({
      label: labels[field.ref],
      value: build.values[field.ref],
    }))
    .filter((item) => Number(item.value) !== 0)
    .filter((item) => item.label && item.label !== "-" && item.label !== "/")
    .sort(
      (a, b) =>
        Number(b.value) - Number(a.value) || String(a.label).localeCompare(String(b.label)),
    );

  if (!active.length) {
    els.skillSummaryList.innerHTML = `<div class="empty-state">No active skills in the selected build.</div>`;
    return;
  }

  els.skillSummaryList.className = "skill-summary-list";
  els.skillSummaryList.innerHTML = active
    .map(
      (item) => `
        <div class="skill-summary-item">
          <span>${escapeHtml(String(item.label))}</span>
          <span class="level">Lv ${escapeHtml(String(item.value))}</span>
        </div>
      `,
    )
    .join("");
}

function renderLibraryList(targetEl, items, selectedId, editHandlerName, deleteHandlerName) {
  if (!items.length) {
    targetEl.innerHTML = `<div class="empty-state">Nothing saved yet.</div>`;
    return;
  }

  targetEl.innerHTML = items
    .map(
      (item) => `
        <div class="library-item ${item.id === selectedId ? "selected" : ""}">
          <div class="library-item-main">
            <label class="library-item-check">
              <input type="checkbox" data-action="toggle-compare" data-id="${item.id}" ${item.compareEnabled !== false ? "checked" : ""} />
            </label>
            <div class="library-item-name">${escapeHtml(item.name)}</div>
          </div>
          <div class="library-item-actions">
            <button class="secondary" type="button" data-action="${editHandlerName}" data-id="${item.id}">Edit</button>
            <button class="danger" type="button" data-action="${deleteHandlerName}" data-id="${item.id}">Delete</button>
          </div>
        </div>
      `,
    )
    .join("");
}

function renderBuildList() {
  renderLibraryList(
    els.buildList,
    state.builds,
    state.selectedBuildId,
    "edit-build",
    "delete-build",
  );
}

function renderWeaponList() {
  renderLibraryList(
    els.weaponList,
    state.weapons,
    state.selectedWeaponId,
    "edit-weapon",
    "delete-weapon",
  );
}

function renderBuildForm() {
  if (!state.buildDraft) {
    els.buildForm.classList.add("hidden");
    els.buildForm.innerHTML = "";
    return;
  }

  els.buildForm.classList.remove("hidden");
  state.buildEditorColumnCount = getBuildEditorColumnCount();
  const labels = getBuildLabels(state.buildDraft, buildDefaultWeapon());

  const fieldsMarkup = orderBuildFieldsByVisibleColumn(state.data.buildFields)
    .map((field) => {
      const label = labels[field.ref] ?? field.labelRef;
      const value = state.buildDraft.values[field.ref];
      const isActive = Number(value) !== 0;
      const inputMarkup = field.options
        ? `<select data-build-field="${field.ref}">
            ${field.options
              .map((option) => {
                const selected = String(option) === String(value) ? "selected" : "";
                return `<option value="${escapeHtml(String(option))}" ${selected}>${escapeHtml(String(option))}</option>`;
              })
              .join("")}
          </select>`
        : `<input type="number" step="any" value="${escapeHtml(String(value ?? ""))}" data-build-field="${field.ref}" />`;

      return `
        <div class="editor-row ${isActive ? "editor-row-active" : ""}">
          <label class="editor-label" for="build-${field.ref}">${escapeHtml(String(label))}</label>
          ${inputMarkup.replace("data-build-field", `id="build-${field.ref}" data-build-field`)}
        </div>
      `;
    })
    .join("");

  els.buildForm.innerHTML = `
    <div class="editor-toolbar">
      <label class="field">
        <span>Build Name</span>
        <input id="build-name" type="text" value="${escapeHtml(state.buildDraft.name)}" />
      </label>
      <button type="submit">${state.editingBuildId ? "Save Build" : "Create Build"}</button>
      <button class="secondary" id="duplicate-build-form" type="button">Duplicate</button>
      <button class="secondary" id="discard-build-form" type="button">Discard</button>
    </div>
    <div class="editor-grid build-editor-grid">${fieldsMarkup}</div>
  `;

  els.buildForm.onsubmit = (event) => {
    event.preventDefault();
    saveBuildDraft();
  };

  els.buildForm.querySelector("#build-name").addEventListener("input", (event) => {
    state.buildDraft.name = event.target.value;
  });

  els.buildForm.querySelector("#duplicate-build-form").addEventListener("click", () => {
    duplicateCurrentBuildDraft();
  });

  els.buildForm.querySelector("#discard-build-form").addEventListener("click", () => {
    state.editingBuildId = null;
    state.buildDraft = null;
    renderAll();
  });

  els.buildForm.querySelectorAll("[data-build-field]").forEach((input) => {
    input.addEventListener("change", (event) => {
      const ref = event.target.dataset.buildField;
      const schema = buildSchemaMap()[ref];
      const nextValue = valueFromDefault(event.target.value, schema.defaultValue);
      state.buildDraft.values[ref] = nextValue;
      event.target
        .closest(".editor-row")
        ?.classList.toggle("editor-row-active", Number(nextValue) !== 0);
    });
  });

  syncBuildEditorSeparators();
}

function renderWeaponForm() {
  if (!state.weaponDraft) {
    els.weaponForm.classList.add("hidden");
    els.weaponForm.innerHTML = "";
    return;
  }

  els.weaponForm.classList.remove("hidden");
  const labels = getWeaponLabels(state.weaponDraft);
  const fieldsMarkup = state.data.weaponFields
    .map((field) => {
      const baseLabel = labels[field.ref] ?? field.labelRef;
      const label = field.ref === "E6" ? `${baseLabel} %` : baseLabel;
      const value = state.weaponDraft.values[field.ref];
      const options =
        field.ref === "E5"
          ? (field.options ?? []).filter(
              (option) => !["Poison", "Paralysis", "Sleep", "Blast"].includes(String(option)),
            )
          : field.options;
      let inputMarkup;
      if (options) {
        inputMarkup = `<select data-weapon-field="${field.ref}">
          ${options
            .map((option) => {
              const selected = String(option) === String(value) ? "selected" : "";
              return `<option value="${escapeHtml(String(option))}" ${selected}>${escapeHtml(String(option))}</option>`;
            })
            .join("")}
        </select>`;
      } else {
        const increment = field.ref === "E3" || field.ref === "E4" ? 100 : field.ref === "E6" ? 10 : null;
        const minAttr = field.ref === "E6" ? "" : ' min="0"';
        if (increment) {
          inputMarkup = `
            <div class="weapon-adjuster">
              <input type="number"${minAttr} step="any" value="${escapeHtml(displayWeaponValue(field.ref, value))}" data-weapon-field="${field.ref}" />
              <button class="secondary weapon-adjuster-button" type="button" data-weapon-step="${field.ref}" data-step-direction="-1">-${increment}</button>
              <button class="secondary weapon-adjuster-button" type="button" data-weapon-step="${field.ref}" data-step-direction="1">+${increment}</button>
            </div>
          `;
        } else {
          inputMarkup = `<input type="number"${minAttr} step="any" value="${escapeHtml(displayWeaponValue(field.ref, value))}" data-weapon-field="${field.ref}" />`;
        }
      }

      return `
        <div class="editor-row weapon-row">
          <label class="editor-label" for="weapon-${field.ref}">${escapeHtml(String(label))}</label>
          ${inputMarkup.replace("data-weapon-field", `id="weapon-${field.ref}" data-weapon-field`)}
        </div>
      `;
    })
    .join("");

  els.weaponForm.innerHTML = `
    <div class="editor-toolbar">
      <label class="field">
        <span>Weapon Name</span>
        <input id="weapon-name" type="text" value="${escapeHtml(state.weaponDraft.name)}" />
      </label>
    </div>
    <div class="editor-actions-row">
      <button type="submit">${state.editingWeaponId ? "Save Weapon" : "Create Weapon"}</button>
      <button class="secondary" id="duplicate-weapon-form" type="button">Duplicate</button>
      <button class="secondary" id="discard-weapon-form" type="button">Discard</button>
    </div>
    <div class="checkbox-row">
      <input id="weapon-is-rift" type="checkbox" ${state.weaponDraft.isRift ? "checked" : ""} />
      <label class="editor-label" for="weapon-is-rift">Rift base weapon (no upgrades)</label>
    </div>
    <div class="editor-grid weapon-editor-stack">${fieldsMarkup}</div>
  `;

  els.weaponForm.onsubmit = (event) => {
    event.preventDefault();
    saveWeaponDraft();
  };

  els.weaponForm.querySelector("#weapon-name").addEventListener("input", (event) => {
    state.weaponDraft.name = event.target.value;
  });

  els.weaponForm.querySelector("#duplicate-weapon-form").addEventListener("click", () => {
    duplicateCurrentWeaponDraft();
  });

  els.weaponForm.querySelector("#weapon-is-rift").addEventListener("change", (event) => {
    state.weaponDraft.isRift = event.target.checked;
  });

  els.weaponForm.querySelector("#discard-weapon-form").addEventListener("click", () => {
    state.editingWeaponId = null;
    state.weaponDraft = null;
    renderAll();
  });

  els.weaponForm.querySelectorAll("[data-weapon-field]").forEach((input) => {
    input.addEventListener("change", (event) => {
      const ref = event.target.dataset.weaponField;
      const schema = weaponSchemaMap()[ref];
      state.weaponDraft.values[ref] = normalizeWeaponValue(ref, event.target.value, schema.defaultValue);
      if (ref === "E5") {
        renderWeaponForm();
      }
    });
  });

  els.weaponForm.querySelectorAll("[data-weapon-step]").forEach((button) => {
    button.addEventListener("click", (event) => {
      const ref = event.currentTarget.dataset.weaponStep;
      const direction = Number(event.currentTarget.dataset.stepDirection);
      const amount = ref === "E3" || ref === "E4" ? 100 : ref === "E6" ? 10 : 0;
      const input = els.weaponForm.querySelector(`[data-weapon-field="${ref}"]`);
      const currentValue = Number(input?.value ?? 0);
      const nextValue =
        ref === "E6"
          ? (Number.isFinite(currentValue) ? currentValue : 0) + direction * amount
          : Math.max(0, (Number.isFinite(currentValue) ? currentValue : 0) + direction * amount);
      if (input) {
        input.value = String(nextValue);
      }
      const schema = weaponSchemaMap()[ref];
      state.weaponDraft.values[ref] = normalizeWeaponValue(ref, nextValue, schema.defaultValue);
    });
  });
}

function persistBuilds() {
  saveStoredItems(STORAGE_KEYS.builds, state.builds);
}

function persistWeapons() {
  saveStoredItems(STORAGE_KEYS.weapons, state.weapons);
}

function saveBuildDraft() {
  const cleanName = state.buildDraft.name.trim() || "Unnamed Build";
  const payload = {
    ...state.buildDraft,
    name: cleanName,
  };

  if (state.editingBuildId) {
    state.builds = state.builds.map((item) => (item.id === state.editingBuildId ? payload : item));
  } else {
    payload.id = makeId();
    state.builds = [...state.builds, payload];
    state.selectedBuildId = payload.id;
    state.editingBuildId = payload.id;
  }

  persistBuilds();
  saveStoredValue(STORAGE_KEYS.selectedBuildId, state.selectedBuildId);
  state.editingBuildId = null;
  state.buildDraft = null;
  renderAll();
}

function saveWeaponDraft() {
  const cleanName = state.weaponDraft.name.trim() || "Unnamed Weapon";
  const payload = {
    ...state.weaponDraft,
    name: cleanName,
  };

  if (state.editingWeaponId) {
    state.weapons = state.weapons.map((item) => (item.id === state.editingWeaponId ? payload : item));
  } else {
    payload.id = makeId();
    state.weapons = [...state.weapons, payload];
    state.selectedWeaponId = payload.id;
    state.editingWeaponId = payload.id;
  }

  persistWeapons();
  saveStoredValue(STORAGE_KEYS.selectedWeaponId, state.selectedWeaponId);
  state.editingWeaponId = null;
  state.weaponDraft = null;
  renderAll();
}

function duplicateCurrentBuildDraft() {
  if (!state.buildDraft) {
    return;
  }

  const payload = {
    ...deepClone(state.buildDraft),
    id: makeId(),
    name: `Duplicate of ${state.buildDraft.name || "Unnamed Build"}`,
  };

  state.builds = [...state.builds, payload];
  state.selectedBuildId = payload.id;
  state.editingBuildId = payload.id;
  state.buildDraft = deepClone(payload);
  persistBuilds();
  saveStoredValue(STORAGE_KEYS.selectedBuildId, state.selectedBuildId);
  renderAll();
  scrollEditorIntoView(els.buildForm);
}

function duplicateCurrentWeaponDraft() {
  if (!state.weaponDraft) {
    return;
  }

  const payload = {
    ...deepClone(state.weaponDraft),
    id: makeId(),
    name: `Duplicate of ${state.weaponDraft.name || "Unnamed Weapon"}`,
  };

  state.weapons = [...state.weapons, payload];
  state.selectedWeaponId = payload.id;
  state.editingWeaponId = payload.id;
  state.weaponDraft = deepClone(payload);
  persistWeapons();
  saveStoredValue(STORAGE_KEYS.selectedWeaponId, state.selectedWeaponId);
  renderAll();
  scrollEditorIntoView(els.weaponForm);
}

function editBuild(id) {
  const build = getBuildById(id);
  if (!build) {
    return;
  }
  state.editingBuildId = id;
  state.selectedBuildId = id;
  state.buildDraft = deepClone(build);
  renderAll();
  scrollEditorIntoView(els.buildForm);
}

function editWeapon(id) {
  const weapon = getWeaponById(id);
  if (!weapon) {
    return;
  }
  state.editingWeaponId = id;
  state.selectedWeaponId = id;
  state.weaponDraft = deepClone(weapon);
  renderAll();
  scrollEditorIntoView(els.weaponForm);
}

function deleteBuild(id) {
  state.builds = state.builds.filter((item) => item.id !== id);
  if (!state.builds.length) {
    const seeded = buildDefaultBuild();
    state.builds = [seeded];
  }
  if (state.selectedBuildId === id || !getBuildById(state.selectedBuildId)) {
    state.selectedBuildId = state.builds[0].id;
    saveStoredValue(STORAGE_KEYS.selectedBuildId, state.selectedBuildId);
  }
  if (state.editingBuildId === id) {
    state.editingBuildId = null;
    state.buildDraft = null;
  }
  persistBuilds();
  renderAll();
}

function deleteWeapon(id) {
  state.weapons = state.weapons.filter((item) => item.id !== id);
  if (!state.weapons.length) {
    const seeded = buildDefaultWeapon();
    state.weapons = [seeded];
  }
  if (state.selectedWeaponId === id || !getWeaponById(state.selectedWeaponId)) {
    state.selectedWeaponId = state.weapons[0].id;
    saveStoredValue(STORAGE_KEYS.selectedWeaponId, state.selectedWeaponId);
  }
  if (state.editingWeaponId === id) {
    state.editingWeaponId = null;
    state.weaponDraft = null;
  }
  persistWeapons();
  renderAll();
}

function renderAll() {
  renderSelectionActions();
  renderCalculatorSelectors();
  renderResultGrid();
  renderCalculatorActions();
  renderSkillSummary();
  renderBuildList();
  renderWeaponList();
  renderBuildForm();
  renderWeaponForm();
}

function generateRiftVariants(weapon) {
  const type = weapon.values.E5;
  const options = type === "Raw" ? ["attack", "affinity"] : ["attack", "element", "affinity"];
  const variants = [];

  function helper(index, remaining, counts) {
    if (index === options.length - 1) {
      counts[options[index]] = remaining;
      variants.push({ ...counts });
      return;
    }
    for (let count = 0; count <= remaining; count += 1) {
      counts[options[index]] = count;
      helper(index + 1, remaining - count, counts);
    }
  }

  helper(0, 3, {});

  return variants.map((counts) => {
    const nextWeapon = deepClone(weapon);
    nextWeapon.values.E3 = Number(nextWeapon.values.E3) + (counts.attack ?? 0) * 100;
    if (type !== "Raw") {
      nextWeapon.values.E4 = Number(nextWeapon.values.E4) + (counts.element ?? 0) * 100;
    }
    nextWeapon.values.E6 = Number(nextWeapon.values.E6) + (counts.affinity ?? 0) * 0.1;
    return {
      counts,
      weapon: nextWeapon,
    };
  });
}

function describeRiftVariant(counts, isRaw) {
  const parts = [];
  if (counts.attack) {
    parts.push(`+${counts.attack * 100} Attack`);
  }
  if (!isRaw && counts.element) {
    parts.push(`+${counts.element * 100} Element`);
  }
  if (counts.affinity) {
    parts.push(`+${counts.affinity * 10} Affinity`);
  }
  return parts.join(", ");
}

function describeRiftUpgradeSequence(counts, isRaw) {
  const upgrades = [];
  for (let index = 0; index < (counts.attack ?? 0); index += 1) {
    upgrades.push("attack");
  }
  if (!isRaw) {
    for (let index = 0; index < (counts.element ?? 0); index += 1) {
      upgrades.push("element");
    }
  }
  for (let index = 0; index < (counts.affinity ?? 0); index += 1) {
    upgrades.push("affinity");
  }
  return upgrades.join(", ");
}

function startWeaponDraftFromRiftVariant(baseWeapon, variant) {
  const label = describeRiftVariant(variant.counts, baseWeapon.values.E5 === "Raw");
  state.editingWeaponId = null;
  state.weaponDraft = {
    ...deepClone(variant.weapon),
    id: makeId(),
    name: `${baseWeapon.name} (${label || "Rift"})`,
    isRift: false,
  };
  closeModal();
  renderWeaponForm();
  renderWeaponList();
  scrollEditorIntoView(els.weaponForm);
}

function openRiftComparison() {
  const build = getBuildById(state.selectedBuildId);
  const weapon = getWeaponById(state.selectedWeaponId);
  if (!build || !weapon || !weapon.isRift) {
    return;
  }

  const engine = createEngine();
  const variants = generateRiftVariants(weapon)
    .map((variant) => {
      applyScenario(engine, build, variant.weapon);
      const isRaw = weapon.values.E5 === "Raw";
      return {
        ...variant,
        label: describeRiftVariant(variant.counts, isRaw),
        upgradeSequence: describeRiftUpgradeSequence(variant.counts, isRaw),
        h12: readCell(engine, "Calculator1", state.data.resultCell),
      };
    })
    .sort((a, b) => {
      const av = typeof a.h12 === "number" ? a.h12 : -Infinity;
      const bv = typeof b.h12 === "number" ? b.h12 : -Infinity;
      return bv - av;
    });
  const topValue = variants.find((variant) => typeof variant.h12 === "number")?.h12;

  els.modalTitle.textContent = "Rift Combinations";
  els.riftModal.querySelector(".modal-window")?.classList.add("modal-window-compact");
  els.riftModalContent.innerHTML = `
    <div class="rift-context">
      <div><span class="rift-context-label">Build:</span> ${escapeHtml(build.name)} <span class="rift-context-separator">|</span> <span class="rift-context-label">Weapon:</span> ${escapeHtml(weapon.name)}</div>
    </div>
    <div class="rift-results">
      ${variants
        .map(
          (variant, index) => {
            const relativePercent =
              typeof variant.h12 === "number" && typeof topValue === "number" && topValue !== 0
                ? ((variant.h12 - topValue) / topValue) * 100
                : Number.NaN;
            const relativeLabel =
              index === 0 && typeof variant.h12 === "number" ? "100.00%" : formatSignedPercent(relativePercent);

            return `
            <div class="rift-result-item">
              <div>
                <div>${escapeHtml(variant.label || "No bonus")}</div>
                <div class="rift-result-label">${escapeHtml(variant.upgradeSequence || "attack, attack, attack")}</div>
              </div>
              <div class="rift-result-actions">
                <div class="rift-result-value-block">
                  <div class="rift-result-value">${escapeHtml(formatResult(variant.h12))}</div>
                  <div class="rift-result-relative">${escapeHtml(relativeLabel)}</div>
                </div>
                <button class="secondary" type="button" data-rift-variant-index="${index}">Use In Weapon Editor</button>
              </div>
            </div>
          `;
          },
        )
        .join("")}
    </div>
  `;
  els.riftModalContent.querySelectorAll("[data-rift-variant-index]").forEach((button) => {
    button.addEventListener("click", () => {
      const variant = variants[Number(button.dataset.riftVariantIndex)];
      if (!variant) {
        return;
      }
      startWeaponDraftFromRiftVariant(weapon, variant);
    });
  });
  els.riftModal.classList.remove("hidden");
}

function buildMatrixKey(buildId, weaponId) {
  return `${weaponId}::${buildId}`;
}

function openBuildWeaponComparison() {
  const buildsToCompare = state.builds.filter((build) => build.compareEnabled !== false);
  const weaponsToCompare = state.weapons.filter((weapon) => weapon.compareEnabled !== false);
  if (!buildsToCompare.length || !weaponsToCompare.length) {
    els.modalTitle.textContent = "Build and Weapon Comparison";
    els.riftModal.querySelector(".modal-window")?.classList.add("modal-window-medium");
    els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-compact");
    els.riftModalContent.innerHTML = `
      <div class="comparison-intro">Select at least one build and one weapon with the checkboxes before running the comparison.</div>
    `;
    els.riftModal.classList.remove("hidden");
    return;
  }

  const engine = createEngine();
  const results = {};
  for (const weapon of weaponsToCompare) {
    for (const build of buildsToCompare) {
      applyScenario(engine, build, weapon);
      results[buildMatrixKey(build.id, weapon.id)] = readCell(engine, "Calculator1", state.data.resultCell);
    }
  }

  state.matrixComparison = {
    results,
    buildIds: buildsToCompare.map((build) => build.id),
    weaponIds: weaponsToCompare.map((weapon) => weapon.id),
    referenceBuildId:
      buildsToCompare.find((build) => build.id === state.selectedBuildId)?.id ?? buildsToCompare[0].id,
    referenceWeaponId:
      weaponsToCompare.find((weapon) => weapon.id === state.selectedWeaponId)?.id ?? weaponsToCompare[0].id,
  };
  renderBuildWeaponComparison();
}

function openExportModal(payload = exportAppData(), description = "Copy this string and keep it somewhere safe. It contains all currently stored builds and weapons.") {
  els.modalTitle.textContent = "Export Data";
  els.riftModal.querySelector(".modal-window")?.classList.add("modal-window-medium");
  els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-compact");
  els.riftModalContent.innerHTML = `
    <div class="transfer-copy">
      <p class="comparison-intro">${escapeHtml(description)}</p>
      <textarea class="transfer-textarea" id="export-payload" readonly></textarea>
      <div class="transfer-actions">
        <button type="button" id="copy-export-payload">Copy</button>
      </div>
      <div class="transfer-feedback" id="transfer-feedback"></div>
    </div>
  `;
  els.riftModalContent.querySelector("#export-payload").value = payload;
  els.riftModalContent.querySelector("#copy-export-payload").addEventListener("click", async () => {
    const feedback = els.riftModalContent.querySelector("#transfer-feedback");
    try {
      await copyTextToClipboard(payload);
      feedback.textContent = "Copied to clipboard.";
    } catch {
      feedback.textContent = "Copy failed. Select the text and copy it manually.";
    }
  });
  els.riftModal.classList.remove("hidden");
}

function openImportModal() {
  els.modalTitle.textContent = "Import Data";
  els.riftModal.querySelector(".modal-window")?.classList.add("modal-window-medium");
  els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-compact");
  els.riftModalContent.innerHTML = `
    <div class="transfer-copy">
      <p class="comparison-intro">Paste an exported string below. Imported builds and weapons will be added to the current data.</p>
      <textarea class="transfer-textarea" id="import-payload" placeholder="Paste export string here"></textarea>
      <div class="transfer-actions">
        <button type="button" id="submit-import-payload">Import</button>
      </div>
      <div class="transfer-feedback" id="transfer-feedback"></div>
    </div>
  `;
  els.riftModalContent.querySelector("#submit-import-payload").addEventListener("click", () => {
    const payload = els.riftModalContent.querySelector("#import-payload").value.trim();
    const feedback = els.riftModalContent.querySelector("#transfer-feedback");
    if (!payload) {
      feedback.textContent = "Paste an export string first.";
      return;
    }

    try {
      const parsed = JSON.parse(payload);
      const existingBuildFingerprints = new Set(state.builds.map(buildImportFingerprint));
      const existingWeaponFingerprints = new Set(state.weapons.map(weaponImportFingerprint));
      const importedBuilds = [];
      const importedWeapons = [];

      if (Array.isArray(parsed.builds)) {
        for (const rawBuild of parsed.builds) {
          const normalizedBuild = normalizeImportedBuild(rawBuild);
          const fingerprint = buildImportFingerprint(normalizedBuild);
          if (existingBuildFingerprints.has(fingerprint)) {
            continue;
          }
          existingBuildFingerprints.add(fingerprint);
          importedBuilds.push(normalizedBuild);
        }
      }

      if (Array.isArray(parsed.weapons)) {
        for (const rawWeapon of parsed.weapons) {
          const normalizedWeapon = normalizeImportedWeapon(rawWeapon);
          const fingerprint = weaponImportFingerprint(normalizedWeapon);
          if (existingWeaponFingerprints.has(fingerprint)) {
            continue;
          }
          existingWeaponFingerprints.add(fingerprint);
          importedWeapons.push(normalizedWeapon);
        }
      }

      if (!importedBuilds.length && !importedWeapons.length) {
        feedback.textContent = "No new builds or weapons were imported.";
        return;
      }

      state.builds = [...state.builds, ...importedBuilds];
      state.weapons = [...state.weapons, ...importedWeapons];
      persistBuilds();
      persistWeapons();
      renderAll();
      feedback.textContent = `Imported ${importedBuilds.length} builds and ${importedWeapons.length} weapons.`;
    } catch {
      feedback.textContent = "Import failed. Check that the pasted string is a valid export.";
    }
  });
  els.riftModal.classList.remove("hidden");
}

function renderBuildWeaponComparison() {
  if (!state.matrixComparison) {
    return;
  }

  const { referenceBuildId, referenceWeaponId, results, buildIds, weaponIds } = state.matrixComparison;
  const buildsToCompare = buildIds.map((id) => getBuildById(id)).filter(Boolean);
  const weaponsToCompare = weaponIds.map((id) => getWeaponById(id)).filter(Boolean);
  const referenceValue = results[buildMatrixKey(referenceBuildId, referenceWeaponId)];

  const headerCells = buildsToCompare
    .map(
      (build) => `
        <th scope="col" class="comparison-header-cell" title="${escapeHtml(build.name)}">
          <span class="comparison-header-cell-text">${escapeHtml(build.name)}</span>
        </th>
      `,
    )
    .join("");

  const bodyRows = weaponsToCompare
    .map((weapon) => {
      const cells = buildsToCompare
        .map((build) => {
          const key = buildMatrixKey(build.id, weapon.id);
          const value = results[key];
          const isReference = build.id === referenceBuildId && weapon.id === referenceWeaponId;
          const relativePercent =
            typeof value === "number" && typeof referenceValue === "number" && referenceValue !== 0
              ? ((value - referenceValue) / referenceValue) * 100
              : Number.NaN;
          const deltaLabel = isReference ? "100.00%" : formatSignedPercent(relativePercent);
          const background = isReference
            ? "rgba(79, 140, 255, 0.22)"
            : getDeltaBackground(relativePercent);

          return `
            <td>
              <button
                type="button"
                class="comparison-cell ${isReference ? "is-reference" : ""}"
                data-matrix-build-id="${build.id}"
                data-matrix-weapon-id="${weapon.id}"
                style="background:${background};"
              >
                <span class="comparison-dps">${escapeHtml(formatResult(value))}</span>
                <span class="comparison-delta">${escapeHtml(deltaLabel)}</span>
              </button>
            </td>
          `;
        })
        .join("");

      return `
        <tr>
          <th scope="row" class="comparison-header-row" title="${escapeHtml(weapon.name)}">
            <span class="comparison-header-row-text">${escapeHtml(weapon.name)}</span>
          </th>
          ${cells}
        </tr>
      `;
    })
    .join("");

  els.modalTitle.textContent = "Build and Weapon Comparison";
  els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-compact");
  els.riftModalContent.innerHTML = `
    <div class="comparison-intro">
      Click any cell to use it as the reference. The selected reference shows 100.00%, and all other cells show relative DPS difference against it.
    </div>
    <div class="comparison-table-wrap">
      <table class="comparison-table">
        <thead>
          <tr>
            <th class="comparison-corner"></th>
            ${headerCells}
          </tr>
        </thead>
        <tbody>
          ${bodyRows}
        </tbody>
      </table>
    </div>
  `;

  els.riftModalContent.querySelectorAll("[data-matrix-build-id]").forEach((button) => {
    button.addEventListener("click", () => {
      state.matrixComparison.referenceBuildId = button.dataset.matrixBuildId;
      state.matrixComparison.referenceWeaponId = button.dataset.matrixWeaponId;
      renderBuildWeaponComparison();
    });
  });

  els.riftModal.classList.remove("hidden");
}

function closeModal() {
  state.matrixComparison = null;
  state.uptimeDraft = null;
  els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-compact");
  els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-medium");
  els.riftModal.querySelector(".modal-window")?.classList.remove("modal-window-uptime");
  els.riftModal.classList.add("hidden");
  els.riftModalContent.innerHTML = "";
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function wireGlobalEvents() {
  els.calculatorBuild.addEventListener("change", (event) => {
    state.selectedBuildId = event.target.value;
    saveStoredValue(STORAGE_KEYS.selectedBuildId, state.selectedBuildId);
    renderResultGrid();
    renderCalculatorActions();
    renderSkillSummary();
    renderBuildList();
  });

  els.calculatorWeapon.addEventListener("change", (event) => {
    state.selectedWeaponId = event.target.value;
    saveStoredValue(STORAGE_KEYS.selectedWeaponId, state.selectedWeaponId);
    renderResultGrid();
    renderCalculatorActions();
    renderWeaponList();
  });

  els.newBuild.addEventListener("click", () => {
    state.editingBuildId = null;
    state.buildDraft = buildDefaultBuild();
    renderBuildForm();
    renderBuildList();
    scrollEditorIntoView(els.buildForm);
  });

  els.newWeapon.addEventListener("click", () => {
    state.editingWeaponId = null;
    state.weaponDraft = buildDefaultWeapon();
    renderWeaponForm();
    renderWeaponList();
    scrollEditorIntoView(els.weaponForm);
  });

  els.exportSelectedBuilds.addEventListener("click", () => {
    const builds = state.builds.filter((build) => build.compareEnabled !== false);
    openExportModal(
      exportScopedData({ builds, weapons: [] }),
      "Copy this string to export only the selected builds.",
    );
  });

  els.exportSelectedWeapons.addEventListener("click", () => {
    const weapons = state.weapons.filter((weapon) => weapon.compareEnabled !== false);
    openExportModal(
      exportScopedData({ builds: [], weapons }),
      "Copy this string to export only the selected weapons.",
    );
  });

  els.exportData.addEventListener("click", () => {
    openExportModal();
  });

  els.importData.addEventListener("click", () => {
    openImportModal();
  });

  document.body.addEventListener("click", (event) => {
    const button = event.target.closest("button[data-action]");
    if (button) {
      const { action, id } = button.dataset;
      if (action === "edit-build") {
        editBuild(id);
      } else if (action === "delete-build") {
        deleteBuild(id);
      } else if (action === "edit-weapon") {
        editWeapon(id);
      } else if (action === "delete-weapon") {
        deleteWeapon(id);
      }
      return;
    }

    const checkbox = event.target.closest('input[data-action="toggle-compare"]');
    if (!checkbox) {
      return;
    }

    const containerId = checkbox.closest("[id]")?.id;
    if (containerId === "build-list") {
      state.builds = state.builds.map((build) =>
        build.id === checkbox.dataset.id ? { ...build, compareEnabled: checkbox.checked } : build,
      );
      persistBuilds();
      renderBuildList();
      return;
    }

    if (containerId === "weapon-list") {
      state.weapons = state.weapons.map((weapon) =>
        weapon.id === checkbox.dataset.id ? { ...weapon, compareEnabled: checkbox.checked } : weapon,
      );
      persistWeapons();
      renderWeaponList();
    }
  });

  window.addEventListener("resize", () => {
    if (state.buildDraft) {
      const nextColumnCount = getBuildEditorColumnCount();
      if (nextColumnCount !== state.buildEditorColumnCount) {
        renderBuildForm();
      } else {
        syncBuildEditorSeparators();
      }
      return;
    }
    syncBuildEditorSeparators();
  });
}

async function init() {
  if (!window.HyperFormula) {
    throw new Error("HyperFormula failed to load.");
  }

  if (window.WORKBOOK_DATA) {
    state.data = window.WORKBOOK_DATA;
  } else {
    const response = await fetch("./data/workbook-data.json");
    state.data = await response.json();
  }

  if (!state.data) {
    throw new Error("Workbook data failed to load.");
  }

  state.builds = loadStoredItems(STORAGE_KEYS.builds);
  state.weapons = loadStoredItems(STORAGE_KEYS.weapons);
  state.uptimeFields = collectDefaultUptimeFields();
  state.uptimeValues = {
    ...buildDefaultUptimeValues(),
    ...loadStoredObject(STORAGE_KEYS.uptimes, {}),
  };

  if (!state.builds.length) {
    state.builds = [buildDefaultBuild()];
    persistBuilds();
  }
  if (!state.weapons.length) {
    state.weapons = [buildDefaultWeapon()];
    persistWeapons();
  }

  state.builds = state.builds.map((build) => ({
    ...buildDefaultBuild(),
    ...build,
    values: { ...buildDefaultBuild().values, ...build.values },
    compareEnabled: typeof build.compareEnabled === "boolean" ? build.compareEnabled : true,
  }));
  persistBuilds();

  state.weapons = state.weapons.map((weapon) => ({
    ...buildDefaultWeapon(),
    ...weapon,
    values: { ...buildDefaultWeapon().values, ...weapon.values },
    isRift: Boolean(weapon.isRift),
    compareEnabled: typeof weapon.compareEnabled === "boolean" ? weapon.compareEnabled : true,
  }));
  persistWeapons();

  const storedBuildId = loadStoredValue(STORAGE_KEYS.selectedBuildId);
  const storedWeaponId = loadStoredValue(STORAGE_KEYS.selectedWeaponId);
  state.selectedBuildId = getBuildById(storedBuildId)?.id ?? state.builds[0].id;
  state.selectedWeaponId = getWeaponById(storedWeaponId)?.id ?? state.weapons[0].id;
  saveStoredValue(STORAGE_KEYS.selectedBuildId, state.selectedBuildId);
  saveStoredValue(STORAGE_KEYS.selectedWeaponId, state.selectedWeaponId);
  state.editingBuildId = state.builds[0].id;
  state.editingWeaponId = state.weapons[0].id;
  state.buildDraft = null;
  state.weaponDraft = null;

  wireGlobalEvents();
  els.closeRiftModal.addEventListener("click", closeModal);
  els.riftModal.addEventListener("click", (event) => {
    if (event.target === els.riftModal) {
      closeModal();
    }
  });
  renderAll();
}

init().catch((error) => {
  console.error(error);
  document.body.innerHTML = `<pre style="padding:24px;color:#fff;background:#101216;">${escapeHtml(
    error.stack || String(error),
  )}</pre>`;
});
