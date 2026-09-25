const STORAGE_KEYS = {
  builds: "mhn_builds_v2",
  weapons: "mhn_weapons_v2",
  selectedBuildId: "mhn_selected_build_v2",
  selectedWeaponId: "mhn_selected_weapon_v2",
  uptimes: "mhn_uptimes_v2",
};

const PREVIOUS_STORAGE_KEYS = {
  builds: "mhn_builds_v363",
  weapons: "mhn_weapons_v363",
  selectedBuildId: "mhn_selected_build_v363",
  selectedWeaponId: "mhn_selected_weapon_v363",
  uptimes: "mhn_uptimes_v363",
};

const LEGACY_STORAGE_KEYS = {
  builds: "mhn_builds_v1",
  weapons: "mhn_weapons_v1",
  selectedBuildId: "mhn_selected_build_v1",
  selectedWeaponId: "mhn_selected_weapon_v1",
  uptimes: "mhn_uptimes_v1",
};

const EXPORT_FORMAT_VERSION = 2;

const LEGACY_BUILD_KEYS = [
  "elementalAttack",
  "advancedElementalAttack",
  "aggressiveDodger",
  "artillery",
  "attackBoost",
  "advancedAttackBoost",
  "attackEfficacy",
  "bleedingEdge",
  "bluntForce",
  "bubblyDance",
  "buildupBoost",
  "burst",
  "advancedBurst",
  "chameleosVenomist",
  "chargeMaster",
  "coalescence",
  "criticalBoost",
  "criticalElement",
  "criticalEye",
  "criticalFerocity",
  "criticalStrength",
  "dauntless",
  "fightingSpirit",
  "fortify",
  "headstrong",
  "hellfireCloak",
  "heroics",
  "huntersUnity",
  "kirinFlashstorm",
  "kushalaFrostwind",
  "latentPower",
  "malzenoCrimsonblood",
  "namielleElectrowave",
  "nergiganteAvidity",
  "morphBoost",
  "offensiveDodger",
  "offensiveGuard",
  "paralysisExploit",
  "partbreaker",
  "peakPerformance",
  "poisonExploit",
  "pursuit",
  "rawPower",
  "reckless",
  "resentment",
  "resuscitate",
  "risingTide",
  "sharedSword",
  "skywardStriker",
  "sneakAttack",
  "solidarity",
  "specialBoost",
  "specialPartbreaker",
  "statusSneakAttack",
  "sweltingSummer",
  "teostraBlastpowder",
  "timedCharger",
  "valor",
  "vitalElement",
  "weaknessExploit",
];

const LEGACY_BUILD_REF_TO_KEY = Object.freeze(
  Object.fromEntries(LEGACY_BUILD_KEYS.map((key, index) => [`B${index + 3}`, key])),
);

const LEGACY_WEAPON_REF_TO_KEY = Object.freeze({
  E3: "weaponAttack",
  E4: "weaponElement",
  E5: "damageType",
  E6: "weaponAffinity",
  E7: "weaponType",
});

const LEGACY_UPTIME_REF_TO_KEY = Object.freeze({
  E15: "remainingHealth",
  E24: "aggressiveDodger",
  E25: "artillery",
  E26: "bluntForce",
  E27: "burst",
  E28: "chargeMaster",
  E29: "dauntless",
  E30: "fightingSpirit",
  E31: "fortify",
  E32: "headstrong",
  E33: "heroics",
  E34: "latentPower",
  E35: "morphBoost",
  E36: "offensiveDodger",
  E37: "offensiveGuard",
  E38: "paralysisExploit",
  E39: "poisonExploit",
  E40: "pursuit",
  E41: "resentment",
  E42: "skywardStriker",
  E43: "statusSneakAttack",
  E44: "specialBoost",
  E45: "specialPartbreaker",
  E46: "valor",
  E47: "weaknessExploit",
});

const MODAL_MODE_CLASSES = [
  "modal-window-compact",
  "modal-window-medium",
  "modal-window-uptime",
];

const WEAPON_LIBRARY_TYPE_MAP = Object.freeze({
  SwordAndShield: "Sword & Shield",
  DualBlades: "Dual Blades",
  GreatSword: "Greatsword",
  LongSword: "Longsword",
  Hammer: "Hammer",
  HuntingHorn: "Hunting Horn",
  Lance: "Lance",
  Gunlance: "Gunlance",
  SwitchAxe: "Switch Axe",
  LightBowgun: "Light Bowgun",
  HeavyBowgun: "Heavy Bowgun",
  Bow: "Bow",
});

const WEAPON_LIBRARY_ATTRIBUTE_LABELS = Object.freeze({
  none: "Raw",
  fire: "Fire",
  water: "Water",
  thunder: "Thunder",
  ice: "Ice",
  dragon: "Dragon",
  poison: "Poison",
  paralysis: "Paralysis",
  sleep: "Sleep",
  blast: "Blast",
});

const LIBRARY_ACTION_HANDLERS = {
  "edit-build": (id) => editBuild(id),
  "delete-build": (id) => deleteBuild(id),
  "move-build-up": (id) => moveBuild(id, -1),
  "move-build-down": (id) => moveBuild(id, 1),
  "edit-weapon": (id) => editWeapon(id),
  "delete-weapon": (id) => deleteWeapon(id),
  "move-weapon-up": (id) => moveWeapon(id, -1),
  "move-weapon-down": (id) => moveWeapon(id, 1),
};

const COMPARE_COLLECTION_BY_LIST_ID = {
  "build-list": "builds",
  "weapon-list": "weapons",
};

const state = {
  data: null,
  extensionStatus: null,
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
  weaponLibraryType: "all",
  weaponLibraryAttribute: "all",
  weaponLibraryQuery: "",
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
  selectAllBuilds: document.getElementById("select-all-builds"),
  deselectAllBuilds: document.getElementById("deselect-all-builds"),
  newWeapon: document.getElementById("new-weapon"),
  selectAllWeapons: document.getElementById("select-all-weapons"),
  deselectAllWeapons: document.getElementById("deselect-all-weapons"),
  selectWeaponType: document.getElementById("select-weapon-type"),
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

function sortBuildFieldsAlphabetically(fields, labels) {
  return [...fields].sort((left, right) => {
    const leftLabel = String(labels[left.ref] ?? left.labelRef);
    const rightLabel = String(labels[right.ref] ?? right.labelRef);
    return leftLabel.localeCompare(rightLabel, undefined, { sensitivity: "base" });
  });
}

function orderBuildFieldsByVisibleColumn(fields) {
  const columnCount = getBuildEditorColumnCount();
  if (columnCount <= 1) {
    return fields;
  }

  const rowCount = Math.ceil(fields.length / columnCount);
  const shortColumnSize = Math.floor(fields.length / columnCount);
  const longColumnCount = fields.length % columnCount;
  const columnSizes = Array.from(
    { length: columnCount },
    (_, index) => shortColumnSize + (index < longColumnCount ? 1 : 0),
  );
  const columnStarts = [];
  columnSizes.reduce((start, size) => {
    columnStarts.push(start);
    return start + size;
  }, 0);
  const ordered = [];

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const sourceIndex = columnStarts[columnIndex] + rowIndex;
      ordered.push(rowIndex < columnSizes[columnIndex] ? fields[sourceIndex] : null);
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
  const sheets = deepClone(state.data.sheets);
  if (state.extensionStatus?.enabled) {
    window.PHASK_SKILL_EXTENSIONS.prepareSheets(sheets, state.data.sheetVersion);
  }
  return HyperFormula.buildFromSheets(sheets, {
    licenseKey: "gpl-v3",
  });
}

function calculatorSheetName() {
  return state.data?.calculatorSheet ?? "Calculator1";
}

function readCell(engine, sheetName, ref) {
  const sheetId = engine.getSheetId(sheetName);
  return engine.getCellValue(toAddress(sheetId, ref));
}

function writeCell(engine, sheetName, ref, value) {
  const sheetId = engine.getSheetId(sheetName);
  engine.setCellContents(toAddress(sheetId, ref), value);
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

function loadFirstStoredValue(keys) {
  for (const key of keys) {
    const value = loadStoredValue(key);
    if (value) {
      return value;
    }
  }
  return null;
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

function fieldRefToKeyMap(fields) {
  return Object.fromEntries(fields.map((field) => [field.ref, field.key ?? field.ref]));
}

function fieldKeyToRefMap(fields) {
  return Object.fromEntries(fields.map((field) => [field.key ?? field.ref, field.ref]));
}

function previousVersionRefToKeyMap(fields, insertedKey) {
  const insertedField = fields.find((field) => field.key === insertedKey);
  const insertedMatch = insertedField?.ref?.match(/^([A-Z]+)(\d+)$/);
  if (!insertedMatch) {
    return fieldRefToKeyMap(fields);
  }

  const [, insertedColumn, insertedRowText] = insertedMatch;
  const insertedRow = Number(insertedRowText);
  return Object.fromEntries(
    fields
      .filter((field) => field.key !== insertedKey)
      .map((field) => {
        const match = field.ref.match(/^([A-Z]+)(\d+)$/);
        if (!match || match[1] !== insertedColumn || Number(match[2]) <= insertedRow) {
          return [field.ref, field.key ?? field.ref];
        }
        return [`${match[1]}${Number(match[2]) - 1}`, field.key ?? field.ref];
      }),
  );
}

function addImportWarning(warnings, context, field) {
  warnings.push(`${context}: ${field}`);
}

function refToKeyMapForSource(fields, source, previousInsertedKey, legacyMap) {
  if (source === "legacy") {
    return legacyMap;
  }
  if (source === "previous") {
    return previousVersionRefToKeyMap(fields, previousInsertedKey);
  }
  return fieldRefToKeyMap(fields);
}

function semanticValuesFromCurrent(values, fields) {
  const refToKey = fieldRefToKeyMap(fields);
  return Object.fromEntries(
    fields.map((field) => [field.key ?? field.ref, values?.[field.ref] ?? field.defaultValue ?? 0]),
  );
}

function semanticValuesFromInput(values, fields, refToKey, warnings, context) {
  const semanticValues = {};
  const currentKeys = new Set(fields.map((field) => field.key ?? field.ref));
  if (!values || typeof values !== "object") {
    return semanticValues;
  }

  Object.entries(values).forEach(([field, value]) => {
    if (refToKey[field]) {
      semanticValues[refToKey[field]] = value;
    } else if (currentKeys.has(field)) {
      semanticValues[field] = value;
    } else {
      addImportWarning(warnings, context, field);
    }
  });

  return semanticValues;
}

function currentValuesFromSemantic(semanticValues, fields, warnings, context) {
  const keyToRef = fieldKeyToRefMap(fields);
  const values = {};
  Object.entries(semanticValues).forEach(([key, value]) => {
    const ref = keyToRef[key];
    if (!ref) {
      addImportWarning(warnings, context, key);
      return;
    }
    values[ref] = value;
  });
  return values;
}

function serializeBuild(build) {
  return {
    name: build.name,
    compareEnabled: build.compareEnabled !== false,
    skills: semanticValuesFromCurrent(build.values, state.data.buildFields),
  };
}

function serializeStoredBuild(build) {
  return {
    id: build.id,
    ...serializeBuild(build),
  };
}

function serializeWeapon(weapon) {
  return {
    name: weapon.name,
    compareEnabled: weapon.compareEnabled !== false,
    isRift: Boolean(weapon.isRift),
    ...(weapon.libraryId ? { libraryId: weapon.libraryId } : {}),
    ...(weapon.libraryVariant ? { libraryVariant: weapon.libraryVariant } : {}),
    ...(Number.isFinite(weapon.riftLevel) ? { riftLevel: weapon.riftLevel } : {}),
    ...(typeof weapon.riftNodesApplied === "boolean"
      ? { riftNodesApplied: weapon.riftNodesApplied }
      : {}),
    properties: semanticValuesFromCurrent(weapon.values, state.data.weaponFields),
  };
}

function serializeStoredWeapon(weapon) {
  return {
    id: weapon.id,
    ...serializeWeapon(weapon),
  };
}

function exportAppData() {
  return JSON.stringify(
    {
      formatVersion: EXPORT_FORMAT_VERSION,
      sheetVersion: state.data?.sheetVersion ?? "unknown",
      builds: state.builds.map(serializeBuild),
      weapons: state.weapons.map(serializeWeapon),
    },
    null,
    0,
  );
}

function exportScopedData({ builds = [], weapons = [] }) {
  return JSON.stringify(
    {
      formatVersion: EXPORT_FORMAT_VERSION,
      sheetVersion: state.data?.sheetVersion ?? "unknown",
      builds: builds.map(serializeBuild),
      weapons: weapons.map(serializeWeapon),
    },
    null,
    0,
  );
}

function normalizeBuildItem(item, { source = "current", preserveId = false, warnings = [] } = {}) {
  const base = buildDefaultBuild();
  const name = String(item?.name || "Imported Build");
  const context = `Build "${name}"`;
  const sourceRefToKey = refToKeyMapForSource(
    state.data.buildFields,
    source,
    "elementalRelease",
    LEGACY_BUILD_REF_TO_KEY,
  );
  const semanticValues =
    item?.skills && typeof item.skills === "object"
      ? semanticValuesFromInput(item.skills, state.data.buildFields, {}, warnings, context)
      : semanticValuesFromInput(item?.values, state.data.buildFields, sourceRefToKey, warnings, context);

  return {
    ...base,
    id: preserveId && item?.id ? String(item.id) : makeId(),
    name,
    compareEnabled: typeof item?.compareEnabled === "boolean" ? item.compareEnabled : true,
    values: {
      ...base.values,
      ...currentValuesFromSemantic(semanticValues, state.data.buildFields, warnings, context),
    },
  };
}

function normalizeBuildUptimes(values, { source = "current", warnings = [] } = {}) {
  const context = "Uptimes";
  const sourceRefToKey = refToKeyMapForSource(
    state.uptimeFields,
    source,
    "elementalRelease",
    LEGACY_UPTIME_REF_TO_KEY,
  );
  const semanticValues = semanticValuesFromInput(
    values,
    state.uptimeFields,
    sourceRefToKey,
    warnings,
    context,
  );
  return currentValuesFromSemantic(semanticValues, state.uptimeFields, warnings, context);
}

function normalizeImportedBuild(item, warnings = [], source = "legacy") {
  return normalizeBuildItem(item, { source, warnings });
}

function normalizeWeaponItem(item, { source = "current", preserveId = false, warnings = [] } = {}) {
  const base = buildDefaultWeapon();
  const name = String(item?.name || "Imported Weapon");
  const context = `Weapon "${name}"`;
  const sourceRefToKey = refToKeyMapForSource(
    state.data.weaponFields,
    source,
    null,
    LEGACY_WEAPON_REF_TO_KEY,
  );
  const semanticValues =
    item?.properties && typeof item.properties === "object"
      ? semanticValuesFromInput(item.properties, state.data.weaponFields, {}, warnings, context)
      : semanticValuesFromInput(item?.values, state.data.weaponFields, sourceRefToKey, warnings, context);

  return {
    ...base,
    id: preserveId && item?.id ? String(item.id) : makeId(),
    name,
    compareEnabled: typeof item?.compareEnabled === "boolean" ? item.compareEnabled : true,
    values: {
      ...base.values,
      ...currentValuesFromSemantic(semanticValues, state.data.weaponFields, warnings, context),
    },
    isRift: Boolean(item?.isRift),
    libraryId: item?.libraryId ? String(item.libraryId) : null,
    libraryVariant: item?.libraryVariant ? String(item.libraryVariant) : null,
    riftLevel:
      item?.riftLevel != null && Number.isFinite(Number(item.riftLevel))
        ? Number(item.riftLevel)
        : null,
    riftNodesApplied:
      typeof item?.riftNodesApplied === "boolean" ? item.riftNodesApplied : null,
  };
}

function normalizeImportedWeapon(item, warnings = [], source = "legacy") {
  return normalizeWeaponItem(item, { source, warnings });
}

function loadMigratedItems(sources, normalizeItem, warningLabel) {
  for (const source of sources) {
    const items = loadStoredItems(source.key);
    if (!items.length) {
      continue;
    }

    const warnings = [];
    const migrated = items.map((item) =>
      normalizeItem(item, {
        source: source.source,
        preserveId: true,
        warnings,
      }),
    );
    if (warnings.length) {
      console.warn(`Some ${warningLabel} fields could not be migrated:`, warnings);
    }
    return migrated;
  }
  return [];
}

function loadMigratedObject(sources, normalizeObject, warningLabel, fallback = {}) {
  for (const source of sources) {
    const value = loadStoredObject(source.key, null);
    if (!value) {
      continue;
    }

    const warnings = [];
    const migrated = normalizeObject(value, { source: source.source, warnings });
    if (warnings.length) {
      console.warn(`Some ${warningLabel} fields could not be migrated:`, warnings);
    }
    return migrated;
  }
  return fallback;
}

function loadBuildsForCurrentVersion() {
  return loadMigratedItems(
    [
      { key: STORAGE_KEYS.builds, source: "current" },
      { key: PREVIOUS_STORAGE_KEYS.builds, source: "previous" },
      { key: LEGACY_STORAGE_KEYS.builds, source: "legacy" },
    ],
    normalizeBuildItem,
    "build",
  );
}

function loadWeaponsForCurrentVersion() {
  return loadMigratedItems(
    [
      { key: STORAGE_KEYS.weapons, source: "current" },
      { key: PREVIOUS_STORAGE_KEYS.weapons, source: "current" },
      { key: LEGACY_STORAGE_KEYS.weapons, source: "legacy" },
    ],
    normalizeWeaponItem,
    "weapon",
  );
}

function loadUptimesForCurrentVersion() {
  return loadMigratedObject(
    [
      { key: STORAGE_KEYS.uptimes, source: "current" },
      { key: PREVIOUS_STORAGE_KEYS.uptimes, source: "previous" },
      { key: LEGACY_STORAGE_KEYS.uptimes, source: "legacy" },
    ],
    normalizeBuildUptimes,
    "uptime",
  );
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
  if (Array.isArray(state.data?.uptimeFields)) {
    return state.data.uptimeFields.map((field) => ({
      ...field,
      label: String(field.label),
      defaultValue: typeof field.defaultValue === "number" ? field.defaultValue : Number(field.defaultValue) || 0,
      displayScale: field.displayScale ?? 100,
    }));
  }

  const sheet = state.data?.sheets?.[calculatorSheetName()];
  if (!sheet) {
    return [];
  }

  const fields = [];
  const remainingHealthLabel = sheet[14]?.[3];
  fields.push({
    ref: "E15",
    label: String(remainingHealthLabel || "Remaining Health"),
    defaultValue: 100,
    displayScale: 1,
    maxValue: 160,
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

function getUptimeMinValue(field) {
  return field?.minValue ?? (field?.displayScale === 1 ? 1 : 0);
}

function getUptimeMaxValue(field) {
  return field?.maxValue ?? 100;
}

function getDefaultUptimeFeedback(field, value) {
  if (!isUsingDefaultUptime(field, value)) {
    return "";
  }
  return field?.defaultFeedback ?? "Using Krea default";
}

function updateUptimeFeedback(input, field) {
  const ref = input.dataset.uptimeField;
  const numeric = Number(input.value);
  const minValue = getUptimeMinValue(field);
  const maxValue = getUptimeMaxValue(field);
  const isInvalid = input.value.trim() !== "" && (!Number.isFinite(numeric) || numeric < minValue || numeric > maxValue);
  const feedback = els.riftModalContent.querySelector(`[data-uptime-feedback="${ref}"]`);

  input.classList.toggle("uptime-input-invalid", isInvalid);
  input.setAttribute("aria-invalid", isInvalid ? "true" : "false");

  if (!feedback) {
    return;
  }

  feedback.classList.toggle("uptime-validation-error", isInvalid);
  feedback.textContent = isInvalid
    ? `Only values ${minValue}-${maxValue} are valid.`
    : getDefaultUptimeFeedback(field, state.uptimeDraft[ref]);
}

function persistUptimes() {
  localStorage.setItem(
    STORAGE_KEYS.uptimes,
    JSON.stringify(semanticValuesFromCurrent(state.uptimeValues, state.uptimeFields)),
  );
}

function setAllBuildsCompareEnabled(compareEnabled) {
  state.builds = state.builds.map((build) => ({ ...build, compareEnabled }));
  persistBuilds();
  renderBuildList();
}

function setAllWeaponsCompareEnabled(compareEnabled) {
  state.weapons = state.weapons.map((weapon) => ({ ...weapon, compareEnabled }));
  persistWeapons();
  renderWeaponList();
}

function getAvailableWeaponTypes() {
  const availableTypes = new Set(
    state.weapons
      .map((weapon) => String(weapon.values?.E7 ?? "").trim())
      .filter(Boolean),
  );
  const orderedTypes = (weaponSchemaMap().E7?.options ?? [])
    .map((option) => String(option))
    .filter((option) => availableTypes.has(option));
  const extraTypes = [...availableTypes]
    .filter((type) => !orderedTypes.includes(type))
    .sort((a, b) => a.localeCompare(b));

  return [...orderedTypes, ...extraTypes];
}

function renderWeaponTypeShortcut() {
  if (!els.selectWeaponType) {
    return;
  }

  const weaponTypes = getAvailableWeaponTypes();
  els.selectWeaponType.disabled = weaponTypes.length === 0;
  els.selectWeaponType.innerHTML = [
    `<option value="">Type...</option>`,
    ...weaponTypes.map((weaponType) => `<option value="${escapeHtml(weaponType)}">${escapeHtml(weaponType)}</option>`),
  ].join("");
}

function selectWeaponsByType(weaponType) {
  state.weapons = state.weapons.map((weapon) => ({
    ...weapon,
    compareEnabled: String(weapon.values?.E7 ?? "") === weaponType,
  }));
  persistWeapons();
  renderWeaponList();
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

function normalizeRuntimeBuild(build) {
  const base = buildDefaultBuild();
  return {
    ...base,
    ...build,
    values: { ...base.values, ...build.values },
    compareEnabled: typeof build.compareEnabled === "boolean" ? build.compareEnabled : true,
  };
}

function normalizeRuntimeWeapon(weapon) {
  const base = buildDefaultWeapon();
  return {
    ...base,
    ...weapon,
    values: { ...base.values, ...weapon.values },
    isRift: Boolean(weapon.isRift),
    compareEnabled: typeof weapon.compareEnabled === "boolean" ? weapon.compareEnabled : true,
    libraryId: weapon.libraryId ? String(weapon.libraryId) : null,
    libraryVariant: weapon.libraryVariant ? String(weapon.libraryVariant) : null,
    riftLevel:
      weapon.riftLevel != null && Number.isFinite(Number(weapon.riftLevel))
        ? Number(weapon.riftLevel)
        : null,
    riftNodesApplied:
      typeof weapon.riftNodesApplied === "boolean" ? weapon.riftNodesApplied : null,
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
    if (field.extension) {
      continue;
    }
    writeCell(engine, calculatorSheetName(), field.ref, build.values[field.ref]);
  }
  for (const field of state.data.weaponFields) {
    writeCell(engine, calculatorSheetName(), field.ref, weapon.values[field.ref]);
  }
  for (const field of state.uptimeFields) {
    if (field.extension) {
      continue;
    }
    writeCell(engine, calculatorSheetName(), field.ref, state.uptimeValues[field.ref]);
  }

  if (state.extensionStatus?.enabled) {
    const extensions = window.PHASK_SKILL_EXTENSIONS;
    const modifiers = extensions.calculateScenarioModifiers(
      build.values,
      weapon.values,
      state.uptimeValues,
    );
    writeCell(
      engine,
      extensions.extensionSheet,
      extensions.targetCells.velkhanaAegis,
      modifiers.velkhanaAegis,
    );
    writeCell(
      engine,
      extensions.extensionSheet,
      extensions.targetCells.meditation,
      modifiers.meditation,
    );
    writeCell(
      engine,
      extensions.extensionSheet,
      extensions.targetCells.blastExploit,
      modifiers.blastExploit,
    );
  }
}

function getBuildLabels(build, weapon) {
  const engine = createEngine();
  applyScenario(engine, build, weapon);
  const labels = {};
  for (const field of state.data.buildFields) {
    labels[field.ref] = field.extension
      ? field.label
      : readCell(engine, calculatorSheetName(), field.labelRef);
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
    labels[field.ref] = readCell(engine, calculatorSheetName(), field.labelRef);
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
    h12: readCell(engine, calculatorSheetName(), state.data.resultCell),
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
  return Object.fromEntries(refs.map((ref) => [ref, readCell(engine, calculatorSheetName(), ref)]));
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

  const extensionWarning = state.extensionStatus?.error
    ? `<div class="calculation-warning">${escapeHtml(state.extensionStatus.error)} Local skill extensions are disabled until their workbook adapter is updated.</div>`
    : "";

  els.resultGrid.innerHTML = `
    <div class="result-card">
      <div class="label">Effective Damage</div>
      <div class="value">${escapeHtml(formatResult(result.h12))}</div>
    </div>
    ${extensionWarning}
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
      <button id="open-weapon-library" type="button">Weapon Library</button>
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
  els.selectionActions
    .querySelector("#open-weapon-library")
    .addEventListener("click", () => {
      openWeaponLibrary();
    });
}

function getWeaponLibraryEntries() {
  const weapons = Array.isArray(window.WEAPON_LIBRARY?.weapons)
    ? window.WEAPON_LIBRARY.weapons
    : [];
  return weapons.filter(
    (weapon) =>
      weapon.weaponTypeKey !== "InsectGlaive" &&
      (weapon.weaponTypeKey === "ChargeBlade" || WEAPON_LIBRARY_TYPE_MAP[weapon.weaponTypeKey]),
  );
}

function mapLibraryWeaponType(weapon) {
  if (weapon.weaponTypeKey === "ChargeBlade") {
    return weapon.mechanics?.phial === "Impact Phial"
      ? "Charge Blade (Impact)"
      : "Charge Blade (Power)";
  }
  return WEAPON_LIBRARY_TYPE_MAP[weapon.weaponTypeKey] ?? null;
}

function getLibraryEntry(libraryId) {
  return getWeaponLibraryEntries().find((weapon) => weapon.id === libraryId) ?? null;
}

function getLibraryRiftLevel(weapon) {
  const level = Number(weapon.customization?.max_level);
  return Number.isFinite(level) && level > 0 ? level : null;
}

function getRiftFixedStats(weapon) {
  const maxLevel = getLibraryRiftLevel(weapon);
  const totals = {
    attack: Number(weapon.attack) || 0,
    attribute: Number(weapon.attribute?.value) || 0,
    affinityPercent: Number(weapon.affinityPercent) || 0,
  };
  if (!maxLevel) {
    return totals;
  }
  for (const rule of weapon.customization?.level_rules ?? []) {
    if (Number(rule.level_no) > maxLevel) {
      continue;
    }
    totals.attack += Number(rule.attack_add) || 0;
    totals.attribute += Number(rule.element_add) || 0;
    totals.affinityPercent += Number(rule.affinity_add) || 0;
  }
  return totals;
}

function getLibraryVariantKey(weapon, useRiftMaximum) {
  return useRiftMaximum
    ? `rift-lv${getLibraryRiftLevel(weapon)}-no-nodes`
    : "g10-5";
}

function buildWeaponFromLibrary(weapon, useRiftMaximum = false) {
  const riftLevel = useRiftMaximum ? getLibraryRiftLevel(weapon) : null;
  const stats = useRiftMaximum
    ? getRiftFixedStats(weapon)
    : {
        attack: Number(weapon.attack) || 0,
        attribute: Number(weapon.attribute?.value) || 0,
        affinityPercent: Number(weapon.affinityPercent) || 0,
      };
  const damageType = WEAPON_LIBRARY_ATTRIBUTE_LABELS[weapon.attribute?.type] ?? "Raw";
  return {
    id: makeId(),
    name: useRiftMaximum
      ? `${weapon.name} (Rift Lv${riftLevel}, No Nodes)`
      : weapon.name,
    compareEnabled: true,
    isRift: useRiftMaximum,
    libraryId: weapon.id,
    libraryVariant: getLibraryVariantKey(weapon, useRiftMaximum),
    riftLevel,
    riftNodesApplied: useRiftMaximum ? false : null,
    values: {
      E3: stats.attack,
      E4: stats.attribute,
      E5: damageType,
      E6: stats.affinityPercent / 100,
      E7: mapLibraryWeaponType(weapon),
    },
  };
}

function isLibraryVariantSaved(weapon, useRiftMaximum = false) {
  const variant = getLibraryVariantKey(weapon, useRiftMaximum);
  const candidate = buildWeaponFromLibrary(weapon, useRiftMaximum);
  if (
    state.weapons.some(
      (saved) =>
        saved.libraryId === weapon.id &&
        saved.libraryVariant === variant &&
        Boolean(saved.isRift) === Boolean(candidate.isRift) &&
        stableStringify(saved.values) === stableStringify(candidate.values),
    )
  ) {
    return true;
  }
  const candidateFingerprint = weaponImportFingerprint(candidate);
  return state.weapons.some(
    (saved) => weaponImportFingerprint(saved) === candidateFingerprint,
  );
}

function addWeaponFromLibrary(weapon, useRiftMaximum, button) {
  if (isLibraryVariantSaved(weapon, useRiftMaximum)) {
    button.textContent = "Added";
    button.disabled = true;
    button.classList.add("weapon-library-add-saved");
    return;
  }

  const savedWeapon = buildWeaponFromLibrary(weapon, useRiftMaximum);
  state.weapons = [...state.weapons, savedWeapon];
  state.selectedWeaponId = savedWeapon.id;
  saveStoredValue(STORAGE_KEYS.selectedWeaponId, state.selectedWeaponId);
  persistWeapons();
  renderCalculatorSelectors();
  renderResultGrid();
  renderCalculatorActions();
  renderSkillSummary();
  renderWeaponList();
  renderWeaponTypeShortcut();

  button.textContent = "Added";
  button.disabled = true;
  button.classList.add("weapon-library-add-saved");
  button.closest(".weapon-library-card")?.classList.add("weapon-library-card-added");
  const feedback = els.riftModalContent.querySelector("#weapon-library-feedback");
  if (feedback) {
    feedback.textContent = `${savedWeapon.name} was added to your Weapons list.`;
  }
}

function getFilteredWeaponLibraryEntries() {
  const query = state.weaponLibraryQuery.trim().toLocaleLowerCase();
  return getWeaponLibraryEntries().filter((weapon) => {
    if (
      state.weaponLibraryType !== "all" &&
      weapon.weaponTypeKey !== state.weaponLibraryType
    ) {
      return false;
    }
    if (
      state.weaponLibraryAttribute !== "all" &&
      weapon.attribute?.type !== state.weaponLibraryAttribute
    ) {
      return false;
    }
    if (!query) {
      return true;
    }
    const searchable = [
      weapon.name,
      weapon.monster,
      weapon.series?.name,
      weapon.weaponType,
      WEAPON_LIBRARY_ATTRIBUTE_LABELS[weapon.attribute?.type],
      ...(weapon.skills ?? []).map((skill) => skill.name),
    ]
      .filter(Boolean)
      .join(" ")
      .toLocaleLowerCase();
    return searchable.includes(query);
  });
}

function renderWeaponLibraryResults() {
  const results = els.riftModalContent.querySelector("#weapon-library-results");
  const count = els.riftModalContent.querySelector("#weapon-library-count");
  if (!results || !count) {
    return;
  }
  const weapons = getFilteredWeaponLibraryEntries();
  count.textContent = `${weapons.length} weapon${weapons.length === 1 ? "" : "s"}`;
  if (!weapons.length) {
    results.innerHTML = `<div class="weapon-library-empty">No weapons match these filters.</div>`;
    return;
  }

  results.innerHTML = weapons
    .map((weapon) => {
      const riftLevel = getLibraryRiftLevel(weapon);
      const baseSaved = isLibraryVariantSaved(weapon, false);
      const riftSaved = riftLevel ? isLibraryVariantSaved(weapon, true) : false;
      const attributeLabel = WEAPON_LIBRARY_ATTRIBUTE_LABELS[weapon.attribute?.type] ?? "Raw";
      const attributeMarkup =
        weapon.attribute?.category === "raw"
          ? ""
          : `<span class="weapon-library-stat weapon-library-attribute-${escapeHtml(weapon.attribute.type)}">${escapeHtml(attributeLabel)} ${escapeHtml(weapon.attribute.value)}</span>`;
      const skills = (weapon.skills ?? [])
        .map((skill) => `${skill.name} Lv${skill.level}`)
        .join(" · ");
      const originName = weapon.monster ?? weapon.series?.name ?? "Unknown";
      return `
        <article class="weapon-library-card ${baseSaved && (!riftLevel || riftSaved) ? "weapon-library-card-added" : ""}">
          <div class="weapon-library-card-header">
            <div>
              <h3>${escapeHtml(originName)} ${escapeHtml(weapon.weaponType)}</h3>
              <div class="weapon-library-series">${escapeHtml(weapon.name)}</div>
            </div>
            ${riftLevel ? `<span class="weapon-library-rift-badge">Rift Lv${riftLevel}</span>` : ""}
          </div>
          <div class="weapon-library-stats">
            <span class="weapon-library-stat">Attack ${escapeHtml(weapon.attack)}</span>
            ${attributeMarkup}
            <span class="weapon-library-stat">Affinity ${escapeHtml(weapon.affinityPercent)}%</span>
          </div>
          <div class="weapon-library-skill">${escapeHtml(skills || "No equipment skill")}</div>
          <div class="weapon-library-card-actions ${riftLevel ? "weapon-library-card-actions-rift" : ""}">
            ${riftLevel ? `<span class="weapon-library-add-label">Add G10.5:</span>` : ""}
            <button class="${baseSaved ? "weapon-library-add-saved" : ""}" type="button" data-library-id="${escapeHtml(weapon.id)}" data-library-variant="base" aria-label="${escapeHtml(riftLevel ? `Add ${weapon.name} at G10.5, Rift level 0` : `Add ${weapon.name} at G10.5`)}" ${baseSaved ? "disabled" : ""}>${baseSaved ? "Added" : riftLevel ? "Rift 0" : "Add G10.5"}</button>
            ${
              riftLevel
                ? `<button class="${riftSaved ? "weapon-library-add-saved" : ""}" type="button" data-library-id="${escapeHtml(weapon.id)}" data-library-variant="rift" aria-label="${escapeHtml(`Add ${weapon.name} at G10.5, Rift level ${riftLevel}`)}" ${riftSaved ? "disabled" : ""}>${riftSaved ? "Added" : `Rift ${riftLevel}`}</button>`
                : ""
            }
          </div>
        </article>
      `;
    })
    .join("");

  results.querySelectorAll("[data-library-id]").forEach((button) => {
    button.addEventListener("click", () => {
      const weapon = getLibraryEntry(button.dataset.libraryId);
      if (!weapon) {
        return;
      }
      addWeaponFromLibrary(weapon, button.dataset.libraryVariant === "rift", button);
    });
  });
}

function renderWeaponLibrary() {
  const entries = getWeaponLibraryEntries();
  const typeMetadata = (window.WEAPON_LIBRARY?.weaponTypes ?? []).filter(
    (type) => type.key !== "InsectGlaive" && entries.some((weapon) => weapon.weaponTypeKey === type.key),
  );
  const attributeTypes = Object.keys(WEAPON_LIBRARY_ATTRIBUTE_LABELS).filter((type) =>
    entries.some((weapon) => weapon.attribute?.type === type),
  );

  openModal({
    title: "Weapon Library",
    content: `
      <div class="weapon-library-shell">
        <div class="weapon-library-filter-bar">
          <div class="weapon-library-toolbar">
            <label class="weapon-library-search">
              <span class="visually-hidden">Search weapons</span>
              <input id="weapon-library-search" type="search" value="${escapeHtml(state.weaponLibraryQuery)}" placeholder="Search weapon, series, or skill" />
            </label>
            <label class="weapon-library-attribute-filter">
              <span class="visually-hidden">Filter by damage type</span>
              <select id="weapon-library-attribute">
                <option value="all">All damage types</option>
                ${attributeTypes
                  .map(
                    (type) =>
                      `<option value="${escapeHtml(type)}" ${state.weaponLibraryAttribute === type ? "selected" : ""}>${escapeHtml(WEAPON_LIBRARY_ATTRIBUTE_LABELS[type])}</option>`,
                  )
                  .join("")}
              </select>
            </label>
          </div>
          <div class="weapon-library-types" role="group" aria-label="Weapon type">
            <button class="weapon-library-type ${state.weaponLibraryType === "all" ? "active" : ""}" type="button" data-library-type="all">All</button>
            ${typeMetadata
              .map(
                (type) =>
                  `<button class="weapon-library-type ${state.weaponLibraryType === type.key ? "active" : ""}" type="button" data-library-type="${escapeHtml(type.key)}">${escapeHtml(type.name)}</button>`,
              )
              .join("")}
          </div>
          <label class="weapon-library-mobile-type">
            <span>Weapon Type</span>
            <select id="weapon-library-type-select">
              <option value="all">All weapon types</option>
              ${typeMetadata
                .map(
                  (type) =>
                    `<option value="${escapeHtml(type.key)}" ${state.weaponLibraryType === type.key ? "selected" : ""}>${escapeHtml(type.name)}</option>`,
                )
                .join("")}
            </select>
          </label>
        </div>
        <div class="weapon-library-summary">
          <span id="weapon-library-count"></span>
          <span id="weapon-library-feedback" class="weapon-library-feedback" role="status" aria-live="polite"></span>
        </div>
        <div class="weapon-library-grid" id="weapon-library-results"></div>
      </div>
    `,
  });

  const search = els.riftModalContent.querySelector("#weapon-library-search");
  search.addEventListener("input", () => {
    state.weaponLibraryQuery = search.value;
    renderWeaponLibraryResults();
  });
  els.riftModalContent.querySelector("#weapon-library-attribute").addEventListener("change", (event) => {
    state.weaponLibraryAttribute = event.target.value;
    renderWeaponLibraryResults();
  });
  els.riftModalContent.querySelectorAll("[data-library-type]").forEach((button) => {
    button.addEventListener("click", () => {
      state.weaponLibraryType = button.dataset.libraryType;
      const mobileSelect = els.riftModalContent.querySelector("#weapon-library-type-select");
      if (mobileSelect) {
        mobileSelect.value = state.weaponLibraryType;
      }
      els.riftModalContent.querySelectorAll("[data-library-type]").forEach((item) => {
        item.classList.toggle("active", item === button);
      });
      renderWeaponLibraryResults();
    });
  });
  els.riftModalContent.querySelector("#weapon-library-type-select").addEventListener("change", (event) => {
    state.weaponLibraryType = event.target.value;
    els.riftModalContent.querySelectorAll("[data-library-type]").forEach((item) => {
      item.classList.toggle("active", item.dataset.libraryType === state.weaponLibraryType);
    });
    renderWeaponLibraryResults();
  });
  renderWeaponLibraryResults();
  requestAnimationFrame(() => search.focus());
}

function openWeaponLibrary() {
  if (!window.WEAPON_LIBRARY) {
    openModal({
      title: "Weapon Library",
      mode: "medium",
      content: `<div class="comparison-intro">The weapon catalog could not be loaded.</div>`,
    });
    return;
  }
  renderWeaponLibrary();
}

function openRiftUnavailableMessage() {
  openModal({
    title: "Rift Combinations",
    mode: "medium",
    content: `
    <div class="comparison-intro">This option is only available with rift weapons that do not have node upgrades applied.</div>
  `,
  });
}

function openUptimesModal() {
  state.uptimeDraft = { ...state.uptimeValues };
  renderUptimesModal();
}

function renderUptimesModal() {
  if (!state.uptimeDraft) {
    return;
  }

  openModal({
    title: "View/Edit Uptimes",
    mode: "uptime",
    content: `
    <div class="editor-actions-row uptime-actions-row">
      <button type="button" id="save-uptimes">Save Values</button>
      <button class="secondary" type="button" id="revert-uptimes">Revert to Defaults</button>
      <button class="secondary" type="button" id="discard-uptimes">Discard</button>
    </div>
    <div class="uptime-list">
      ${state.uptimeFields
        .map((field) => {
          const minValue = getUptimeMinValue(field);
          const maxValue = getUptimeMaxValue(field);
          return `
          <label class="uptime-row" for="uptime-${field.ref}">
            <span class="uptime-label">${escapeHtml(field.label)}</span>
            <input id="uptime-${field.ref}" type="number" min="${minValue}" max="${maxValue}" step="${field.step ?? (field.displayScale === 1 ? "1" : "0.1")}" data-uptime-field="${field.ref}" aria-invalid="false" aria-describedby="uptime-feedback-${field.ref}${field.description ? ` uptime-description-${field.ref}` : ""}" value="${escapeHtml(
              field.displayScale === 1
                ? String(Math.round(state.uptimeDraft[field.ref] ?? 0))
                : ((state.uptimeDraft[field.ref] ?? 0) * (field.displayScale ?? 100)).toFixed(1),
            )}" />
            <span class="uptime-unit">${field.displayScale === 1 ? "" : "%"}</span>
            <span id="uptime-feedback-${field.ref}" class="uptime-feedback" data-uptime-feedback="${field.ref}">${escapeHtml(getDefaultUptimeFeedback(field, state.uptimeDraft[field.ref]))}</span>
            ${field.description ? `<span id="uptime-description-${field.ref}" class="uptime-description">${escapeHtml(field.description)}</span>` : ""}
          </label>
        `;
        })
        .join("")}
    </div>
  `,
  });

  els.riftModalContent.querySelectorAll("[data-uptime-field]").forEach((input) => {
    input.addEventListener("input", (event) => {
      const ref = event.target.dataset.uptimeField;
      const field = state.uptimeFields.find((item) => item.ref === ref);
      const numeric = Number(event.target.value);
      const minValue = getUptimeMinValue(field);
      const maxValue = getUptimeMaxValue(field);
      const fallbackValue = minValue;
      const boundedValue = Number.isFinite(numeric) ? Math.min(maxValue, Math.max(minValue, numeric)) : fallbackValue;
      state.uptimeDraft[ref] = boundedValue / (field?.displayScale ?? 100);
      updateUptimeFeedback(event.target, field);
    });
    const field = state.uptimeFields.find((item) => item.ref === input.dataset.uptimeField);
    updateUptimeFeedback(input, field);
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

function renderLibraryList(
  targetEl,
  items,
  selectedId,
  editHandlerName,
  deleteHandlerName,
  moveUpHandlerName,
  moveDownHandlerName,
) {
  if (!items.length) {
    targetEl.innerHTML = `<div class="empty-state">Nothing saved yet.</div>`;
    return;
  }

  targetEl.innerHTML = items
    .map(
      (item, index) => {
        const escapedName = escapeHtml(item.name);
        return `
        <div class="library-item ${item.id === selectedId ? "selected" : ""}">
          <div class="library-item-main">
            <label class="library-item-check">
              <input type="checkbox" data-action="toggle-compare" data-id="${item.id}" ${item.compareEnabled !== false ? "checked" : ""} />
            </label>
            <div class="library-item-name">${escapedName}</div>
          </div>
          <div class="library-item-actions">
            <div class="library-reorder-actions">
              <button class="secondary library-reorder-button" type="button" data-action="${moveUpHandlerName}" data-id="${item.id}" aria-label="Move ${escapedName} up" title="Move up" ${index === 0 ? "disabled" : ""}>↑</button>
              <button class="secondary library-reorder-button" type="button" data-action="${moveDownHandlerName}" data-id="${item.id}" aria-label="Move ${escapedName} down" title="Move down" ${index === items.length - 1 ? "disabled" : ""}>↓</button>
            </div>
            <button class="secondary" type="button" data-action="${editHandlerName}" data-id="${item.id}">Edit</button>
            <button class="danger" type="button" data-action="${deleteHandlerName}" data-id="${item.id}">Del</button>
          </div>
        </div>
      `;
      },
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
    "move-build-up",
    "move-build-down",
  );
}

function renderWeaponList() {
  renderLibraryList(
    els.weaponList,
    state.weapons,
    state.selectedWeaponId,
    "edit-weapon",
    "delete-weapon",
    "move-weapon-up",
    "move-weapon-down",
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
  const sortedFields = sortBuildFieldsAlphabetically(state.data.buildFields, labels);

  const fieldsMarkup = orderBuildFieldsByVisibleColumn(sortedFields)
    .map((field) => {
      if (!field) {
        return '<div class="editor-row editor-row-placeholder" aria-hidden="true"></div>';
      }
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
      const options = field.options;
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
      <label class="editor-label" for="weapon-is-rift">Rift weapon (node upgrades not applied)</label>
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
  saveStoredItems(STORAGE_KEYS.builds, state.builds.map(serializeStoredBuild));
}

function persistWeapons() {
  saveStoredItems(STORAGE_KEYS.weapons, state.weapons.map(serializeStoredWeapon));
}

function moveLibraryItem(items, id, direction) {
  const currentIndex = items.findIndex((item) => item.id === id);
  const nextIndex = currentIndex + direction;
  if (currentIndex < 0 || nextIndex < 0 || nextIndex >= items.length) {
    return items;
  }

  const reordered = [...items];
  [reordered[currentIndex], reordered[nextIndex]] = [reordered[nextIndex], reordered[currentIndex]];
  return reordered;
}

function moveBuild(id, direction) {
  const reordered = moveLibraryItem(state.builds, id, direction);
  if (reordered === state.builds) {
    return;
  }

  state.builds = reordered;
  persistBuilds();
  renderAll();
}

function moveWeapon(id, direction) {
  const reordered = moveLibraryItem(state.weapons, id, direction);
  if (reordered === state.weapons) {
    return;
  }

  state.weapons = reordered;
  persistWeapons();
  renderAll();
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

function nextCopyName(name, fallbackName) {
  const currentName = String(name || fallbackName);
  const copyMatch = /^(.*) \(Copy(?: (\d+))?\)$/.exec(currentName);
  if (!copyMatch) {
    return `${currentName} (Copy)`;
  }

  const nextNumber = copyMatch[2] ? Number(copyMatch[2]) + 1 : 2;
  return `${copyMatch[1]} (Copy ${nextNumber})`;
}

function duplicateCurrentBuildDraft() {
  if (!state.buildDraft) {
    return;
  }

  const payload = {
    ...deepClone(state.buildDraft),
    id: makeId(),
    name: nextCopyName(state.buildDraft.name, "Unnamed Build"),
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
    name: nextCopyName(state.weaponDraft.name, "Unnamed Weapon"),
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
  renderWeaponTypeShortcut();
  renderBuildForm();
  renderWeaponForm();
}

function generateRiftVariants(weapon) {
  const libraryWeapon = weapon.libraryId ? getLibraryEntry(weapon.libraryId) : null;
  const customization = libraryWeapon?.customization;
  const milestoneCount = customization
    ? (customization.milestone_levels ?? []).filter(
        (level) => Number(level) <= Number(weapon.riftLevel ?? customization.max_level),
      ).length
    : 3;
  const choices = customization
    ? (customization.choices ?? []).map((choice) => ({
        key: String(choice.key).toLocaleLowerCase(),
        label:
          choice.key === "Element" && weapon.values.E5 !== "Raw"
            ? weapon.values.E5
            : choice.label_en,
        value: Number(choice.value) || 0,
      }))
    : [
        { key: "attack", label: "Attack", value: 100 },
        ...(weapon.values.E5 === "Raw"
          ? []
          : [{ key: "element", label: weapon.values.E5, value: 100 }]),
        { key: "critical", label: "Affinity", value: 10 },
      ];
  if (!milestoneCount || !choices.length) {
    return [];
  }

  const allocations = [];

  function helper(index, remaining, counts) {
    if (index === choices.length - 1) {
      counts[choices[index].key] = remaining;
      allocations.push({ ...counts });
      return;
    }
    for (let count = 0; count <= remaining; count += 1) {
      counts[choices[index].key] = count;
      helper(index + 1, remaining - count, counts);
    }
  }

  helper(0, milestoneCount, {});

  return allocations.map((counts) => {
    const nextWeapon = deepClone(weapon);
    const labels = [];
    const sequence = [];
    for (const choice of choices) {
      const count = counts[choice.key] ?? 0;
      if (!count) {
        continue;
      }
      const total = count * choice.value;
      if (choice.key === "attack") {
        nextWeapon.values.E3 = Number(nextWeapon.values.E3) + total;
      } else if (choice.key === "element") {
        nextWeapon.values.E4 = Number(nextWeapon.values.E4) + total;
      } else if (choice.key === "critical") {
        nextWeapon.values.E6 = Number(nextWeapon.values.E6) + total / 100;
      }
      labels.push(`+${total}${choice.key === "critical" ? "%" : ""} ${choice.label}`);
      for (let index = 0; index < count; index += 1) {
        sequence.push(String(choice.label).toLocaleLowerCase());
      }
    }
    return {
      counts,
      weapon: nextWeapon,
      label: labels.join(", "),
      upgradeSequence: sequence.join(", "),
    };
  });
}

function buildSavedRiftWeapon(baseWeapon, variant) {
  const nodeKey = Object.entries(variant.counts)
    .filter(([, count]) => count)
    .map(([key, count]) => `${key}-${count}`)
    .join("-");
  return {
    ...deepClone(variant.weapon),
    name: `${baseWeapon.name} (${variant.label || "Rift"})`,
    isRift: false,
    libraryVariant: baseWeapon.libraryVariant
      ? `${baseWeapon.libraryVariant}-nodes-${nodeKey || "none"}`
      : null,
    riftNodesApplied: true,
    compareEnabled: true,
  };
}

function isRiftVariantSaved(baseWeapon, variant) {
  const fingerprint = weaponImportFingerprint(buildSavedRiftWeapon(baseWeapon, variant));
  return state.weapons.some((weapon) => weaponImportFingerprint(weapon) === fingerprint);
}

function markRiftVariantSaved(button) {
  button.textContent = "Saved";
  button.disabled = true;
  button.classList.add("rift-save-button-saved");
  button.closest(".rift-result-item")?.classList.add("rift-result-item-saved");
}

function saveRiftVariant(baseWeapon, variant, button) {
  const savedWeapon = buildSavedRiftWeapon(baseWeapon, variant);
  const fingerprint = weaponImportFingerprint(savedWeapon);
  const alreadySaved = state.weapons.some((weapon) => weaponImportFingerprint(weapon) === fingerprint);

  if (!alreadySaved) {
    state.weapons = [...state.weapons, { ...savedWeapon, id: makeId() }];
    persistWeapons();
    renderCalculatorSelectors();
    renderWeaponList();
    renderWeaponTypeShortcut();
  }

  markRiftVariantSaved(button);
}

function openRiftComparison() {
  const build = getBuildById(state.selectedBuildId);
  const weapon = getWeaponById(state.selectedWeaponId);
  if (!build || !weapon || !weapon.isRift) {
    return;
  }

  const engine = createEngine();
  const generatedVariants = generateRiftVariants(weapon);
  if (!generatedVariants.length) {
    openModal({
      title: "Rift Combinations",
      mode: "medium",
      content: `<div class="comparison-intro">No node upgrade choices are available for this weapon at its saved rift level.</div>`,
    });
    return;
  }
  const variants = generatedVariants
    .map((variant) => {
      applyScenario(engine, build, variant.weapon);
      return {
        ...variant,
        h12: readCell(engine, calculatorSheetName(), state.data.resultCell),
        isSaved: isRiftVariantSaved(weapon, variant),
      };
    })
    .sort((a, b) => {
      const av = typeof a.h12 === "number" ? a.h12 : -Infinity;
      const bv = typeof b.h12 === "number" ? b.h12 : -Infinity;
      return bv - av;
    });
  const topValue = variants.find((variant) => typeof variant.h12 === "number")?.h12;

  openModal({
    title: "Rift Combinations",
    mode: "compact",
    content: `
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
            <div class="rift-result-item ${variant.isSaved ? "rift-result-item-saved" : ""}">
              <div>
                <div>${escapeHtml(variant.label || "No bonus")}</div>
                <div class="rift-result-label">${escapeHtml(variant.upgradeSequence || "attack, attack, attack")}</div>
              </div>
              <div class="rift-result-actions">
                <div class="rift-result-value-block">
                  <div class="rift-result-value">${escapeHtml(formatResult(variant.h12))}</div>
                  <div class="rift-result-relative">${escapeHtml(relativeLabel)}</div>
                </div>
                <button class="secondary rift-save-button ${variant.isSaved ? "rift-save-button-saved" : ""}" type="button" data-rift-variant-index="${index}" ${variant.isSaved ? "disabled" : ""}>${variant.isSaved ? "Saved" : "Save Weapon"}</button>
              </div>
            </div>
          `;
          },
        )
        .join("")}
    </div>
  `,
  });
  els.riftModalContent.querySelectorAll("[data-rift-variant-index]").forEach((button) => {
    button.addEventListener("click", () => {
      const variant = variants[Number(button.dataset.riftVariantIndex)];
      if (!variant) {
        return;
      }
      saveRiftVariant(weapon, variant, button);
    });
  });
}

function buildMatrixKey(buildId, weaponId) {
  return `${weaponId}::${buildId}`;
}

function openBuildWeaponComparison() {
  const buildsToCompare = state.builds.filter((build) => build.compareEnabled !== false);
  const weaponsToCompare = state.weapons.filter((weapon) => weapon.compareEnabled !== false);
  if (!buildsToCompare.length || !weaponsToCompare.length) {
    openModal({
      title: "Build and Weapon Comparison",
      mode: "medium",
      content: `
      <div class="comparison-intro">Select at least one build and one weapon with the checkboxes before running the comparison.</div>
    `,
    });
    return;
  }

  const engine = createEngine();
  const results = {};
  for (const weapon of weaponsToCompare) {
    for (const build of buildsToCompare) {
      applyScenario(engine, build, weapon);
      results[buildMatrixKey(build.id, weapon.id)] = readCell(engine, calculatorSheetName(), state.data.resultCell);
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
  let exportSummary = "Exporting: 0 builds, 0 weapons";
  try {
    const parsed = JSON.parse(payload);
    const buildCount = Array.isArray(parsed?.builds) ? parsed.builds.length : 0;
    const weaponCount = Array.isArray(parsed?.weapons) ? parsed.weapons.length : 0;
    exportSummary = `Exporting: ${buildCount} builds, ${weaponCount} weapons`;
  } catch {}

  openModal({
    title: "Export Data",
    mode: "medium",
    content: `
    <div class="transfer-copy">
      <p class="comparison-intro">${escapeHtml(description)}</p>
      <div class="transfer-summary">${escapeHtml(exportSummary)}</div>
      <textarea class="transfer-textarea" id="export-payload" readonly></textarea>
      <div class="transfer-actions">
        <button type="button" id="copy-export-payload">Copy</button>
      </div>
      <div class="transfer-feedback" id="transfer-feedback"></div>
    </div>
  `,
  });
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
}

function openImportModal() {
  openModal({
    title: "Import Data",
    mode: "medium",
    content: `
    <div class="transfer-copy">
      <p class="comparison-intro">Paste an exported string below. Imported builds and weapons will be added to the current data.</p>
      <textarea class="transfer-textarea" id="import-payload" placeholder="Paste export string here"></textarea>
      <div class="transfer-actions">
        <button type="button" id="submit-import-payload">Import</button>
      </div>
      <div class="transfer-feedback" id="transfer-feedback"></div>
    </div>
  `,
  });
  els.riftModalContent.querySelector("#submit-import-payload").addEventListener("click", () => {
    const payload = els.riftModalContent.querySelector("#import-payload").value.trim();
    const feedback = els.riftModalContent.querySelector("#transfer-feedback");
    if (!payload) {
      feedback.textContent = "Paste an export string first.";
      return;
    }

    try {
      const parsed = JSON.parse(payload);
      const importSource = parsed?.formatVersion === EXPORT_FORMAT_VERSION ? "semantic" : "legacy";
      const importWarnings = [];
      const existingBuildFingerprints = new Set(state.builds.map(buildImportFingerprint));
      const existingWeaponFingerprints = new Set(state.weapons.map(weaponImportFingerprint));
      const importedBuilds = [];
      const importedWeapons = [];

      if (Array.isArray(parsed.builds)) {
        for (const rawBuild of parsed.builds) {
          const normalizedBuild = normalizeImportedBuild(rawBuild, importWarnings, importSource);
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
          const normalizedWeapon = normalizeImportedWeapon(rawWeapon, importWarnings, importSource);
          const fingerprint = weaponImportFingerprint(normalizedWeapon);
          if (existingWeaponFingerprints.has(fingerprint)) {
            continue;
          }
          existingWeaponFingerprints.add(fingerprint);
          importedWeapons.push(normalizedWeapon);
        }
      }

      if (!importedBuilds.length && !importedWeapons.length) {
        feedback.textContent = importWarnings.length
          ? `No new builds or weapons were imported. Warning: these fields could not be imported: ${importWarnings.join(", ")}.`
          : "No new builds or weapons were imported.";
        return;
      }

      state.builds = [...state.builds, ...importedBuilds];
      state.weapons = [...state.weapons, ...importedWeapons];
      persistBuilds();
      persistWeapons();
      renderAll();
      const migrationNote = importSource === "legacy" ? " Legacy v1 data was converted." : "";
      const warningNote = importWarnings.length
        ? ` Warning: these fields could not be imported: ${importWarnings.join(", ")}.`
        : "";
      feedback.textContent = `Imported ${importedBuilds.length} builds and ${importedWeapons.length} weapons.${migrationNote}${warningNote}`;
    } catch {
      feedback.textContent = "Import failed. Check that the pasted string is a valid export.";
    }
  });
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

  openModal({
    title: "Build and Weapon Comparison",
    content: `
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
  `,
  });

  els.riftModalContent.querySelectorAll("[data-matrix-build-id]").forEach((button) => {
    button.addEventListener("click", () => {
      state.matrixComparison.referenceBuildId = button.dataset.matrixBuildId;
      state.matrixComparison.referenceWeaponId = button.dataset.matrixWeaponId;
      renderBuildWeaponComparison();
    });
  });

}

function setModalMode(mode = null) {
  const modalWindow = els.riftModal.querySelector(".modal-window");
  if (!modalWindow) {
    return;
  }

  modalWindow.classList.remove(...MODAL_MODE_CLASSES);
  if (mode) {
    modalWindow.classList.add(`modal-window-${mode}`);
  }
}

function openModal({ title, content, mode = null }) {
  els.modalTitle.textContent = title;
  setModalMode(mode);
  els.riftModalContent.innerHTML = content;
  els.riftModal.classList.remove("hidden");
}

function closeModal() {
  state.matrixComparison = null;
  state.uptimeDraft = null;
  setModalMode();
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

function handleLibraryAction(action, id) {
  LIBRARY_ACTION_HANDLERS[action]?.(id);
}

function updateCompareSelection(collectionName, id, compareEnabled) {
  if (collectionName === "builds") {
    state.builds = state.builds.map((build) =>
      build.id === id ? { ...build, compareEnabled } : build,
    );
    persistBuilds();
    renderBuildList();
    return;
  }

  if (collectionName === "weapons") {
    state.weapons = state.weapons.map((weapon) =>
      weapon.id === id ? { ...weapon, compareEnabled } : weapon,
    );
    persistWeapons();
    renderWeaponList();
  }
}

function handleCompareToggle(checkbox) {
  const listId = checkbox.closest("[id]")?.id;
  const collectionName = COMPARE_COLLECTION_BY_LIST_ID[listId];
  if (!collectionName) {
    return;
  }

  updateCompareSelection(collectionName, checkbox.dataset.id, checkbox.checked);
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

  els.selectAllBuilds.addEventListener("click", () => {
    setAllBuildsCompareEnabled(true);
  });

  els.deselectAllBuilds.addEventListener("click", () => {
    setAllBuildsCompareEnabled(false);
  });

  els.newWeapon.addEventListener("click", () => {
    state.editingWeaponId = null;
    state.weaponDraft = buildDefaultWeapon();
    renderWeaponForm();
    renderWeaponList();
    scrollEditorIntoView(els.weaponForm);
  });

  els.selectAllWeapons.addEventListener("click", () => {
    setAllWeaponsCompareEnabled(true);
  });

  els.deselectAllWeapons.addEventListener("click", () => {
    setAllWeaponsCompareEnabled(false);
  });

  els.selectWeaponType.addEventListener("change", (event) => {
    const weaponType = event.target.value;
    if (!weaponType) {
      return;
    }

    selectWeaponsByType(weaponType);
    event.target.value = "";
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
      handleLibraryAction(button.dataset.action, button.dataset.id);
      return;
    }

    const checkbox = event.target.closest('input[data-action="toggle-compare"]');
    if (!checkbox) {
      return;
    }

    handleCompareToggle(checkbox);
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
    state.data = deepClone(window.WORKBOOK_DATA);
  } else {
    const response = await fetch("./data/workbook-data.json");
    state.data = await response.json();
  }

  if (!state.data) {
    throw new Error("Workbook data failed to load.");
  }

  const extensionSystem = window.PHASK_SKILL_EXTENSIONS;
  state.extensionStatus = extensionSystem
    ? extensionSystem.initialize(state.data)
    : {
        enabled: false,
        error: "Phask's skill-extension module failed to load.",
      };
  if (extensionSystem) {
    extensionSystem.installSchema(state.data);
  }
  if (!state.extensionStatus.enabled) {
    console.warn(state.extensionStatus.error);
  }

  state.uptimeFields = collectDefaultUptimeFields();
  state.builds = loadBuildsForCurrentVersion();
  state.weapons = loadWeaponsForCurrentVersion();
  state.uptimeValues = {
    ...buildDefaultUptimeValues(),
    ...loadUptimesForCurrentVersion(),
  };
  persistUptimes();

  if (!state.builds.length) {
    state.builds = [buildDefaultBuild()];
    persistBuilds();
  }
  if (!state.weapons.length) {
    state.weapons = [buildDefaultWeapon()];
    persistWeapons();
  }

  state.builds = state.builds.map(normalizeRuntimeBuild);
  persistBuilds();

  state.weapons = state.weapons.map(normalizeRuntimeWeapon);
  persistWeapons();

  const storedBuildId = loadFirstStoredValue([
    STORAGE_KEYS.selectedBuildId,
    PREVIOUS_STORAGE_KEYS.selectedBuildId,
    LEGACY_STORAGE_KEYS.selectedBuildId,
  ]);
  const storedWeaponId = loadFirstStoredValue([
    STORAGE_KEYS.selectedWeaponId,
    PREVIOUS_STORAGE_KEYS.selectedWeaponId,
    LEGACY_STORAGE_KEYS.selectedWeaponId,
  ]);
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
