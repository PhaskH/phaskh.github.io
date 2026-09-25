(function registerPhaskSkillExtensions(global) {
  "use strict";

  // APK-derived local calculator extensions. Before updating the Krea workbook,
  // read ../resources/developer-docs/calculation-extensions.md.

  const EXTENSION_SHEET = "PhaskExtensions";
  const SUPPORTED_SHEET_VERSION = "3.6.4";

  const BUILD_REFS = Object.freeze({
    meditation: "PX_MEDITATION",
    velkhanaAegis: "PX_VELKHANA_AEGIS",
    blastExploit: "PX_BLAST_EXPLOIT",
  });

  const UPTIME_REFS = Object.freeze({
    meditation: "PX_MEDITATION_UPTIME",
    blastExploitExpectedProcs: "PX_BLAST_EXPLOIT_PROCS",
  });

  const TARGET_CELLS = Object.freeze({
    velkhanaAegis: "B2",
    meditation: "B3",
    blastExploit: "B4",
  });

  const MEDITATION_BONUSES = Object.freeze([0, 0.1, 0.15, 0.2, 0.25, 0.35]);
  const VELKHANA_AEGIS_BONUSES = Object.freeze([0, 0.1, 0.15, 0.2]);
  const BLAST_EXPLOIT_ATTACK_PER_STACK = Object.freeze([0, 30, 60, 90, 120, 150]);

  const BUILD_FIELDS = Object.freeze([
    {
      ref: BUILD_REFS.blastExploit,
      key: "blastExploit",
      label: "Blast Exploit",
      options: ["0", "1", "2", "3", "4", "5"],
      defaultValue: 0,
      extension: true,
    },
    {
      ref: BUILD_REFS.meditation,
      key: "meditation",
      label: "Meditation",
      options: ["0", "1", "2", "3", "4", "5"],
      defaultValue: 0,
      extension: true,
    },
    {
      ref: BUILD_REFS.velkhanaAegis,
      key: "velkhanaAegis",
      label: "Velkhana Aegis",
      options: ["0", "1", "2", "3"],
      defaultValue: 0,
      extension: true,
    },
  ]);

  const UPTIME_FIELDS = Object.freeze([
    {
      ref: UPTIME_REFS.meditation,
      key: "meditationUptime",
      label: "Meditation",
      defaultValue: 0.95,
      displayScale: 100,
      minValue: 0,
      maxValue: 100,
      step: 0.1,
      defaultFeedback: "Using default",
      extension: true,
    },
    {
      ref: UPTIME_REFS.blastExploitExpectedProcs,
      key: "blastExploitExpectedProcs",
      label: "Blast Exploit - Procs",
      defaultValue: 0,
      displayScale: 1,
      minValue: 0,
      maxValue: 20,
      step: 1,
      defaultFeedback: "Using default",
      description:
        "Expected number of blast explosions during a 75-second hunt. Procs are assumed to occur evenly throughout the hunt. Blast Exploit gains up to 10 attack-increase stacks; their contribution is averaged based on how long each stack is active.",
      extension: true,
    },
  ]);

  const ADAPTERS = Object.freeze({
    "3.6.4": Object.freeze({
      patches: Object.freeze([
        Object.freeze({
          sheet: "Backyard",
          row: 1,
          column: 54,
          label: "flat raw attack",
          original: "=(($B$2+$S$5)*(1+$P$14)+$S$18)*(1+$P$18)",
          extended:
            "=(($B$2+$S$5)*(1+$P$14)+$S$18+PhaskExtensions!$B$4)*(1+$P$18)",
        }),
        Object.freeze({
          sheet: "Backyard",
          row: 1,
          column: 57,
          label: "final element multiplier",
          original:
            '=if($B$5="Element",((($B$3+$Y$6)*(1+$V$10)+$Y$13)*(1+$V$17)*$B$10),0)',
          extended:
            '=if($B$5="Element",((($B$3+$Y$6)*(1+$V$10)+$Y$13)*(1+$V$17+PhaskExtensions!$B$2)*$B$10),0)',
        }),
        Object.freeze({
          sheet: "Backyard",
          row: 2,
          column: 57,
          label: "critical element multiplier",
          original:
            '=if($B$5="Element",((($B$3+$Y$6)*(1+$V$10+$V$2)+$Y$13)*(1+$V$17)*$B$10),0)',
          extended:
            '=if($B$5="Element",((($B$3+$Y$6)*(1+$V$10+$V$2)+$Y$13)*(1+$V$17+PhaskExtensions!$B$2)*$B$10),0)',
        }),
        Object.freeze({
          sheet: "Backyard",
          row: 25,
          column: 36,
          label: "additive damage multiplier",
          original:
            "=$AK$2+$AK$3+$AK$4+$AK$5+$AK$6+$AK$7+$AK$8+$AK$9+$AK$10+$AK$11+$AK$12+$AK$13+$AK$14+$AK$15+$AK$16+$AK$17+$AK$18+$AK$19+$AK$20+$AK$21+$AK$22+$AK$23+$AK$24+AK27",
          extended:
            "=$AK$2+$AK$3+$AK$4+$AK$5+$AK$6+$AK$7+$AK$8+$AK$9+$AK$10+$AK$11+$AK$12+$AK$13+$AK$14+$AK$15+$AK$16+$AK$17+$AK$18+$AK$19+$AK$20+$AK$21+$AK$22+$AK$23+$AK$24+AK27+PhaskExtensions!$B$3",
        }),
      ]),
    }),
  });

  function copyFields(fields) {
    return fields.map((field) => ({
      ...field,
      ...(field.options ? { options: [...field.options] } : {}),
    }));
  }

  function validateWorkbook(data) {
    const errors = [];
    const version = String(data?.sheetVersion ?? "unknown");
    const adapter = ADAPTERS[version];

    if (!adapter) {
      errors.push(`KreaTV1 sheet ${version} is not supported by Phask's skill extensions.`);
      return errors;
    }

    for (const patch of adapter.patches) {
      const sheet = data?.sheets?.[patch.sheet];
      if (!sheet) {
        errors.push(`Missing expected sheet ${patch.sheet}.`);
        continue;
      }
      const currentFormula = sheet[patch.row]?.[patch.column];
      if (currentFormula !== patch.original) {
        errors.push(`The ${patch.label} formula changed in ${patch.sheet}.`);
      }
    }

    return errors;
  }

  function initialize(data) {
    const errors = validateWorkbook(data);
    return {
      enabled: errors.length === 0,
      error: errors.join(" "),
      sheetVersion: String(data?.sheetVersion ?? "unknown"),
    };
  }

  function installSchema(data) {
    const existingBuildKeys = new Set((data.buildFields ?? []).map((field) => field.key));
    const existingUptimeKeys = new Set((data.uptimeFields ?? []).map((field) => field.key));

    data.buildFields = [
      ...(data.buildFields ?? []),
      ...copyFields(BUILD_FIELDS).filter((field) => !existingBuildKeys.has(field.key)),
    ];
    data.uptimeFields = [
      ...(data.uptimeFields ?? []),
      ...copyFields(UPTIME_FIELDS).filter((field) => !existingUptimeKeys.has(field.key)),
    ];
  }

  function prepareSheets(sheets, sheetVersion) {
    const adapter = ADAPTERS[String(sheetVersion)];
    if (!adapter) {
      throw new Error(`Unsupported KreaTV1 sheet version: ${sheetVersion}`);
    }

    for (const patch of adapter.patches) {
      const sheet = sheets[patch.sheet];
      const currentFormula = sheet?.[patch.row]?.[patch.column];
      if (currentFormula !== patch.original) {
        throw new Error(`Cannot apply ${patch.label} extension: workbook formula changed.`);
      }
      sheet[patch.row][patch.column] = patch.extended;
    }

    sheets[EXTENSION_SHEET] = [
      ["Phask's skill extensions", "Value"],
      ["Velkhana Aegis", 0],
      ["Meditation", 0],
      ["Blast Exploit", 0],
    ];

    return sheets;
  }

  function boundedLevel(value, maximum) {
    const numeric = Math.trunc(Number(value));
    return Number.isFinite(numeric) ? Math.min(maximum, Math.max(0, numeric)) : 0;
  }

  function boundedNumber(value, minimum, maximum) {
    const numeric = Number(value);
    return Number.isFinite(numeric) ? Math.min(maximum, Math.max(minimum, numeric)) : minimum;
  }

  function averageBlastExploitStacks(expectedProcs) {
    const procs = Math.round(boundedNumber(expectedProcs, 0, 20));
    const stackGrantingProcs = Math.min(procs, 10);
    if (stackGrantingProcs === 0) {
      return 0;
    }
    return (
      stackGrantingProcs -
      (stackGrantingProcs * (stackGrantingProcs + 1)) / (2 * (procs + 1))
    );
  }

  function calculateScenarioModifiers(buildValues, weaponValues, uptimeValues) {
    const meditationLevel = boundedLevel(buildValues?.[BUILD_REFS.meditation], 5);
    const velkhanaAegisLevel = boundedLevel(buildValues?.[BUILD_REFS.velkhanaAegis], 3);
    const blastExploitLevel = boundedLevel(buildValues?.[BUILD_REFS.blastExploit], 5);
    const meditationUptime = boundedNumber(uptimeValues?.[UPTIME_REFS.meditation], 0, 1);
    const expectedBlastProcs = Math.round(
      boundedNumber(uptimeValues?.[UPTIME_REFS.blastExploitExpectedProcs], 0, 20),
    );
    const averageBlastStacks = averageBlastExploitStacks(expectedBlastProcs);

    return {
      velkhanaAegis:
        weaponValues?.E5 === "Ice" ? VELKHANA_AEGIS_BONUSES[velkhanaAegisLevel] : 0,
      meditation: MEDITATION_BONUSES[meditationLevel] * meditationUptime,
      blastExploit: BLAST_EXPLOIT_ATTACK_PER_STACK[blastExploitLevel] * averageBlastStacks,
      averageBlastStacks,
    };
  }

  global.PHASK_SKILL_EXTENSIONS = Object.freeze({
    extensionSheet: EXTENSION_SHEET,
    supportedSheetVersion: SUPPORTED_SHEET_VERSION,
    buildRefs: BUILD_REFS,
    uptimeRefs: UPTIME_REFS,
    targetCells: TARGET_CELLS,
    initialize,
    installSchema,
    prepareSheets,
    calculateScenarioModifiers,
    averageBlastExploitStacks,
  });
})(globalThis);
