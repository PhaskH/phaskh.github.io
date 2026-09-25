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
    morphAttackBoost: "PX_MORPH_ATTACK_BOOST",
  });

  const UPTIME_REFS = Object.freeze({
    buildupBoost: "PX_BUILDUP_BOOST_UPTIME",
    meditation: "PX_MEDITATION_UPTIME",
    morphAttackDamageShare: "PX_MORPH_ATTACK_DAMAGE_SHARE",
    blastExploitExpectedProcs: "PX_BLAST_EXPLOIT_PROCS",
  });

  const TARGET_CELLS = Object.freeze({
    velkhanaAegis: "B2",
    meditation: "B3",
    blastExploit: "B4",
    buildupBoost: "B5",
    morphAttackBoostDamage: "B6",
    morphAttackBoostAffinity: "B7",
    morphAttackDamageShare: "B8",
  });

  const MEDITATION_BONUSES = Object.freeze([0, 0.1, 0.15, 0.2, 0.25, 0.35]);
  const VELKHANA_AEGIS_BONUSES = Object.freeze([0, 0.1, 0.15, 0.2]);
  const BLAST_EXPLOIT_ATTACK_PER_STACK = Object.freeze([0, 30, 60, 90, 120, 150]);
  const MORPH_ATTACK_BOOST_DAMAGE = Object.freeze([0, 0.3, 0.5, 0.8]);
  const MORPH_ATTACK_BOOST_AFFINITY = Object.freeze([0, 0.6, 0.7, 0.8]);
  const MORPH_ATTACK_WEAPON_TYPES = new Set([
    "Switch Axe",
    "Charge Blade (Power)",
    "Charge Blade (Impact)",
  ]);
  const BLAST_THRESHOLD_GROWTH = 1.3;

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
      ref: BUILD_REFS.morphAttackBoost,
      key: "morphAttackBoost",
      label: "Morph Attack Boost",
      options: ["0", "1", "2", "3"],
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
      ref: UPTIME_REFS.buildupBoost,
      key: "buildupBoostUptime",
      label: "Buildup Boost",
      defaultValue: 1,
      displayScale: 100,
      minValue: 0,
      maxValue: 100,
      step: 0.1,
      defaultFeedback: "Using Phask default",
      extension: true,
    },
    {
      ref: UPTIME_REFS.meditation,
      key: "meditationUptime",
      label: "Meditation",
      defaultValue: 0.95,
      displayScale: 100,
      minValue: 0,
      maxValue: 100,
      step: 0.1,
      defaultFeedback: "Using Phask default",
      extension: true,
    },
    {
      ref: UPTIME_REFS.morphAttackDamageShare,
      key: "morphAttackDamageShare",
      label: "Morph Attack Dmg Share",
      defaultValue: 0.25,
      displayScale: 100,
      minValue: 0,
      maxValue: 100,
      step: 0.1,
      defaultFeedback: "Using Phask default",
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
      defaultFeedback: "Using Phask default",
      description:
        "Expected blast explosions during a 75-second hunt. Average stack uptime accounts for Blast thresholds increasing by 30% after each proc and assumes the hunt ends midway toward the next proc.",
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
          row: 4,
          column: 36,
          label: "Buildup Boost uptime",
          original:
            "=index(Skills!$AH$3:$AH$8,match(Calculator!$B$13,Skills!$AG$3:$AG$8,0))/100*$BI$6",
          extended:
            "=index(Skills!$AH$3:$AH$8,match(Calculator!$B$13,Skills!$AG$3:$AG$8,0))/100*$BI$6*PhaskExtensions!$B$5",
        }),
        Object.freeze({
          sheet: "Backyard",
          row: 25,
          column: 36,
          label: "additive damage multiplier",
          original:
            "=$AK$2+$AK$3+$AK$4+$AK$5+$AK$6+$AK$7+$AK$8+$AK$9+$AK$10+$AK$11+$AK$12+$AK$13+$AK$14+$AK$15+$AK$16+$AK$17+$AK$18+$AK$19+$AK$20+$AK$21+$AK$22+$AK$23+$AK$24+AK27",
          extended:
            "=$AK$2+$AK$3+$AK$4+$AK$5+$AK$6+$AK$7+$AK$8+$AK$9+$AK$10+$AK$11+$AK$12+$AK$13+$AK$14+$AK$15+$AK$16+$AK$17+$AK$18+$AK$19+$AK$20+$AK$21+$AK$22+$AK$23+$AK$24+AK27+PhaskExtensions!$B$3+PhaskExtensions!$B$6",
        }),
        Object.freeze({
          sheet: "Backyard",
          row: 11,
          column: 40,
          label: "total affinity",
          original: "=max(-1,min(1,$AO$5+$AO$7+$AO$8+AO9+AO10+AO11))",
          extended:
            "=max(-1,min(1,$AO$5+$AO$7+$AO$8+AO9+AO10+AO11+PhaskExtensions!$B$7))",
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
      ["Buildup Boost", 1],
      ["Morph Attack Boost damage", 0],
      ["Morph Attack Boost affinity", 0],
      ["Morph Attack damage share", 0.25],
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

    const cumulativeEffort = (procCount) =>
      (BLAST_THRESHOLD_GROWTH ** procCount - 1) / (BLAST_THRESHOLD_GROWTH - 1);
    const estimatedFightEffort = (cumulativeEffort(procs) + cumulativeEffort(procs + 1)) / 2;
    let stackActivationEffort = 0;
    for (let proc = 1; proc <= stackGrantingProcs; proc += 1) {
      stackActivationEffort += cumulativeEffort(proc);
    }

    return stackGrantingProcs - stackActivationEffort / estimatedFightEffort;
  }

  function calculateScenarioModifiers(buildValues, weaponValues, uptimeValues) {
    const meditationLevel = boundedLevel(buildValues?.[BUILD_REFS.meditation], 5);
    const velkhanaAegisLevel = boundedLevel(buildValues?.[BUILD_REFS.velkhanaAegis], 3);
    const blastExploitLevel = boundedLevel(buildValues?.[BUILD_REFS.blastExploit], 5);
    const morphAttackBoostLevel = boundedLevel(
      buildValues?.[BUILD_REFS.morphAttackBoost],
      3,
    );
    const meditationUptime = boundedNumber(uptimeValues?.[UPTIME_REFS.meditation], 0, 1);
    const buildupBoostUptime = boundedNumber(
      uptimeValues?.[UPTIME_REFS.buildupBoost] ?? 1,
      0,
      1,
    );
    const expectedBlastProcs = Math.round(
      boundedNumber(uptimeValues?.[UPTIME_REFS.blastExploitExpectedProcs], 0, 20),
    );
    const averageBlastStacks = averageBlastExploitStacks(expectedBlastProcs);
    const supportsMorphAttacks = MORPH_ATTACK_WEAPON_TYPES.has(weaponValues?.E7);
    const morphAttackDamageShare = supportsMorphAttacks
      ? boundedNumber(uptimeValues?.[UPTIME_REFS.morphAttackDamageShare] ?? 0.25, 0, 1)
      : 0;

    return {
      velkhanaAegis:
        weaponValues?.E5 === "Ice" ? VELKHANA_AEGIS_BONUSES[velkhanaAegisLevel] : 0,
      meditation: MEDITATION_BONUSES[meditationLevel] * meditationUptime,
      blastExploit: BLAST_EXPLOIT_ATTACK_PER_STACK[blastExploitLevel] * averageBlastStacks,
      buildupBoost: buildupBoostUptime,
      morphAttackBoostDamage: supportsMorphAttacks
        ? MORPH_ATTACK_BOOST_DAMAGE[morphAttackBoostLevel]
        : 0,
      morphAttackBoostAffinity: supportsMorphAttacks
        ? MORPH_ATTACK_BOOST_AFFINITY[morphAttackBoostLevel]
        : 0,
      morphAttackDamageShare,
      averageBlastStacks,
    };
  }

  function blendMorphAttackDamage(nonMorphResult, fullMorphResult, damageShare) {
    if (typeof nonMorphResult !== "number" || typeof fullMorphResult !== "number") {
      return fullMorphResult;
    }
    const boundedShare = boundedNumber(damageShare, 0, 1);
    return nonMorphResult + boundedShare * (fullMorphResult - nonMorphResult);
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
    blendMorphAttackDamage,
    averageBlastExploitStacks,
  });
})(globalThis);
