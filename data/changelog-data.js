// Static snapshot of the GitHub main-branch history. Add newer entries at the top.
// Dates use America/Phoenix calendar dates; commit links are built from each full hash.
window.CHANGELOG_DATA = Object.freeze([
  {
    date: "October 2, 2026",
    commit: "dd383897ec76825f49aff621174b9346c8e2b5be",
    summary: "Changed Armor Library weapon selection so weapons are saved with builds only when needed, without being added immediately.",
  },
  {
    date: "October 2, 2026",
    commit: "6b30730046b3874871b12a33c7999116184e0ac8",
    summary: "Widened the Builds and Weapons panels to make better use of horizontal space.",
  },
  {
    date: "October 2, 2026",
    commit: "a9077f2d4ff39e676af2cccf39d29b7fd9639da0",
    summary: "Added the fan-project disclaimer footer and the linked, scrollable Changelog dialog.",
  },
  {
    date: "October 1, 2026",
    commit: "e9d51a29061cee3e6f3093d4d3811ecce2398688",
    summary: "Armor Library slots now filter the available armor to the matching body part.",
  },
  {
    date: "October 1, 2026",
    commit: "36c2caebafaaeff43be83d21f217b9d0ec640101",
    summary: "Added the Armor Library and visual build editor with armor, driftstone, weapon-skill, save, search, and skill-pinning support.",
  },
  {
    date: "October 1, 2026",
    commit: "56d91e294389e7f284613b731d718a8d185d26d2",
    summary: "Restored quick Edit buttons to the Builds and Weapons sections.",
  },
  {
    date: "October 1, 2026",
    commit: "67fee37015d2d82c0afb6b4654ad91244b033434",
    summary: "Added Critical Range Boost and Charge Stock calculations with weapon-specific uptime controls.",
  },
  {
    date: "September 30, 2026",
    commit: "80507cf755484dadf2c4bb1396b2a627e25d45d1",
    summary: "Added experimental Insect Glaive support and all 54 Insect Glaive weapons.",
  },
  {
    date: "September 30, 2026",
    commit: "4c5c1057ca078b5af228525df49a2acfed0f6115",
    summary: "Added expected Retaliation and Reflection damage calculations with a configurable knockback value.",
  },
  {
    date: "September 30, 2026",
    commit: "b9780c4ed57151b0c3ab85633875ed098677d7cd",
    summary: "Added Crit-Capable Damage Share to scale affinity for attacks that cannot critically hit.",
  },
  {
    date: "September 29, 2026",
    commit: "6c9d2a4b05182daa5d43c9d33a800d7392f5f0f4",
    summary: "Expanded rift combination comparisons to multiple builds with selectable ranking.",
  },
  {
    date: "September 29, 2026",
    commit: "035d6932a14a0dcce546724801d3b9095b87557e",
    summary: "Added mobile-friendly Build and Weapon management dialogs with reordering, editing, deletion, and undo.",
  },
  {
    date: "September 25, 2026",
    commit: "85212816021736f798eb1d2cb050d451b5a7c519",
    summary: "Added named uptime presets, quick preset switching, default-state detection, and backup support.",
  },
  {
    date: "September 25, 2026",
    commit: "59e072e3e05aca3c759ecbace37b9d6b609d0059",
    summary: "Set the default Buildup Boost uptime to 33.3% and prevented it from affecting non-status weapons.",
  },
  {
    date: "September 25, 2026",
    commit: "312b8e4be586dbe077dc31b740fadccd3bca50fc",
    summary: "Added Day Mode for experimental uptime controls and revised how Buildup Boost uptime is applied.",
  },
  {
    date: "September 25, 2026",
    commit: "e67553bd386098e70468fcdd204f024aa6bb5967",
    summary: "Added Morph Attack Boost and a morph-attack damage-share control for Switch Axe and Charge Blade.",
  },
  {
    date: "September 25, 2026",
    commit: "5532b625d86f2577d5cf80def94f6d152b84e48a",
    summary: "Corrected weapon stats, names, skills, terminology, and several weapon mechanics against official data.",
  },
  {
    date: "September 25, 2026",
    commit: "bca794111eb7322514ca50193439088d1071884d",
    summary: "Added an uptime control for Buildup Boost calculations.",
  },
  {
    date: "September 25, 2026",
    commit: "7065a1c3b971d266434e93de35108f4b77ea475e",
    summary: "Improved Blast Exploit proc estimates with cumulative Blast effort and refined alphabetical sorting.",
  },
  {
    date: "September 24, 2026",
    commit: "d08663e896e27a7ed2bd2bd1c1eff55fa9cd904d",
    summary: "Added Velkhana Aegis, Meditation, and Blast Exploit calculations and modularized skill extensions.",
  },
  {
    date: "September 24, 2026",
    commit: "d55e4148d6cff4c82b2380086c484de71f7979f7",
    summary: "Added rift information for the new Barroth, Pukei-Pukei, and Rathian weapons.",
  },
  {
    date: "September 24, 2026",
    commit: "b5c24e80e3fab1899ae818a746e1be76a164d0b5",
    summary: "Added Brachydios, Aurora Somnacanth, 26 weapons, more rift options, and a smaller weapon dataset.",
  },
  {
    date: "September 12, 2026",
    commit: "c0dc3f690f365687337c34423af20a99d0a639f8",
    summary: "Added status editing to the Weapon Editor and monster names to Weapon Library entries.",
  },
  {
    date: "September 11, 2026",
    commit: "8f4047bfb63409d5251742488fcc356063a9d849",
    summary: "Added the Weapon Library for browsing and quickly importing existing weapons.",
  },
  {
    date: "August 28, 2026",
    commit: "d93b0a63b5ad8f6cf4c02d519cc1ee1599b1c842",
    summary: "Enabled saving rift weapon variants directly from the dialog and removed the Beta label.",
  },
  {
    date: "August 27, 2026",
    commit: "7694256c96bf75d7e919363e7fa5dbb27a565ffb",
    summary: "Refactored the application code for maintainability without changing behavior.",
  },
  {
    date: "August 27, 2026",
    commit: "1197ea42b15a7d63b503ec65847d8d0e870609b5",
    summary: "Updated calculations to Krea v3.6.4, added Elemental Release, migrated storage keys, and improved legacy compatibility.",
  },
  {
    date: "May 20, 2026",
    commit: "8b25f2175c1ba73839b3d3921ea1349b28191ad3",
    summary: "Updated calculations to Krea v3.6.3 with Combo Master, build ordering, and weapon-selection shortcuts.",
  },
  {
    date: "May 20, 2026",
    commit: "a47127c492d08af94c8811c2e0fc21bc1fee83d4",
    summary: "Added the website icon.",
  },
  {
    date: "May 20, 2026",
    commit: "0ea32a728c19e1465fed558bfab47e02e571afde",
    summary: "Expanded the Remaining Health input range from 1–100 to 1–160.",
  },
  {
    date: "April 8, 2026",
    commit: "b18d9b669b2b5babc45c38b9fede92e9aaed2e2d",
    summary: "Added select-all and deselect-all controls plus clearer selected-export details.",
  },
  {
    date: "April 8, 2026",
    commit: "e9b5e7a0e1a398f57eab4a5509cd7e5be63d71b8",
    summary: "Added default uptime indicators and Resuscitate, Coalescence, Bleeding Edge, and Bubbly Dance uptime inputs.",
  },
  {
    date: "April 8, 2026",
    commit: "5d529779e36d4421fdf7ab6faaa99db4e743a175",
    summary: "Added Remaining Health to the uptime dialog for Vital Element calculations.",
  },
  {
    date: "April 5, 2026",
    commit: "c6e8e1a0eb9571d106d7f615bea2e7c1fa3cac47",
    summary: "Adjusted the desktop comparison table to prevent build-name truncation.",
  },
  {
    date: "April 5, 2026",
    commit: "7320dc256cf75490464a833a990b4ecaab5ba2fc",
    summary: "Further reduced comparison-table spacing and cell sizes on mobile devices.",
  },
  {
    date: "April 5, 2026",
    commit: "557eb169388ce84725537f1ab21dec7624923d7c",
    summary: "Optimized the comparison table for small screens and mobile devices.",
  },
  {
    date: "April 5, 2026",
    commit: "b066a7a6246485943bcfcba610c300327d912912",
    summary: "Replaced Vital Fire with a generic Vital Element skill that follows the selected weapon element.",
  },
  {
    date: "April 5, 2026",
    commit: "a4e6b2f22e619f625202194cd0327dbbefcb70aa",
    summary: "Updated the site introduction to recognize the lance community.",
  },
  {
    date: "April 5, 2026",
    commit: "256f61d6c909e9b0e10eef4123062bb2cdf081bf",
    summary: "Improved the comparison-table layout for mobile screens.",
  },
  {
    date: "April 5, 2026",
    commit: "466a204732f5f3584a0cc68726e034653e9e9562",
    summary: "Linked the KreaTV1 sheet and standardized the tool's eDPS and Effective Damage labels.",
  },
  {
    date: "April 5, 2026",
    commit: "a0eadfa7d8fb0f4774dd9498bf6d161caace0222",
    summary: "Updated to Krea v3.6.2, added selected exports, refreshed branding, and improved mobile layouts.",
  },
  {
    date: "April 4, 2026",
    commit: "b1f312305673b622c6bbc45950068288b687fdd7",
    summary: "Allowed weapons to use negative affinity values.",
  },
  {
    date: "April 4, 2026",
    commit: "3d6b9c5603ecda907fd26ae5db25725a5d81180a",
    summary: "Polished and repositioned the uptime dialog and its controls.",
  },
  {
    date: "April 4, 2026",
    commit: "d67f34a2b55b12a3a55ff9c7eb5208e933574570",
    summary: "Added a persistent uptime dialog with editable values, defaults, save, and discard behavior.",
  },
  {
    date: "April 4, 2026",
    commit: "7b08e49c68c07549a7a9ebe8d7489e974ba754be",
    summary: "Fixed Android keyboard resizing so Build Name editing retains focus.",
  },
  {
    date: "April 3, 2026",
    commit: "c95f81ae484142c09df6f650f2ab3d941c6a653d",
    summary: "Polished comparison controls, rift feedback, helper text, and calculator naming.",
  },
  {
    date: "April 3, 2026",
    commit: "897ca32f6ce90e3152771d3a4eaac439e8ad6942",
    summary: "Added selective comparisons, skill highlighting, weapon-value shortcuts, and safer rift variant creation.",
  },
  {
    date: "April 3, 2026",
    commit: "86c2c7bfa8e7162dde5f10fa1129d6ead35cbfe0",
    summary: "Created the initial MHN build comparison calculator with local build and weapon management.",
  },
]);
