/*******************************************************
 * Tool      : GDrive-to-GSheet
 * Version   : 1.0.0
 * Date      : 2025-11-30
 * Author    : techpoov+GDrive-to-GSheet@gmail.com
 *
 * TechPoov Menu → Scan GDrive (Resume + Triggers)
 * Bound Apps Script — single file, no external libs
 * -----------------------------------------------------
 * PROFILES INPUT (sheet "Profiles")
 *
 *   *ProfileName | *FolderId | *Scan | Mode | MaxDepth | ShowURL | LastUpdated | Remarks
 *
 *   *ProfileName : Logical name (must be UNIQUE across all rows)
 *   *FolderId    : Google Drive Folder ID (from URL)
 *   *Scan        : "Yes" (case-insensitive) to include in next run
 *   Mode         : Files / Folders / Both (blank = Both)
 *   MaxDepth     : 0 or blank = all levels,
 *                  1 = only that folder,
 *                  2 = that folder + one level down, etc.
 *   ShowURL      : "Yes" = Add URL column in output.
 *                  Blank/"No" = No URL column (only hyperlinks).
 *   LastUpdated  : Set by tool when scan completes
 *   Remarks      : Status / progress, set by tool
 *
 * OUTPUT PER PROFILE (per Mode)
 *   Sheet name: <ProfileName>_<ModeSuffix>
 *
 *   Common first columns (all modes):
 *     SlNo | ID | ...
 *
 *   Mode = Files:
 *     SlNo | ID | Path | File | Extension | DtCreated | DtModified [| URL if ShowURL=Yes]
 *
 *   Mode = Both:
 *     SlNo | ID | Path | Name | Type | Extension | DtCreated | DtModified [| URL if ShowURL=Yes]
 *
 *   Mode = Folders:
 *     SlNo | ID | Path | File_Count | Subfolder_Count | DtCreated | DtModified [| URL if ShowURL=Yes]
 *
 * RESUME / CHECKPOINT
 *   DocumentProperties key: SCAN_STATE::<ProfileName>
 *   {
 *     profileName, folderId, mode, maxDepth,
 *     topName,
 *     queue:[{id, path, depth}],
 *     processedCount, startedAtISO, phase,
 *     sheetName, oldSheetName,
 *     lastRun, writeRow, pendingRows:[],
 *     lastProgressNoteAt,
 *     showURL,
 *     nextSlNo
 *   }
 *
 * SINGLE CONTINUATION TRIGGER
 *   DocumentProperties key: SCAN_WORKER_TRIGGER_ID
 *
 * LOGGING
 *   Sheet "Log": Timestamp | Level | Profile | Phase | Message | Details
 *   Profile column = sheetName when available (e.g. MyProfile_Both)
 *******************************************************/

const TOOL_NAME         = 'GDrive-to-GSheet';
const TOOL_VERSION      = '1.0.0';
const TOOL_VERSION_DATE = '2025-11-30';

const STATE_PREFIX      = 'SCAN_STATE::';
const TRIGGER_KEY       = 'SCAN_WORKER_TRIGGER_ID';
const ACTIVE_SET        = 'SCAN_ACTIVE_PROFILES'; // JSON array of profileNames

// Tracks when a worker slice started (used by timeAlmostUp)
let SLICE_START_TS = 0;

// Worker timing
const MAX_RUNTIME_MS       = 4 * 60 * 1000;  // 4 minutes total budget per slice
const STOP_MARGIN_MS       = 30 * 1000;
const NEXT_RUN_DELAY_MS    = 15 * 1000;      // 15s between slices
const PROGRESS_INTERVAL_MS = 60 * 1000;      // update Remarks roughly every 60s

// Column width configuration (pixels)
const COL_WIDTH_SLNO   = 60;
const COL_WIDTH_ID     = 180; // usually hidden, but keep sane width
const COL_WIDTH_PATH   = 300;
const COL_WIDTH_NAME   = 220;
const COL_WIDTH_TYPE   = 90;
const COL_WIDTH_EXT    = 90;
const COL_WIDTH_COUNT  = 120;
const COL_WIDTH_DATE   = 150;
const COL_WIDTH_URL    = 320;

// Columns (by header name) that should be auto-resized
const AUTO_RESIZE_HEADERS = ['SlNo', 'Extension', 'DtCreated', 'DtModified'];

/*******************************************************
 * ENTRYPOINTS
 *******************************************************/

function onOpen() {
  runSafely('onOpen', () => {
    ensureProfilesStructure();
    buildMenu();
  });
}

function menuScanGdriveFolders() {
  runSafely('menuScanGdriveFolders', () => {
    const ok = validateProfiles();
    if (!ok) return;
    mainOrchestrator('SCAN_PROFILES');
  });
}

function scanWorker() {
  runSafely('scanWorker', () => {
    SLICE_START_TS = Date.now();           // mark start of this worker slice
    mainOrchestrator('WORKER_SLICE');
  });
}

function menuAboutGDriveToGSheet() {
  runSafely('menuAboutGDriveToGSheet', () => {
    showAboutToolDialog();
  });
}

/*******************************************************
 * MAIN ORCHESTRATOR
 *******************************************************/

function mainOrchestrator(action) {
  switch (action) {
    case 'SCAN_PROFILES':
      handleScanProfiles();
      break;

    case 'WORKER_SLICE':
      handleWorkerSlice();
      break;

    default:
      logSimple('', 'MAIN', 'ERROR', `Unknown action: ${action}`, '');
  }
}

/*******************************************************
 * TOP-LEVEL FLOW HANDLERS
 *******************************************************/

function handleScanProfiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profilesSheet = ss.getSheetByName('Profiles');
  if (!profilesSheet) throw new Error('Profiles sheet not found.');

  const { rows, colMap } = readTable(profilesSheet);

  const idxProfileName = getColIndex(colMap, 'ProfileName');
  const idxFolderId    = getColIndex(colMap, 'FolderId');
  const idxScan        = getColIndex(colMap, 'Scan');
  const idxMode        = getColIndex(colMap, 'Mode');
  const idxMaxDepth    = getColIndex(colMap, 'MaxDepth');
  const idxShowURL     = getColIndex(colMap, 'ShowURL');

  if (idxProfileName == null || idxFolderId == null || idxScan == null) {
    throw new Error('Profiles sheet missing mandatory columns. Please reopen the spreadsheet to allow the tool to repair headers.');
  }

  const toRun = [];

  rows.forEach(r => {
    const scanVal = String(r[idxScan] || '').trim().toLowerCase();
    if (scanVal !== 'yes') return;

    const profileName = String(r[idxProfileName] || '').trim();
    const folderId    = String(r[idxFolderId]    || '').trim();
    const modeRaw     = idxMode     != null ? r[idxMode]     : '';
    const maxDepthRaw = idxMaxDepth != null ? r[idxMaxDepth] : '';
    const showURLRaw  = idxShowURL  != null ? r[idxShowURL]  : '';

    const mode     = normalizeMode(modeRaw);     // 'files' | 'folders' | 'both'
    const maxDepth = parseMaxDepth(maxDepthRaw); // 0 or >=1
    const showURL  = String(showURLRaw || '').trim().toLowerCase() === 'yes';

    const ok = initOrResetProfileState({ profileName, folderId, mode, maxDepth, showURL });
    if (ok) {
      // Clear any old hyperlink on Profiles before starting a fresh/continued run
      clearProfileHyperlink(profileName);
      toRun.push(profileName);
    }
  });

  setActiveProfiles(toRun);
  ensureSingleTrigger(); // clean previous worker triggers

  logSimple(
    toRun.join(', '),
    'INIT',
    'INFO',
    `Queued ${toRun.length} profile(s). Starting worker immediately.`,
    ''
  );

  // First slice immediately; continuation by trigger
  handleWorkerSlice();
}

/**
 * Orchestrated worker slice:
 *  - Resolve which profile/state to work on
 *  - Run slice loop
 *  - Finalize triggers, states, and UI
 */
function handleWorkerSlice() {
  const start = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const profilesSheet = ss.getSheetByName('Profiles');

  const resolved = resolveActiveState();
  if (!resolved) {
    finalizeWhenNoActiveProfiles(ss, profilesSheet);
    return;
  }

  const profileName = resolved.profileName;
  let state = resolved.state;
  const activesAtStart = resolved.actives;

  const outSheet = ensureOutputSheetForState(ss, state);

  keepProfilesSheetVisible(ss, profilesSheet);

  state = runProfileSliceLoop(state, outSheet, start);
  setState(profileName, state);

  finalizeAfterSlice(state, activesAtStart, profilesSheet);
}

/**
 * Find the next active profile/state to work on.
 */
function resolveActiveState() {
  let actives = getActiveProfiles();
  if (!actives.length) {
    actives = discoverActiveProfilesFromState();
    setActiveProfiles(actives);
  }

  if (!actives.length) {
    return null;
  }

  const profileName = actives[0];
  const state = getState(profileName);

  if (!state || state.phase === 'DONE') {
    const rest = actives.slice(1);
    setActiveProfiles(rest);
    if (!rest.length) return null;
    return resolveActiveState();
  }

  return { profileName: profileName, state: state, actives: actives };
}

/**
 * When there are no active profiles, clean up trigger and keep Profiles visible.
 */
function finalizeWhenNoActiveProfiles(ss, profilesSheet) {
  ensureSingleTrigger(true);
  logSimple('', 'IDLE', 'INFO', 'No active profiles. Worker trigger cleared.', '');
  keepProfilesSheetVisible(ss, profilesSheet);
}

/**
 * Always keep Profiles sheet as the visible tab for the user.
 */
function keepProfilesSheetVisible(ss, profilesSheet) {
  try {
    if (profilesSheet) ss.setActiveSheet(profilesSheet);
  } catch (e) {
    logSimple('', 'UI', 'WARN', 'Could not set Profiles sheet as active.', e);
  }
}

/**
 * Core worker loop for a single profile/state within the time budget.
 */
function runProfileSliceLoop(state, outSheet, startTs) {
  while (Date.now() - startTs < (MAX_RUNTIME_MS - STOP_MARGIN_MS)) {
    if (state.phase === 'SCANNING') {
      const progressed = processScanSlice(state, outSheet, startTs);
      setState(state.profileName, state);
      if (!progressed) break;
      continue;
    }

    if (state.phase === 'MERGE_BACK') {
      const progressed = processMergeBackSlice(state);
      setState(state.profileName, state);
      if (!progressed) break;
      continue;
    }

    if (state.phase === 'DONE') {
      break;
    }
  }

  state.lastRun = Date.now();
  return state;
}

/**
 * Update active set, schedule next trigger, finalize profile if needed.
 */
function finalizeAfterSlice(state, activesAtStart, profilesSheet) {
  const profileName = state.profileName;
  let actives = getActiveProfiles();
  if (!actives.length && activesAtStart && activesAtStart.length) {
    actives = activesAtStart;
  }

  if (state.phase === 'DONE') {
    const rest = actives.filter(name => name !== profileName);
    setActiveProfiles(rest);
    finalizeProfile(state);
    logSimple(
      state.sheetName || state.profileName,
      'DONE',
      'INFO',
      'Completed. Processed ' + formatInt(state.processedCount) + ' items.',
      ''
    );
  }

  const hasMore = getActiveProfiles().length > 0 || anyScanningInProgress();
  if (hasMore) {
    ensureSingleTrigger();
    scheduleNextWorker(NEXT_RUN_DELAY_MS);
  } else {
    ensureSingleTrigger(true);
    logSimple('', 'IDLE', 'INFO', 'All work complete. Worker trigger cleared.', '');
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  keepProfilesSheetVisible(ss, profilesSheet);
}

/*******************************************************
 * MENU & UI HELPERS
 *******************************************************/

function buildMenu() {
  SpreadsheetApp.getUi()
    .createMenu('TechPoov Menu')
    .addItem('Scan GDrive', 'menuScanGdriveFolders')
    .addItem('Cancel scanning', 'cancelAllScans')
    .addSeparator()
    .addItem('About GDrive-to-GSheet', 'menuAboutGDriveToGSheet')
    .addToUi();
}

function showAboutToolDialog() {
  const htmlContent = `
    <div style="
      font-family:Segoe UI, Arial, sans-serif;
      font-size:13px;
      line-height:1.6;
      padding:10px;
    ">
      <div style="font-size:16px; font-weight:bold; margin-bottom:8px;">
        About ${TOOL_NAME}
      </div>

      <b>Tool:</b> ${TOOL_NAME}<br>
      <b>Version:</b> ${TOOL_VERSION}<br>
      <b>Release Date:</b> ${TOOL_VERSION_DATE}<br>
      <b>Website:</b>
      <a href="https://techpoov.github.io" target="_blank">
        https://techpoov.github.io
      </a><br>
      <b>Contact:</b>
      <a href="mailto:TechPoov+GDrive-to-GSheet@gmail.com">
        TechPoov+GDrive-to-GSheet@gmail.com
      </a>
    </div>
  `;

  const html = HtmlService.createHtmlOutput(htmlContent)
    .setWidth(420)
    .setHeight(240);

  SpreadsheetApp.getUi().showModalDialog(html, 'About ' + TOOL_NAME);
}

/*******************************************************
 * SAFE WRAPPER
 *******************************************************/

function runSafely(contextName, fn) {
  try {
    return fn();
  } catch (err) {
    logSimple('', contextName, 'ERROR',
              `${TOOL_NAME} v${TOOL_VERSION} - ${err.message}`, err);
    throw err;
  }
}

/*******************************************************
 * PROFILES STRUCTURE (onOpen)
 *******************************************************/

function ensureProfilesStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Profiles');
  if (!sheet) {
    sheet = ss.insertSheet('Profiles');
    sheet.getRange(1, 1, 1, 8).setValues([[
      '*ProfileName',
      '*FolderId',
      '*Scan',
      'Mode',
      'MaxDepth',
      'ShowURL',
      'LastUpdated',
      'Remarks'
    ]]);
  }

  const expected = [
    { name: 'ProfileName', mandatory: true },
    { name: 'FolderId',   mandatory: true },
    { name: 'Scan',       mandatory: true },
    { name: 'Mode',       mandatory: false },
    { name: 'MaxDepth',   mandatory: false },
    { name: 'ShowURL',    mandatory: false },
    { name: 'LastUpdated',mandatory: false },
    { name: 'Remarks',    mandatory: false }
  ];

  let hdr = getHeader(sheet);
  let colMapRaw = {};
  hdr.forEach((h, i) => { colMapRaw[h] = i; });

  function ensureColumn(baseName, mandatory) {
    let idx = getColIndex(colMapRaw, baseName);
    if (idx == null) {
      const lc = sheet.getLastColumn() || 0;
      sheet.insertColumnAfter(Math.max(lc, 1));
      const newCol = sheet.getLastColumn();
      const hdrValue = mandatory ? ('*' + baseName) : baseName;
      sheet.getRange(1, newCol).setValue(hdrValue);
      hdr = getHeader(sheet);
      colMapRaw = {};
      hdr.forEach((h, i) => { colMapRaw[h] = i; });
      return newCol - 1;
    } else {
      const currentName = hdr[idx];
      if (mandatory && currentName.charAt(0) !== '*') {
        sheet.getRange(1, idx + 1).setValue('*' + baseName);
        hdr[idx] = '*' + baseName;
      }
      return idx;
    }
  }

  expected.forEach(e => ensureColumn(e.name, e.mandatory));

  // Refresh header + colMapRaw
  hdr = getHeader(sheet);
  colMapRaw = {};
  hdr.forEach((h, i) => { colMapRaw[h] = i; });

  const idxProfileName = getColIndex(colMapRaw, 'ProfileName');
  const idxFolderId    = getColIndex(colMapRaw, 'FolderId');
  const idxScan        = getColIndex(colMapRaw, 'Scan');
  const idxMode        = getColIndex(colMapRaw, 'Mode');
  const idxMaxDepth    = getColIndex(colMapRaw, 'MaxDepth');
  const idxShowURL     = getColIndex(colMapRaw, 'ShowURL');
  const idxLastUpdated = getColIndex(colMapRaw, 'LastUpdated');
  const idxRemarks     = getColIndex(colMapRaw, 'Remarks');

  if (idxProfileName != null) {
    sheet.getRange(1, idxProfileName + 1).setNote(
      'Logical name for this scan profile. Must be unique across all rows.'
    );
  }
  if (idxFolderId != null) {
    sheet.getRange(1, idxFolderId + 1).setNote(
      'Google Drive Folder ID (from the URL). Tool starts scanning from this folder.'
    );
  }
  if (idxScan != null) {
    sheet.getRange(1, idxScan + 1).setNote(
      'Type "Yes" to include this profile in the next run. Leave blank to skip.'
    );
  }
  if (idxMode != null) {
    sheet.getRange(1, idxMode + 1).setNote(
      'Files / Folders / Both. Blank = Both.'
    );
  }
  if (idxMaxDepth != null) {
    sheet.getRange(1, idxMaxDepth + 1).setNote(
      '0 or blank = all levels. 1 = only this folder. 2 = this folder + one level down, etc.'
    );
  }
  if (idxShowURL != null) {
    sheet.getRange(1, idxShowURL + 1).setNote(
      '"Yes" = Add a raw URL column in the output sheet (useful for auditing, comparison, export).\n' +
      'Blank or "No" = Only add clickable hyperlinks, no URL column.'
    );
  }
  if (idxLastUpdated != null) {
    sheet.getRange(1, idxLastUpdated + 1).setNote(
      'Automatically set when scan completes.'
    );
  }
  if (idxRemarks != null) {
    sheet.getRange(1, idxRemarks + 1).setNote(
      'Status messages during scan (progress, queue size, etc.).'
    );
  }
}

/*******************************************************
 * PROFILES VALIDATION (before scan) — Orchestrated
 *******************************************************/

function validateProfiles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Profiles');

  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      'Profiles sheet not found. Please reopen the spreadsheet to allow the tool to create it.'
    );
    return false;
  }

  const meta = collectProfileMeta(sheet);
  if (!meta.rows.length) {
    SpreadsheetApp.getUi().alert('No profiles found in Profiles sheet.');
    return false;
  }

  const errors = []
    .concat(validateProfileUniqueness(meta))
    .concat(validateActiveProfiles(meta));

  if (!meta.activeCount) {
    errors.push('No profiles with Scan="Yes" found. Set Scan="Yes" for at least one profile to run.');
  }

  if (errors.length) {
    const msg = 'Cannot start scan due to the following issues:\n\n' + errors.join('\n');
    SpreadsheetApp.getUi().alert(msg);
    return false;
  }

  return true;
}

function collectProfileMeta(sheet) {
  const { rows, colMap, a1StartRow } = readTable(sheet);

  const idxProfileName = getColIndex(colMap, 'ProfileName');
  const idxFolderId    = getColIndex(colMap, 'FolderId');
  const idxScan        = getColIndex(colMap, 'Scan');
  const idxMode        = getColIndex(colMap, 'Mode');
  const idxMaxDepth    = getColIndex(colMap, 'MaxDepth');

  if (idxProfileName == null || idxFolderId == null || idxScan == null) {
    SpreadsheetApp.getUi().alert(
      'Profiles sheet is missing mandatory columns (*ProfileName, *FolderId, *Scan). ' +
      'Please reopen the sheet to let the tool repair headers.'
    );
    return {
      rows: [],
      colMap: colMap,
      a1StartRow: a1StartRow,
      idxProfileName: idxProfileName,
      idxFolderId: idxFolderId,
      idxScan: idxScan,
      idxMode: idxMode,
      idxMaxDepth: idxMaxDepth,
      activeCount: 0
    };
  }

  return {
    rows: rows,
    colMap: colMap,
    a1StartRow: a1StartRow,
    idxProfileName: idxProfileName,
    idxFolderId: idxFolderId,
    idxScan: idxScan,
    idxMode: idxMode,
    idxMaxDepth: idxMaxDepth,
    activeCount: 0
  };
}

function validateProfileUniqueness(meta) {
  const errors = [];
  const allNames = {};

  meta.rows.forEach((r, idx) => {
    const rowNum = meta.a1StartRow + idx;
    const profileName = String(r[meta.idxProfileName] || '').trim();
    if (!profileName) return;

    if (allNames[profileName]) {
      errors.push(
        `Duplicate ProfileName "${profileName}" in rows ${allNames[profileName]} and ${rowNum}. ` +
        'Each profile must have a unique name (e.g. MyProfile_Files, MyProfile_Both).'
      );
    } else {
      allNames[profileName] = rowNum;
    }
  });

  return errors;
}

function validateActiveProfiles(meta) {
  const errors = [];

  meta.rows.forEach((r, idx) => {
    const rowNum  = meta.a1StartRow + idx;
    const scanVal = String(r[meta.idxScan] || '').trim().toLowerCase();
    if (scanVal !== 'yes') return;

    meta.activeCount++;

    const profileName = String(r[meta.idxProfileName] || '').trim();
    const folderId    = String(r[meta.idxFolderId]    || '').trim();
    const modeRaw     = meta.idxMode     != null ? r[meta.idxMode]     : '';
    const maxDepthRaw = meta.idxMaxDepth != null ? r[meta.idxMaxDepth] : '';

    if (!profileName) {
      errors.push(`Row ${rowNum}: ProfileName is mandatory when Scan="Yes".`);
    }

    if (!folderId) {
      errors.push(`Row ${rowNum}: FolderId is mandatory when Scan="Yes".`);
    } else {
      try {
        DriveApp.getFolderById(folderId);
      } catch (_e) {
        errors.push(`Row ${rowNum}: Invalid or inaccessible FolderId for profile "${profileName}".`);
      }
    }

    const modeStr = String(modeRaw || '').trim();
    if (modeStr) {
      const normalized = normalizeMode(modeStr);
      const lower = modeStr.toLowerCase();
      const allowed = ['files', 'file', 'folders', 'folder', 'both'];
      if (allowed.indexOf(lower) === -1 ||
          (normalized !== 'files' && normalized !== 'folders' && normalized !== 'both')) {
        errors.push(
          `Row ${rowNum}: Invalid Mode "${modeStr}". ` +
          'Allowed: Files, Folders, Both (or leave blank for Both).'
        );
      }
    }

    const mdStr = String(maxDepthRaw || '').trim();
    if (mdStr) {
      const n = parseInt(mdStr, 10);
      if (isNaN(n) || n < 0) {
        errors.push(
          `Row ${rowNum}: MaxDepth must be 0 or a positive whole number (got "${mdStr}").`
        );
      }
    }
  });

  return errors;
}

/*******************************************************
 * PROFILE STATE INIT / RESET
 *******************************************************/

function initOrResetProfileState({ profileName, folderId, mode, maxDepth, showURL }) {
  let state = getState(profileName);
  const normalizedMode = normalizeMode(mode); // 'files'|'folders'|'both'
  const depthLimit = parseMaxDepth(maxDepth); // 0 or >=1
  const showURLFlag = !!showURL;

  if (state && state.phase && state.phase !== 'DONE') {
    if (!state.mode || state.mode !== normalizedMode || state.maxDepth !== depthLimit) {
      logSimple(
        profileName,
        'INIT',
        'INFO',
        `Mode/MaxDepth changed. Resetting state (old mode="${state.mode}", old maxDepth="${state.maxDepth}", ` +
        `new mode="${normalizedMode}", new maxDepth="${depthLimit}").`,
        ''
      );
      clearState(profileName);
      state = null;
    } else {
      if (typeof state.showURL === 'undefined') state.showURL = showURLFlag;
      if (typeof state.nextSlNo === 'undefined') state.nextSlNo = 1;
      updateProgressNoteIfDue(state, true);
      logSimple(
        state.sheetName || profileName,
        state.phase,
        'INFO',
        'Resuming existing profile state.',
        ''
      );
      return true;
    }
  }

  let topName = '';
  try {
    topName = DriveApp.getFolderById(folderId).getName();
  } catch (e) {
    logSimple(
      profileName,
      'INIT',
      'ERROR',
      `Invalid or inaccessible FolderId: ${folderId}`,
      e
    );
    return false;
  }

  const baseName = profileName + '_' + modeToSheetSuffix(normalizedMode);
  const { sheetName, oldSheetName } = prepareRolloverSheets(baseName, normalizedMode, showURLFlag);

  const queue = [{ id: folderId, path: topName, depth: 1 }];

  state = {
    profileName,
    folderId,
    mode: normalizedMode,
    maxDepth: depthLimit,
    topName,
    queue,
    processedCount: 0,
    startedAtISO: new Date().toISOString(),
    phase: 'SCANNING',
    sheetName,
    oldSheetName,
    lastRun: Date.now(),
    writeRow: 2,
    pendingRows: [],
    lastProgressNoteAt: 0,
    showURL: showURLFlag,
    nextSlNo: 1
  };

  setState(profileName, state);
  updateProgressNoteIfDue(state, true);
  return true;
}

/*******************************************************
 * SCAN SLICE (SCANNING PHASE)
 *******************************************************/

function processScanSlice(state, outSheet, sliceStartTs) {
  const mode = state.mode || 'both'; // 'files'|'folders'|'both'

  if (state.pendingRows && state.pendingRows.length) {
    flushRows(outSheet, state);
    updateProgressNoteIfDue(state);
    return true;
  }

  if (!state.queue.length) {
    state.phase = 'MERGE_BACK';
    updateProgressNoteIfDue(state);
    return true;
  }

  const node = state.queue.shift(); // {id, path, depth}
  let folder;
  try {
    folder = DriveApp.getFolderById(node.id);
  } catch (e) {
    logSimple(
      state.sheetName || state.profileName,
      'SCANNING',
      'ERROR',
      `Cannot access folderId ${node.id}; skipping.`,
      e
    );
    updateProgressNoteIfDue(state);
    return true;
  }

  const folderUrl     = folder.getUrl();
  const folderCreated = folder.getDateCreated ? folder.getDateCreated() : null;
  const folderUpdated = folder.getLastUpdated ? folder.getLastUpdated() : null;
  const canGoDeeper   = (state.maxDepth === 0 || node.depth < state.maxDepth);

  if (mode === 'files') {
    try {
      const fileIter = folder.getFiles();
      while (fileIter.hasNext()) {
        const f    = fileIter.next();
        const name = f.getName();
        const ext  = deriveExtension(name, f.getMimeType());
        const url  = f.getUrl();
        const dc   = f.getDateCreated ? f.getDateCreated() : null;
        const dm   = f.getLastUpdated ? f.getLastUpdated() : null;

        state.pendingRows.push({
          id: f.getId(),
          values: [
            node.path, // Path
            name,      // File
            ext,       // Extension
            dc,        // DtCreated
            dm         // DtModified
          ],
          url: url
        });

        if (state.pendingRows.length >= 1000) flushRows(outSheet, state);
        if (Date.now() - sliceStartTs > (MAX_RUNTIME_MS - STOP_MARGIN_MS)) break;
      }

      if (canGoDeeper) {
        const folderIter = folder.getFolders();
        while (folderIter.hasNext()) {
          const sub     = folderIter.next();
          const subName = sub.getName();
          state.queue.push({ id: sub.getId(), path: `${node.path}/${subName}`, depth: node.depth + 1 });
          if (Date.now() - sliceStartTs > (MAX_RUNTIME_MS - STOP_MARGIN_MS)) break;
        }
      }
    } catch (e) {
      logSimple(
        state.sheetName || state.profileName,
        'SCANNING',
        'WARN',
        'Listing children failed in Files mode; will retry.',
        e
      );
    }

  } else if (mode === 'folders') {
    let fileCount      = 0;
    let subfolderCount = 0;

    try {
      const fileIter = folder.getFiles();
      while (fileIter.hasNext()) {
        fileIter.next();
        fileCount++;
        if (Date.now() - sliceStartTs > (MAX_RUNTIME_MS - STOP_MARGIN_MS)) break;
      }

      const folderIter = folder.getFolders();
      while (folderIter.hasNext()) {
        const sub = folderIter.next();
        subfolderCount++;
        if (canGoDeeper) {
          const subName = sub.getName();
          state.queue.push({ id: sub.getId(), path: `${node.path}/${subName}`, depth: node.depth + 1 });
        }
        if (Date.now() - sliceStartTs > (MAX_RUNTIME_MS - STOP_MARGIN_MS)) break;
      }
    } catch (e) {
      logSimple(
        state.sheetName || state.profileName,
        'SCANNING',
        'WARN',
        'Listing children failed in Folders mode; will retry counts.',
        e
      );
    }

    state.pendingRows.push({
      id: folder.getId(),
      values: [
        node.path,       // Path
        fileCount,       // File_Count
        subfolderCount,  // Subfolder_Count
        folderCreated,   // DtCreated
        folderUpdated    // DtModified
      ],
      url: folderUrl
    });

  } else {
    // BOTH mode
    try {
      // Folder itself as a row
      state.pendingRows.push({
        id: folder.getId(),
        values: [
          node.path,        // Path
          folder.getName(), // Name
          'Folder',         // Type
          '',               // Extension
          folderCreated,    // DtCreated
          folderUpdated     // DtModified
        ],
        url: folderUrl
      });
      if (state.pendingRows.length >= 1000) flushRows(outSheet, state);

      const fileIter = folder.getFiles();
      while (fileIter.hasNext()) {
        const f    = fileIter.next();
        const name = f.getName();
        const mime = f.getMimeType();
        const ext  = deriveExtension(name, mime);
        const url  = f.getUrl();
        const dc   = f.getDateCreated ? f.getDateCreated() : null;
        const dm   = f.getLastUpdated ? f.getLastUpdated() : null;

        state.pendingRows.push({
          id: f.getId(),
          values: [
            node.path, // Path
            name,      // Name
            'File',    // Type
            ext,       // Extension
            dc,        // DtCreated
            dm         // DtModified
          ],
          url: url
        });

        if (state.pendingRows.length >= 1000) flushRows(outSheet, state);
        if (Date.now() - sliceStartTs > (MAX_RUNTIME_MS - STOP_MARGIN_MS)) break;
      }

      if (canGoDeeper) {
        const folderIter = folder.getFolders();
        while (folderIter.hasNext()) {
          const sub     = folderIter.next();
          const subName = sub.getName();
          state.queue.push({ id: sub.getId(), path: `${node.path}/${subName}`, depth: node.depth + 1 });
          if (Date.now() - sliceStartTs > (MAX_RUNTIME_MS - STOP_MARGIN_MS)) break;
        }
      }
    } catch (e) {
      logSimple(
        state.sheetName || state.profileName,
        'SCANNING',
        'WARN',
        'Listing children failed in Both mode; will retry.',
        e
      );
    }
  }

  if (state.pendingRows.length) {
    flushRows(outSheet, state);
  }

  updateProgressNoteIfDue(state);
  return true;
}

/*******************************************************
 * FLUSH BUFFERED ROWS
 *******************************************************/

function flushRows(sheet, state) {
  try {
    const entries = state.pendingRows;
    if (!entries || !entries.length) return;

    const header = getEffectiveHeaderForState(state);

    const rows = entries.map(e => {
      const slNo = state.nextSlNo++;
      const base = [slNo, e.id].concat(e.values);
      if (state.showURL) {
        base.push(e.url || '');
      }
      return base;
    });

    const startRow = state.writeRow;
    const rng = sheet.getRange(startRow, 1, rows.length, header.length);
    rng.setValues(rows);

    // Hyperlink: Files/Both => "Name/File" column, Folders => Path column
    let linkColIndex;
    if (state.mode === 'folders') {
      linkColIndex = header.indexOf('Path');
    } else {
      linkColIndex = header.indexOf('File');
      if (linkColIndex === -1) {
        linkColIndex = header.indexOf('Name');
      }
    }
    const linkCol = linkColIndex >= 0 ? (linkColIndex + 1) : null;

    entries.forEach((entry, idx) => {
      if (!entry.url || linkCol == null) return;
      const rowIndex = startRow + idx;
      const cell = sheet.getRange(rowIndex, linkCol);
      const text = String(cell.getValue() || '');

      if (text) {
        const rich = SpreadsheetApp.newRichTextValue()
          .setText(text)
          .setLinkUrl(entry.url)
          .build();
        cell.setRichTextValue(rich);
      }
    });

    state.writeRow += rows.length;
    state.processedCount += rows.length;
    state.pendingRows = [];
  } catch (e) {
    logSimple(
      state.sheetName || state.profileName,
      state.phase || 'FLUSH',
      'WARN',
      'Buffered write failed; will retry.',
      e
    );
  }
}

/*******************************************************
 * MERGE-BACK SLICE (EXTRA COLUMNS)
 *******************************************************/

function processMergeBackSlice(state) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const newSheet = ss.getSheetByName(state.sheetName);
  if (!newSheet) {
    state.phase = 'DONE';
    return false;
  }

  const oldSheet = ss.getSheetByName(state.oldSheetName);
  if (!oldSheet) {
    state.phase = 'DONE';
    return false;
  }

  const mode       = state.mode || 'both';
  const baseHeader = getHeaderForMode(mode);
  const newHdr     = getHeader(newSheet);
  const oldHdr     = getHeader(oldSheet);

  const baseSet = new Set(baseHeader.concat(['URL']));
  const extraCols = oldHdr.filter(h => !baseSet.has(h) && h !== 'WebLinkURL');

  if (extraCols.length) {
    const newHdrBefore = getHeader(newSheet);
    const missing = extraCols.filter(c => newHdrBefore.indexOf(c) === -1);
    if (missing.length) {
      newSheet.insertColumnsAfter(newHdrBefore.length, missing.length);
      const fullHdr = newHdrBefore.concat(missing);
      newSheet.getRange(1, 1, 1, fullHdr.length).setValues([fullHdr]);
    }
  }

  const finalHdr  = getHeader(newSheet);
  const newColMap = indexMap(finalHdr);
  const oldColMap = indexMap(oldHdr);

  const oldKeyIdxs = getKeyIndexesForMode(mode, oldColMap);
  const newKeyIdxs = getKeyIndexesForMode(mode, newColMap);

  if (!oldKeyIdxs.length || !newKeyIdxs.length) {
    state.phase = 'DONE';
    return false;
  }

  const extraIdxs  = extraCols.map(c => oldColMap[c]);
  const targetIdxs = extraCols.map(c => newColMap[c]);

  const oldValues = readData(oldSheet);
  const oldMap = new Map();
  oldValues.forEach(row => {
    const key = makeKeyFromRow(row, oldKeyIdxs);
    if (!key) return;
    const extras = extraIdxs.map(i => row[i] || '');
    oldMap.set(key, extras);
  });

  const newValues = readData(newSheet);
  if (!newValues.length) {
    state.phase = 'DONE';
    return false;
  }

  const CHUNK = 1000;
  for (let i = 0; i < newValues.length; i += CHUNK) {
    const slice = newValues.slice(i, i + CHUNK);
    const writeCells = [];

    slice.forEach((row, offset) => {
      const key = makeKeyFromRow(row, newKeyIdxs);
      if (!key) return;
      const extras = oldMap.get(key);
      if (!extras) return;

      extras.forEach((val, exIdx) => {
        const r = 2 + i + offset;
        const c = 1 + targetIdxs[exIdx];
        writeCells.push({ r, c, v: val });
      });
    });

    writeCells.forEach(cell => {
      newSheet.getRange(cell.r, cell.c).setValue(cell.v);
    });

    if (timeAlmostUp()) break;
  }

  state.phase = 'DONE';
  return true;
}

/*******************************************************
 * FINALIZE PROFILE
 *******************************************************/

function finalizeProfile(state) {
  const profileName = state.profileName;
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Profiles');
    if (sheet) {
      const { rows, colMap, a1StartRow } = readTable(sheet);
      const idxProfileName = getColIndex(colMap, 'ProfileName');
      const idxScan        = getColIndex(colMap, 'Scan');
      const idxLastUpdated = getColIndex(colMap, 'LastUpdated');

      if (idxProfileName != null && idxScan != null && idxLastUpdated != null) {
        for (let i = 0; i < rows.length; i++) {
          const name = String(rows[i][idxProfileName] || '').trim();
          const scan = String(rows[i][idxScan]        || '').trim().toLowerCase();
          if (name === profileName && scan === 'yes') {
            sheet.getRange(a1StartRow + i, idxLastUpdated + 1).setValue(new Date());
            break;
          }
        }
      }
    }

    if (state.sheetName) {
      linkProfileToOutputSheet(profileName, state.sheetName);
    }

    const finishedAt = new Date();
    const startedAt  = state.startedAtISO ? new Date(state.startedAtISO) : finishedAt;
    const durationMs = finishedAt.getTime() - startedAt.getTime();
    const durationTxt = formatDuration(durationMs);

    const completedMsg =
      `Completed (${state.mode || 'both'}). ` +
      `${formatInt(state.processedCount)} items processed in ${durationTxt} at ${formatTime(finishedAt)}`;

    noteProgressOnProfiles(profileName, completedMsg, null);

    hideTechnicalColumns(state);

    if (state.sheetName) {
      const outSheet = ss.getSheetByName(state.sheetName);
      if (outSheet) {
        autoResizeColumns(outSheet, AUTO_RESIZE_HEADERS);
      }
    }

  } catch (e) {
    logSimple(
      state.sheetName || profileName,
      'FINALIZE',
      'WARN',
      'Could not update Profiles.LastUpdated or Remarks',
      e
    );
  } finally {
    clearState(profileName);
  }
}

/*******************************************************
 * LINK PROFILE ROW → OUTPUT SHEET
 *******************************************************/

function linkProfileToOutputSheet(profileName, outputSheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const profilesSheet = ss.getSheetByName('Profiles');
    if (!profilesSheet) return;

    const outputSheet = ss.getSheetByName(outputSheetName);
    if (!outputSheet) return;

    const { rows, colMap, a1StartRow } = readTable(profilesSheet);
    const idxProfileName = getColIndex(colMap, 'ProfileName');
    if (idxProfileName == null) return;

    for (let i = 0; i < rows.length; i++) {
      const name = String(rows[i][idxProfileName] || '').trim();
      if (name === profileName) {
        const rowNum = a1StartRow + i;
        const cell = profilesSheet.getRange(rowNum, idxProfileName + 1);

        const currentText = String(cell.getValue() || '').trim() || profileName;

        const rich = SpreadsheetApp.newRichTextValue()
          .setText(currentText)
          .setLinkUrl('#gid=' + outputSheet.getSheetId())
          .build();

        cell.setRichTextValue(rich);
        break;
      }
    }
  } catch (e) {
    logSimple(
      outputSheetName || profileName,
      'FINALIZE_LINK',
      'WARN',
      'Could not set hyperlink on Profiles sheet',
      e
    );
  }
}

/*******************************************************
 * CLEAR PROFILE HYPERLINK (before starting scan)
 *******************************************************/

function clearProfileHyperlink(profileName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Profiles');
    if (!sheet) return;

    const { rows, colMap, a1StartRow } = readTable(sheet);
    const idxProfileName = getColIndex(colMap, 'ProfileName');
    if (idxProfileName == null) return;

    for (let i = 0; i < rows.length; i++) {
      const name = String(rows[i][idxProfileName] || '').trim();
      if (name === profileName) {
        const rowNum = a1StartRow + i;
        const cell = sheet.getRange(rowNum, idxProfileName + 1);

        const currentText =
          String(cell.getDisplayValue() || cell.getValue() || '').trim() ||
          profileName;

        cell.setValue(currentText);
        cell.setFontColor(null);
        cell.setFontLine('none');
        break;
      }
    }
  } catch (e) {
    logSimple(
      profileName,
      'CLEAR_PROFILE_LINK',
      'WARN',
      'Could not clear hyperlink on Profiles sheet',
      e
    );
  }
}

/*******************************************************
 * PROGRESS NOTES (Profiles.Remarks)
 *******************************************************/

function updateProgressNoteIfDue(state, forceStartMessage) {
  const now = Date.now();
  const profileName = state.profileName;
  const mode = state.mode || 'both';
  const runningColor = 'red';

  if (forceStartMessage) {
    const msg = `Started (${mode}) at ${formatTime(new Date(state.startedAtISO || now))}`;
    noteProgressOnProfiles(profileName, msg, runningColor);
    state.lastProgressNoteAt = now;
    return;
  }

  if (!state.lastProgressNoteAt || (now - state.lastProgressNoteAt) >= PROGRESS_INTERVAL_MS) {
    const startedTime = formatTimeShort(new Date(state.startedAtISO || now));
    const lastTime    = formatTimeShort(new Date());

    const msg =
      `Running (${mode})… ` +
      `${formatInt(state.processedCount)} processed, ` +
      `queue=${formatInt(state.queue.length)} ` +
      `(Started ${startedTime}, Last updated ${lastTime})`;

    noteProgressOnProfiles(profileName, msg, runningColor);
    state.lastProgressNoteAt = now;
  }
}

/*******************************************************
 * MODE HELPERS & HEADERS
 *******************************************************/

function normalizeMode(modeRaw) {
  const m = String(modeRaw || '').trim().toLowerCase();
  if (m === 'files' || m === 'file')     return 'files';
  if (m === 'folders' || m === 'folder') return 'folders';
  return 'both';
}

function parseMaxDepth(maxDepthRaw) {
  const s = String(maxDepthRaw || '').trim();
  if (!s) return 0;
  const n = parseInt(s, 10);
  if (isNaN(n) || n < 0) return 0;
  return n;
}

function modeToSheetSuffix(mode) {
  if (mode === 'files')   return 'Files';
  if (mode === 'folders') return 'Folders';
  return 'Both';
}

function getHeaderForMode(mode) {
  switch (mode) {
    case 'files':
      return ['SlNo', 'ID', 'Path', 'File', 'Extension', 'DtCreated', 'DtModified'];
    case 'folders':
      return ['SlNo', 'ID', 'Path', 'File_Count', 'Subfolder_Count', 'DtCreated', 'DtModified'];
    default:
      return ['SlNo', 'ID', 'Path', 'Name', 'Type', 'Extension', 'DtCreated', 'DtModified'];
  }
}

function getEffectiveHeaderForState(state) {
  const base = getHeaderForMode(state.mode || 'both');
  if (state.showURL) {
    return base.concat(['URL']);
  }
  return base;
}

/*******************************************************
 * KEY GENERATION FOR MERGE-BACK
 *******************************************************/

function getKeyIndexesForMode(mode, colMap) {
  const idxs = [];

  if (colMap['ID'] != null) {
    idxs.push(colMap['ID']);
    return idxs;
  }

  function addFirst(names) {
    for (let i = 0; i < names.length; i++) {
      const n = names[i];
      if (colMap[n] != null) {
        idxs.push(colMap[n]);
        return;
      }
    }
  }

  if (mode === 'files') {
    addFirst(['Path', 'Folder_Name', 'GivenFolder']);
    addFirst(['File', 'Name']);
    if (colMap['Extension'] != null) idxs.push(colMap['Extension']);
  } else if (mode === 'folders') {
    addFirst(['Path', 'Folder_Name', 'GivenFolder']);
  } else {
    addFirst(['Path', 'Folder_Name', 'GivenFolder']);
    addFirst(['Name', 'File']);
    if (colMap['Type'] != null) idxs.push(colMap['Type']);
    if (colMap['Extension'] != null) idxs.push(colMap['Extension']);
  }

  return idxs;
}

function makeKeyFromRow(row, idxs) {
  if (!idxs || !idxs.length) return null;
  const parts = idxs.map(i => String(row[i] || ''));
  return parts.join('||');
}

/*******************************************************
 * SHEET / HEADER HELPERS
 *******************************************************/

function ensureOutputSheetForState(ss, state) {
  const header = getEffectiveHeaderForState(state);
  const sheet  = getOrCreateSheetWithHeader(ss, state.sheetName, header);
  hideIdColumnIfPresent(sheet);
  applyColumnLayoutForMode(sheet, state.mode || 'both', !!state.showURL);
  return sheet;
}

function getOrCreateSheetWithHeader(ss, name, header) {
  let sh = ss.getSheetByName(name);
  if (!sh) {
    const active = ss.getActiveSheet();

    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    sh.setFrozenRows(1);

    if (active) {
      ss.setActiveSheet(active);
      active.activate();
    }
    return sh;
  }

  const existingHdr = getHeader(sh);
  if (existingHdr.length < header.length) {
    sh.getRange(1, 1, 1, header.length).setValues([header]);
  }
  return sh;
}

function prepareRolloverSheets(baseName, mode, showURL) {
  const ss   = SpreadsheetApp.getActiveSpreadsheet();
  const curr = ss.getSheetByName(baseName);
  const oldName = baseName + '_Old';
  const old  = ss.getSheetByName(oldName);

  if (old) ss.deleteSheet(old);
  if (curr) curr.setName(oldName);

  const baseHeader = getHeaderForMode(mode);
  const header     = showURL ? baseHeader.concat(['URL']) : baseHeader;
  const fresh      = getOrCreateSheetWithHeader(ss, baseName, header);

  hideIdColumnIfPresent(fresh);
  applyColumnLayoutForMode(fresh, mode, showURL);
  autoResizeColumns(fresh, AUTO_RESIZE_HEADERS);

  return { sheetName: fresh.getName(), oldSheetName: oldName };
}

function getHeader(sheet) {
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return [];
  const hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  return hdr.map(h => String(h || '').trim());
}

function readData(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
}

function applyColumnLayoutForMode(sheet, mode, showURL) {
  try {
    const header = getHeader(sheet);
    if (!header || !header.length) return;
    const colMap = indexMap(header);

    function setWidth(hName, width) {
      const idx = colMap[hName];
      if (idx == null) return;
      sheet.setColumnWidth(idx + 1, width);
    }

    setWidth('SlNo', COL_WIDTH_SLNO);
    setWidth('ID',   COL_WIDTH_ID);
    setWidth('Path', COL_WIDTH_PATH);

    if (mode === 'files') {
      setWidth('File',      COL_WIDTH_NAME);
      setWidth('Extension', COL_WIDTH_EXT);
    } else if (mode === 'folders') {
      setWidth('File_Count',      COL_WIDTH_COUNT);
      setWidth('Subfolder_Count', COL_WIDTH_COUNT);
    } else {
      setWidth('Name',      COL_WIDTH_NAME);
      setWidth('Type',      COL_WIDTH_TYPE);
      setWidth('Extension', COL_WIDTH_EXT);
    }

    setWidth('DtCreated',  COL_WIDTH_DATE);
    setWidth('DtModified', COL_WIDTH_DATE);

    if (showURL) {
      setWidth('URL', COL_WIDTH_URL);
    }
  } catch (e) {
    logSimple(
      sheet ? sheet.getName() : '',
      'LAYOUT',
      'WARN',
      'Could not apply column layout.',
      e
    );
  }
}

/*******************************************************
 * GENERIC TABLE HELPERS
 *******************************************************/

function indexMap(arr) {
  const m = {};
  arr.forEach((v, i) => { m[String(v)] = i; });
  return m;
}

function readTable(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2) return { rows: [], colMap: {}, a1StartRow: 2 };
  const hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(x => String(x).trim());
  const colMap = {};
  hdr.forEach((h, i) => colMap[h] = i);
  const rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  return { rows, colMap, a1StartRow: 2 };
}

function getColIndex(colMap, baseName) {
  const target = String(baseName || '').toLowerCase();
  for (const key in colMap) {
    const clean = String(key || '').replace(/^\*/, '').toLowerCase();
    if (clean === target) return colMap[key];
  }
  return null;
}

/*******************************************************
 * GENERIC HELPERS
 *******************************************************/

function deriveExtension(name, mimeType) {
  if (mimeType && mimeType.indexOf('application/vnd.google-apps') === 0) {
    if (mimeType === 'application/vnd.google-apps.spreadsheet')  return 'gsheet';
    if (mimeType === 'application/vnd.google-apps.document')     return 'gdoc';
    if (mimeType === 'application/vnd.google-apps.presentation') return 'gslide';
    return '';
  }
  const dot = name.lastIndexOf('.');
  if (dot > 0 && dot < name.length - 1) return name.substring(dot + 1);
  return '';
}

function formatTime(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function formatTimeShort(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm:ss');
}

function formatDuration(ms) {
  if (ms < 0 || !isFinite(ms)) return '0s';
  let seconds = Math.floor(ms / 1000);
  let minutes = Math.floor(seconds / 60);
  let hours   = Math.floor(minutes / 60);

  seconds = seconds % 60;
  minutes = minutes % 60;

  const parts = [];
  if (hours)   parts.push(hours + 'h');
  if (minutes) parts.push(minutes + 'm');
  if (!hours && !minutes) {
    parts.push(seconds + 's');
  } else if (seconds) {
    parts.push(seconds + 's');
  }

  return parts.join(' ');
}

function formatInt(n) {
  return String(n || 0).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
}

/*******************************************************
 * STATE (DocumentProperties)
 *******************************************************/

function getState(profileName) {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(STATE_PREFIX + profileName);
  return raw ? JSON.parse(raw) : null;
}

function setState(profileName, state) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(STATE_PREFIX + profileName, JSON.stringify(state));
}

function clearState(profileName) {
  const props = PropertiesService.getDocumentProperties();
  props.deleteProperty(STATE_PREFIX + profileName);
}

function getActiveProfiles() {
  const props = PropertiesService.getDocumentProperties();
  const raw = props.getProperty(ACTIVE_SET);
  try { return raw ? JSON.parse(raw) : []; } catch { return []; }
}

function setActiveProfiles(arr) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty(ACTIVE_SET, JSON.stringify(arr || []));
}

function discoverActiveProfilesFromState() {
  const props = PropertiesService.getDocumentProperties().getProperties();
  const names = [];
  Object.keys(props).forEach(k => {
    if (k.startsWith(STATE_PREFIX)) {
      const st = JSON.parse(props[k]);
      if (st && st.phase && st.phase !== 'DONE') names.push(st.profileName);
    }
  });
  return names;
}

function anyScanningInProgress() {
  const props = PropertiesService.getDocumentProperties().getProperties();
  return Object.keys(props).some(k => {
    if (!k.startsWith(STATE_PREFIX)) return false;
    const st = JSON.parse(props[k]);
    return st && st.phase && st.phase !== 'DONE';
  });
}

/*******************************************************
 * TRIGGERS — single worker
 *******************************************************/

function ensureSingleTrigger(removeOnly) {
  const props = PropertiesService.getDocumentProperties();
  const storedId = props.getProperty(TRIGGER_KEY);
  const all = ScriptApp.getProjectTriggers();

  all.forEach(t => {
    if (t.getHandlerFunction() === 'scanWorker') {
      ScriptApp.deleteTrigger(t);
    }
  });
  if (storedId) props.deleteProperty(TRIGGER_KEY);

  if (removeOnly) return;
}

function scheduleNextWorker(ms) {
  const props = PropertiesService.getDocumentProperties();
  const trig = ScriptApp.newTrigger('scanWorker')
    .timeBased()
    .after(ms)
    .create();
  props.setProperty(TRIGGER_KEY, trig.getUniqueId ? trig.getUniqueId() : String(Date.now()));
}

/*******************************************************
 * TIME GUARD
 *******************************************************/

function timeAlmostUp() {
  if (!SLICE_START_TS) return false;
  return (Date.now() - SLICE_START_TS) > (MAX_RUNTIME_MS - STOP_MARGIN_MS);
}

/*******************************************************
 * LOGGING & PROFILES REMARKS
 *******************************************************/

function logEvent({ level='INFO', profile='', phase='', msg='', details='' }) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = getOrCreateSheetWithHeader(ss, 'Log', ['Timestamp','Level','Profile','Phase','Message','Details']);
    const row = [new Date(), level, profile, phase, msg, details];
    sh.appendRow(row);
  } catch (e) {
    // ignore logging errors
  }
}

function logSimple(profile, phase, level, msg, err) {
  const details =
    err && err.stack ? err.stack :
    err && err.message ? err.message :
    (err || '');

  logEvent({
    level: level || 'INFO',
    profile: profile || '',
    phase: phase || '',
    msg: msg || '',
    details: details
  });
}

/**
 * Write Remarks for a profile, and optionally control font color.
 */
function noteProgressOnProfiles(profileName, text, fontColor) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Profiles');
    if (!sheet) return;
    const { rows, colMap, a1StartRow } = readTable(sheet);
    const idxRemarks     = getColIndex(colMap, 'Remarks');
    const idxProfileName = getColIndex(colMap, 'ProfileName');
    if (idxRemarks == null || idxProfileName == null) return;
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][idxProfileName]).trim() === profileName) {
        const cell = sheet.getRange(a1StartRow + i, idxRemarks + 1);
        cell.setValue(text);
        if (fontColor !== undefined) {
          cell.setFontColor(fontColor);
        }
        break;
      }
    }
  } catch (_e) {}
}

/*******************************************************
 * HIDE TECHNICAL COLUMNS (like ID)
 *******************************************************/

function hideIdColumnIfPresent(sheet) {
  try {
    const header = getHeader(sheet);
    if (!header || !header.length) return;

    const idxId = header.indexOf('ID');
    if (idxId >= 0) {
      const col = idxId + 1;
      sheet.hideColumn(sheet.getRange(1, col));
    }
  } catch (e) {
    logSimple(
      sheet ? sheet.getName() : '',
      'HIDE_ID_COLUMN',
      'WARN',
      'Could not hide ID column.',
      e
    );
  }
}

function hideTechnicalColumns(state) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(state.sheetName);
    if (!sheet) return;
    hideIdColumnIfPresent(sheet);
  } catch (e) {
    logSimple(
      state.sheetName || state.profileName,
      'HIDE_TECH_COLUMNS',
      'WARN',
      'Could not hide technical columns.',
      e
    );
  }
}

function autoResizeColumns(sheet, headerNamesArray) {
  try {
    const header = getHeader(sheet);
    if (!header || !header.length) return;

    headerNamesArray.forEach(name => {
      const idx = header.indexOf(name);
      if (idx >= 0) {
        sheet.autoResizeColumn(idx + 1);
      }
    });
  } catch (e) {
    logSimple(
      sheet ? sheet.getName() : '',
      'AUTO_RESIZE_COLUMNS',
      'WARN',
      'Failed to auto-resize columns',
      e
    );
  }
}

/*******************************************************
 * CANCEL SCANNING
 *******************************************************/

function cancelAllScans() {
  runSafely('cancelAllScans', () => {
    const props = PropertiesService.getDocumentProperties();
    const allProps = props.getProperties();

    const cancelledProfiles = [];
    Object.keys(allProps).forEach(key => {
      if (key.startsWith(STATE_PREFIX)) {
        try {
          const st = JSON.parse(allProps[key]);
          if (st && st.profileName) {
            cancelledProfiles.push(st.profileName);
          }
        } catch (e) {
          // ignore bad JSON, just delete it
        }
        props.deleteProperty(key);
      }
    });

    props.deleteProperty(ACTIVE_SET);

    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(t => {
      if (t.getHandlerFunction() === 'scanWorker') {
        ScriptApp.deleteTrigger(t);
      }
    });
    props.deleteProperty(TRIGGER_KEY);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Profiles');
    let displayCancelled = [];

    if (sheet && cancelledProfiles.length) {
      const { rows, colMap, a1StartRow } = readTable(sheet);
      const idxProfileName = getColIndex(colMap, 'ProfileName');
      const idxRemarks     = getColIndex(colMap, 'Remarks');
      const idxScan        = getColIndex(colMap, 'Scan');

      if (idxProfileName != null && idxRemarks != null) {
        for (let i = 0; i < rows.length; i++) {
          const name = String(rows[i][idxProfileName] || '').trim();
          const scanVal = idxScan != null
            ? String(rows[i][idxScan] || '').trim().toLowerCase()
            : '';

          if (cancelledProfiles.indexOf(name) !== -1) {
            const rowNum = a1StartRow + i;
            sheet.getRange(rowNum, idxRemarks + 1).setValue('');

            if (scanVal === 'yes') {
              displayCancelled.push(name);
            }
          }
        }
      }
    }

    const ui = SpreadsheetApp.getUi();
    if (displayCancelled.length) {
      ui.alert('Scanning cancelled for Scan="Yes" profiles: ' + displayCancelled.join(', '));
    } else if (cancelledProfiles.length) {
      ui.alert('Scanning cancelled. (No profiles currently marked Scan="Yes".)');
    } else {
      ui.alert('No active scans found. Nothing to cancel.');
    }
  });
}
