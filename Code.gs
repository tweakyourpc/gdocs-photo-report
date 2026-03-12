'use strict';

const PHOTO_REPORT_CONFIG = Object.freeze({
  folderPropertyKey: 'PHOTO_REPORT_FOLDER_ID',
  folderNamePropertyKey: 'PHOTO_REPORT_FOLDER_NAME',
  runStatePropertyKey: 'PHOTO_REPORT_RUN_STATE',
  userPickerKeyPropertyKey: 'PHOTO_REPORT_PICKER_API_KEY',
  userPickerProjectPropertyKey: 'PHOTO_REPORT_PICKER_PROJECT_NUMBER',
  markerPrefix: 'PHOTO_REPORT',
  captionPattern: /^\s*(\d+)\s*[-.):]\s*(.+\S)\s*$/,
  maxImageWidthPx: 520,
  maxExecutionMs: 270000,
  maxBatchMs: 15000,
  maxInsertionsPerBatch: 12,
  maxRemovalsPerBatch: 20,
  sidebarTitle: 'Photo Report',
  defaultPickerDeveloperKey: '',
  defaultPickerCloudProjectNumber: '',
  diagnosticDocPrefix: 'Photo Report Diagnostic Log',
  pickerDialogWidth: 720,
  pickerDialogHeight: 560,
  maxConsoleLines: 24,
  maxStoredErrors: 24,
  maxStoredErrorDetails: 12,
});

const PHOTO_REPORT_ACTIONS = Object.freeze({
  insert: Object.freeze({
    key: 'insert',
    label: 'Insert Missing Images',
    rebuild: false,
  }),
  rebuild: Object.freeze({
    key: 'rebuild',
    label: 'Rebuild Images',
    rebuild: true,
  }),
});

function onOpen() {
  DocumentApp.getUi()
      .createMenu(PHOTO_REPORT_CONFIG.sidebarTitle)
      .addItem('Open Sidebar', 'showSidebar')
      .addSeparator()
      .addItem('Set Image Folder', 'showPicker')
      .addItem('Insert Missing Images', 'showSidebarForInsertMissingImages')
      .addItem('Rebuild Images', 'showSidebarForRebuildImages')
      .addToUi();
}

function showSidebar() {
  showSidebar_('');
}

function showSidebarForFolderSelection() {
  showSidebar_('folder');
}

function showSidebarForInsertMissingImages() {
  showSidebar_('insert');
}

function showSidebarForRebuildImages() {
  showSidebar_('rebuild');
}

function showSidebar_(requestedAction) {
  const template = HtmlService.createTemplateFromFile('Sidebar');
  template.bootstrapJson = JSON.stringify(buildSidebarModel_(requestedAction || ''));

  const html = template.evaluate()
      .setTitle(PHOTO_REPORT_CONFIG.sidebarTitle)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  DocumentApp.getUi().showSidebar(html);
}

function showPicker() {
  const template = HtmlService.createTemplateFromFile('Picker');
  template.bootstrapJson = JSON.stringify(buildPickerDialogModel_());

  const html = template.evaluate()
      .setWidth(PHOTO_REPORT_CONFIG.pickerDialogWidth)
      .setHeight(PHOTO_REPORT_CONFIG.pickerDialogHeight)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  DocumentApp.getUi().showModalDialog(html, 'Select Photo Folder');
}

function openPickerDialog() {
  showPicker();
  return true;
}

function getSidebarState() {
  return buildSidebarModel_('');
}

function buildSidebarModel_(requestedAction) {
  return {
    requestedAction: requestedAction,
    folder: getConfiguredFolderInfo_(),
    picker: buildPickerClientState_(),
    run: buildClientRunState_(getStoredRunState_()),
  };
}

function buildPickerDialogModel_() {
  return {
    folder: getConfiguredFolderInfo_(),
    picker: buildPickerClientState_(),
    title: PHOTO_REPORT_CONFIG.sidebarTitle,
  };
}

function getOAuthToken() {
  return ScriptApp.getOAuthToken();
}

function saveFolderInput(input) {
  const folderId = extractDriveFolderId_(input);
  const folder = saveConfiguredFolderById_(
      folderId,
      'Could not save that Drive folder. Make sure the folder URL or ID is valid and accessible.',
  );

  return buildSidebarModelAfterFolderSave_(folder);
}

function handlePickerResponse(id, folderName) {
  if (!id) {
    throw new Error('No Drive folder was selected.');
  }

  const folder = saveConfiguredFolderById_(
      String(id),
      'The selected Drive folder could not be opened.',
      folderName,
  );

  return buildSidebarModelAfterFolderSave_(folder);
}

function savePickerCredentials(apiKey, cloudProjectNumber) {
  const normalizedApiKey = String(apiKey || '').trim();
  const normalizedProjectNumber = String(cloudProjectNumber || '').trim();

  if (!normalizedApiKey) {
    throw new Error('Enter a Google Picker API key before saving.');
  }

  if (!normalizedProjectNumber || !/^\d+$/.test(normalizedProjectNumber)) {
    throw new Error('Enter the Google Cloud project number as digits only.');
  }

  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty(PHOTO_REPORT_CONFIG.userPickerKeyPropertyKey, normalizedApiKey);
  userProperties.setProperty(
      PHOTO_REPORT_CONFIG.userPickerProjectPropertyKey,
      normalizedProjectNumber,
  );

  return buildPickerClientState_();
}

function clearPickerCredentials() {
  const userProperties = PropertiesService.getUserProperties();
  userProperties.deleteProperty(PHOTO_REPORT_CONFIG.userPickerKeyPropertyKey);
  userProperties.deleteProperty(PHOTO_REPORT_CONFIG.userPickerProjectPropertyKey);

  return buildPickerClientState_();
}

function startPhotoReportRun(actionKey) {
  return withPhotoReportLock_(function() {
    return executePhotoReportBatch_(actionKey, true);
  });
}

function resumePhotoReportRun() {
  return withPhotoReportLock_(function() {
    const runState = getStoredRunState_();
    if (!runState || !runState.action) {
      return buildClientRunState_(runState);
    }

    return executePhotoReportBatch_(runState.action, false);
  });
}

function clearPhotoReportRun() {
  clearStoredRunState_();
  return buildClientRunState_(null);
}

function withPhotoReportLock_(callback) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(1000)) {
    const runState = getStoredRunState_();
    if (runState) {
      runState.status = 'running';
      runState.message = 'Another batch is already running. Waiting for it to finish...';
      runState.updatedAt = new Date().toISOString();
      return buildClientRunState_(runState);
    }

    throw new Error('Another Photo Report batch is already running.');
  }

  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function executePhotoReportBatch_(actionKey, resetRunState) {
  const action = getPhotoReportAction_(actionKey);
  const batchLimits = buildBatchLimits_();
  let runState = resetRunState ? null : getStoredRunState_();

  try {
    const folder = getConfiguredFolder_();
    const body = DocumentApp.getActiveDocument().getBody();
    const captionScan = collectCaptionEntries_(body);
    validateCaptionEntries_(captionScan);

    runState = initializeRunState_(runState, action, folder, captionScan, resetRunState);
    runState.status = 'running';
    runState.phase = action.rebuild && !runState.removalComplete ? 'remove' : 'insert';
    runState.limitReached = false;
    runState.needsContinuation = false;
    runState.updatedAt = new Date().toISOString();
    runState.message = buildRunPhaseMessage_(runState);
    if (resetRunState) {
      appendRunLog_(runState, action.label + ' started.');
    } else {
      appendRunLog_(runState, action.label + ' resumed.');
    }
    saveProgressSnapshot_(runState);

    if (action.rebuild && !runState.removalComplete) {
      appendRunLog_(runState, 'Clearing previously managed images before rebuild.');
      saveProgressSnapshot_(runState);
      const removalSummary = removeManagedImageParagraphsBatch_(body, batchLimits, runState);
      runState.updatedAt = new Date().toISOString();

      if (!removalSummary.complete) {
        return pauseRunState_(
            runState,
            'remove',
            removalSummary.limitReached,
          removalSummary.limitReached ?
            'Execution time limit reached while clearing managed images.' :
            'Removing previously managed images...',
        );
      }

      runState.removalComplete = true;
      runState.phase = 'insert';
      runState.message = buildRunPhaseMessage_(runState);
      appendRunLog_(runState, 'Managed image removal complete. Starting insert phase.');
      saveProgressSnapshot_(runState);
    }

    const folderImages = collectFolderImages_(folder);
    runState.duplicateFolderNumbers = folderImages.duplicateNumbers;

    const captionsDescending = captionScan.entries
        .slice()
        .sort(function(left, right) {
          return right.childIndex - left.childIndex;
        });

    let processedThisBatch = 0;

    for (const captionEntry of captionsDescending) {
      if (runState.completedNumbers.indexOf(captionEntry.number) !== -1) {
        continue;
      }

      const pauseReason = getInsertPauseReason_(processedThisBatch, batchLimits);
      if (pauseReason) {
        return pauseRunState_(
            runState,
            'insert',
            pauseReason === 'limit',
          pauseReason === 'limit' ?
            'Execution time limit reached. Resume to continue processing the remaining photos.' :
            'Continuing in the next batch...',
        );
      }

      runState.message = buildProcessingMessage_(runState);
      saveProgressSnapshot_(runState);
      processCaptionEntry_(body, captionEntry, folderImages, runState, action);
      processedThisBatch += 1;
    }

    runState.phase = 'complete';
    runState.status = 'complete';
    runState.needsContinuation = false;
    runState.limitReached = false;
    runState.updatedAt = new Date().toISOString();
    runState.message = buildCompletionMessage_(runState);
    appendRunLog_(runState, 'Batch complete. ' + runState.message);
    saveRunState_(runState);
    return buildClientRunState_(runState);
  } catch (error) {
    return finalizeRunFailure_(action, runState, error);
  }
}

function buildBatchLimits_() {
  const startedAtMs = Date.now();

  return {
    startedAtMs: startedAtMs,
    hardDeadlineMs: startedAtMs + PHOTO_REPORT_CONFIG.maxExecutionMs,
    softDeadlineMs: startedAtMs + PHOTO_REPORT_CONFIG.maxBatchMs,
    maxInsertionsPerBatch: PHOTO_REPORT_CONFIG.maxInsertionsPerBatch,
    maxRemovalsPerBatch: PHOTO_REPORT_CONFIG.maxRemovalsPerBatch,
  };
}

function initializeRunState_(runState, action, folder, captionScan, resetRunState) {
  const captionNumbers = captionScan.entries.map(function(entry) {
    return entry.number;
  });

  if (resetRunState || !runState || runState.action !== action.key) {
    return createRunState_(action, folder, captionNumbers.length);
  }

  const nextState = runState;
  nextState.actionLabel = action.label;
  nextState.folderId = folder.getId();
  nextState.folderName = folder.getName();
  nextState.totalCount = captionNumbers.length;
  nextState.completedNumbers = filterNumbersToKnownSet_(nextState.completedNumbers, captionNumbers);
  nextState.missingNumbers = filterNumbersToKnownSet_(nextState.missingNumbers, captionNumbers);
  nextState.insertedCount = Number(nextState.insertedCount || 0);
  nextState.removedCount = Number(nextState.removedCount || 0);
  nextState.skippedCount = Number(nextState.skippedCount || 0);
  nextState.errors = ensureStringArray_(nextState.errors);
  nextState.errorDetails = ensureObjectArray_(nextState.errorDetails);
  nextState.consoleLines = ensureRecentStringArray_(nextState.consoleLines);
  nextState.duplicateFolderNumbers = ensureNumberArray_(nextState.duplicateFolderNumbers);
  nextState.removalComplete = Boolean(
      nextState.removalComplete || action.rebuild === false,
  );
  nextState.phase = nextState.removalComplete ? 'insert' : 'remove';
  nextState.startedAt = nextState.startedAt || new Date().toISOString();
  nextState.updatedAt = new Date().toISOString();
  nextState.diagnosticLogUrl = String(nextState.diagnosticLogUrl || '');
  return nextState;
}

function createRunState_(action, folder, totalCount) {
  return {
    action: action.key,
    actionLabel: action.label,
    status: 'running',
    phase: action.rebuild ? 'remove' : 'insert',
    folderId: folder.getId(),
    folderName: folder.getName(),
    totalCount: totalCount,
    completedNumbers: [],
    missingNumbers: [],
    duplicateFolderNumbers: [],
    insertedCount: 0,
    removedCount: 0,
    skippedCount: 0,
    errors: [],
    errorDetails: [],
    consoleLines: [],
    limitReached: false,
    needsContinuation: false,
    removalComplete: action.rebuild === false,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    message: action.rebuild ? 'Removing previously managed images...' : 'Preparing images...',
    diagnosticLogUrl: '',
  };
}

function pauseRunState_(runState, phase, limitReached, message) {
  runState.phase = phase;
  runState.limitReached = limitReached;
  runState.needsContinuation = needsMoreWork_(runState);
  runState.status = runState.needsContinuation ? (limitReached ? 'paused' : 'running') : 'complete';
  runState.updatedAt = new Date().toISOString();
  runState.message = runState.needsContinuation ?
    buildPauseMessage_(runState, message) :
    buildCompletionMessage_(runState);
  appendRunLog_(runState, runState.message);

  if (!runState.needsContinuation) {
    runState.phase = 'complete';
    runState.limitReached = false;
  }

  saveRunState_(runState);
  return buildClientRunState_(runState);
}

function needsMoreWork_(runState) {
  if (!runState.removalComplete) {
    return true;
  }

  return runState.completedNumbers.length < runState.totalCount;
}

function getInsertPauseReason_(processedThisBatch, batchLimits) {
  if (Date.now() > batchLimits.hardDeadlineMs) {
    return 'limit';
  }

  if (
    processedThisBatch >= batchLimits.maxInsertionsPerBatch ||
    Date.now() > batchLimits.softDeadlineMs
  ) {
    return 'batch';
  }

  return '';
}

function processCaptionEntry_(body, captionEntry, folderImages, runState, action) {
  const imageRecord = folderImages.byNumber[captionEntry.number];

  if (!imageRecord) {
    pushUniqueNumber_(runState.missingNumbers, captionEntry.number);
    pushUniqueNumber_(runState.completedNumbers, captionEntry.number);
    runState.message = 'Caption ' + captionEntry.number + ': no matching image found.';
    appendRunLog_(runState, runState.message);
    saveProgressSnapshot_(runState);
    return;
  }

  if (
    !action.rebuild &&
    hasManagedImageAbove_(body, captionEntry.childIndex, captionEntry.number)
  ) {
    runState.skippedCount += 1;
    pushUniqueNumber_(runState.completedNumbers, captionEntry.number);
    runState.message = 'Caption ' + captionEntry.number + ': existing managed image found, skipped.';
    appendRunLog_(runState, runState.message);
    saveProgressSnapshot_(runState);
    return;
  }

  try {
    insertManagedImageAboveCaption_(
        body,
        captionEntry.childIndex,
        captionEntry.number,
        imageRecord.file,
    );
    runState.insertedCount += 1;
    runState.message = 'Caption ' + captionEntry.number + ': inserted ' + imageRecord.file.getName() + '.';
    appendRunLog_(runState, runState.message);
  } catch (error) {
    const errorDetail = {
      type: 'insert',
      number: captionEntry.number,
      fileId: imageRecord.file.getId(),
      fileName: imageRecord.file.getName(),
      message: error.message,
    };

    runState.errors.push(formatDiagnosticError_(errorDetail));
    runState.errorDetails.push(errorDetail);
    runState.message = 'Caption ' + captionEntry.number + ': failed to insert ' +
      imageRecord.file.getName() + '.';
    appendRunLog_(runState, runState.message + ' ' + formatDiagnosticError_(errorDetail));
  }

  pushUniqueNumber_(runState.completedNumbers, captionEntry.number);
  saveProgressSnapshot_(runState);
}

function finalizeRunFailure_(action, runState, error) {
  const nextState = runState || createFallbackRunState_(action);
  const errorMessage = error && error.message ? error.message : String(error);

  nextState.status = 'error';
  nextState.phase = 'error';
  nextState.needsContinuation = false;
  nextState.limitReached = false;
  nextState.updatedAt = new Date().toISOString();
  nextState.errors = ensureStringArray_(nextState.errors);
  nextState.errorDetails = ensureObjectArray_(nextState.errorDetails);
  nextState.consoleLines = ensureRecentStringArray_(nextState.consoleLines);
  nextState.errors.push(errorMessage);
  nextState.errorDetails.push({
    type: 'fatal',
    number: '',
    fileId: '',
    fileName: '',
    message: errorMessage,
  });
  nextState.message = 'Run failed. Review the diagnostic log for details.';
  appendRunLog_(nextState, errorMessage);
  nextState.diagnosticLogUrl = createDiagnosticLogSafely_(nextState);

  saveRunState_(nextState);
  return buildClientRunState_(nextState);
}

function createFallbackRunState_(action) {
  const activeDocument = DocumentApp.getActiveDocument();
  const folderInfo = getConfiguredFolderInfo_();

  return {
    action: action ? action.key : '',
    actionLabel: action ? action.label : 'Photo Report',
    status: 'error',
    phase: 'error',
    folderId: folderInfo ? folderInfo.id : '',
    folderName: folderInfo ? folderInfo.name : '',
    totalCount: 0,
    completedNumbers: [],
    missingNumbers: [],
    duplicateFolderNumbers: [],
    insertedCount: 0,
    removedCount: 0,
    skippedCount: 0,
    errors: [],
    errorDetails: [],
    consoleLines: [],
    limitReached: false,
    needsContinuation: false,
    removalComplete: action ? action.rebuild === false : true,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    message: '',
    diagnosticLogUrl: '',
    documentId: activeDocument.getId(),
    documentName: activeDocument.getName(),
  };
}

function createDiagnosticLogSafely_(runState) {
  try {
    return createDiagnosticLog_(runState);
  } catch (diagnosticError) {
    runState.errors.push(
        'Could not create diagnostic log: ' + (diagnosticError.message || String(diagnosticError)),
    );
    return '';
  }
}

function createDiagnosticLog_(runState) {
  const activeDocument = DocumentApp.getActiveDocument();
  const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss',
  );
  const diagnosticDocument = DocumentApp.create(
      PHOTO_REPORT_CONFIG.diagnosticDocPrefix + ' - ' + timestamp,
  );
  const diagnosticBody = diagnosticDocument.getBody();

  diagnosticBody.appendParagraph('Photo Report Diagnostic Log')
      .setHeading(DocumentApp.ParagraphHeading.TITLE);
  diagnosticBody.appendParagraph('Generated: ' + timestamp);
  diagnosticBody.appendParagraph('Action: ' + (runState.actionLabel || runState.action || 'Unknown'));
  diagnosticBody.appendParagraph(
      'Source document: ' + activeDocument.getName() + ' (' + activeDocument.getId() + ')',
  );

  if (runState.folderName || runState.folderId) {
    diagnosticBody.appendParagraph(
        'Photo folder: ' + (runState.folderName || 'Unknown') + ' (' + (runState.folderId || 'n/a') + ')',
    );
  }

  diagnosticBody.appendParagraph(
      'Progress: ' +
      String(runState.completedNumbers ? runState.completedNumbers.length : 0) +
      ' of ' +
      String(runState.totalCount || 0),
  );

  if (runState.errors && runState.errors.length) {
    diagnosticBody.appendParagraph('Errors').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    for (const errorMessage of runState.errors) {
      diagnosticBody.appendListItem(errorMessage);
    }
  }

  if (runState.errorDetails && runState.errorDetails.length) {
    diagnosticBody.appendParagraph('Detailed file diagnostics')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);

    for (const detail of runState.errorDetails) {
      diagnosticBody.appendListItem(
          [
            'Type: ' + String(detail.type || 'unknown'),
            'Caption: ' + String(detail.number || 'n/a'),
            'File name: ' + String(detail.fileName || 'n/a'),
            'File ID: ' + String(detail.fileId || 'n/a'),
            'Message: ' + String(detail.message || 'n/a'),
          ].join(' | '),
      );
    }
  }

  if (runState.duplicateFolderNumbers && runState.duplicateFolderNumbers.length) {
    diagnosticBody.appendParagraph('Duplicate image numbers in folder')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    diagnosticBody.appendParagraph(runState.duplicateFolderNumbers.join(', '));
  }

  if (runState.missingNumbers && runState.missingNumbers.length) {
    diagnosticBody.appendParagraph('Missing image numbers')
        .setHeading(DocumentApp.ParagraphHeading.HEADING1);
    diagnosticBody.appendParagraph(runState.missingNumbers.join(', '));
  }

  diagnosticDocument.saveAndClose();
  return diagnosticDocument.getUrl();
}

function removeManagedImageParagraphsBatch_(body, batchLimits, runState) {
  let removedCount = 0;

  for (let index = body.getNumChildren() - 1; index >= 0; index -= 1) {
    const pauseReason = getRemovalPauseReason_(removedCount, batchLimits);
    if (pauseReason) {
      return {
        removedCount: removedCount,
        complete: false,
        limitReached: pauseReason === 'limit',
      };
    }

    const child = body.getChild(index);
    if (!getManagedImageInfoFromBodyChild_(child)) {
      continue;
    }

    body.removeChild(child);
    removedCount += 1;
    runState.removedCount += 1;
    runState.message = 'Removing previously managed images... ' + removedCount + ' removed this batch.';
    saveProgressSnapshot_(runState);
  }

  return {
    removedCount: removedCount,
    complete: true,
    limitReached: false,
  };
}

function getRemovalPauseReason_(removedCount, batchLimits) {
  if (Date.now() > batchLimits.hardDeadlineMs) {
    return 'limit';
  }

  if (
    removedCount >= batchLimits.maxRemovalsPerBatch ||
    Date.now() > batchLimits.softDeadlineMs
  ) {
    return 'batch';
  }

  return '';
}

function collectCaptionEntries_(body) {
  const entries = [];
  const seenNumbers = {};
  const duplicateNumbers = [];

  for (let index = 0; index < body.getNumChildren(); index += 1) {
    const child = body.getChild(index);
    const text = getBodyChildText_(child);

    if (!text) {
      continue;
    }

    const match = text.match(PHOTO_REPORT_CONFIG.captionPattern);
    if (!match) {
      continue;
    }

    const number = Number(match[1]);
    if (seenNumbers[number]) {
      duplicateNumbers.push(number);
    } else {
      seenNumbers[number] = true;
    }

    entries.push({
      childIndex: index,
      number: number,
      caption: match[2].trim(),
    });
  }

  return {
    entries: entries,
    duplicateNumbers: uniqueNumbers_(duplicateNumbers),
  };
}

function validateCaptionEntries_(captionScan) {
  if (!captionScan.entries.length) {
    throw new Error(
        'No caption lines were found. Use plain text lines like "1 - Front view of house" in the body of the document.',
    );
  }

  if (captionScan.duplicateNumbers.length) {
    throw new Error(
        'Duplicate caption numbers in document: ' +
        captionScan.duplicateNumbers.join(', ') +
        '. Fix the duplicate caption lines, then run the report again.',
    );
  }
}

function getBodyChildText_(child) {
  if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
    return child.asParagraph().getText().trim();
  }

  if (child.getType() === DocumentApp.ElementType.LIST_ITEM) {
    return child.asListItem().getText().trim();
  }

  return '';
}

function collectFolderImages_(folder) {
  const imageRecords = [];
  const duplicateNumbers = [];
  const iterator = folder.getFiles();

  while (iterator.hasNext()) {
    const file = iterator.next();

    if (!isSupportedImageFile_(file)) {
      continue;
    }

    const number = extractTrailingNumber_(file.getName());
    if (number === null) {
      continue;
    }

    imageRecords.push({
      number: number,
      file: file,
      preferredName: isPreferredImageName_(file.getName(), number),
      updatedAtMs: file.getLastUpdated().getTime(),
    });
  }

  imageRecords.sort(function(left, right) {
    return (
      left.number - right.number ||
      Number(right.preferredName) - Number(left.preferredName) ||
      right.updatedAtMs - left.updatedAtMs ||
      left.file.getName().localeCompare(right.file.getName()) ||
      left.file.getId().localeCompare(right.file.getId())
    );
  });

  const byNumber = {};

  for (const record of imageRecords) {
    if (byNumber[record.number]) {
      duplicateNumbers.push(record.number);
      continue;
    }

    byNumber[record.number] = record;
  }

  return {
    byNumber: byNumber,
    duplicateNumbers: uniqueNumbers_(duplicateNumbers),
  };
}

function isSupportedImageFile_(file) {
  const mimeType = file.getMimeType();
  return /^image\//i.test(mimeType);
}

function isPreferredImageName_(fileName, number) {
  const escapedNumber = String(number).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const exactNamePattern = new RegExp('^Image\\s*' + escapedNumber + '(\\.[^.]+)?$', 'i');
  return exactNamePattern.test(fileName);
}

function extractTrailingNumber_(fileName) {
  const baseName = fileName.replace(/\.[^.]+$/, '');
  const match = baseName.match(/(\d+)(?!.*\d)/);
  return match ? Number(match[1]) : null;
}

function hasManagedImageAbove_(body, childIndex, number) {
  if (childIndex === 0) {
    return false;
  }

  const managedInfo = getManagedImageInfoFromBodyChild_(body.getChild(childIndex - 1));
  return Boolean(managedInfo && managedInfo.number === number);
}

function insertManagedImageAboveCaption_(body, childIndex, number, file) {
  const paragraph = body.insertParagraph(childIndex, '');
  paragraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  paragraph.setSpacingAfter(6);

  const image = paragraph.insertInlineImage(0, file.getBlob());
  image.setAltTitle('Photo Report Image ' + number);
  image.setAltDescription(buildManagedMarker_(number, file.getId()));
  resizeInlineImage_(image, PHOTO_REPORT_CONFIG.maxImageWidthPx);
}

function resizeInlineImage_(image, maxWidthPx) {
  const currentWidth = image.getWidth();
  const currentHeight = image.getHeight();

  if (!currentWidth || !currentHeight || currentWidth <= maxWidthPx) {
    return;
  }

  const scale = maxWidthPx / currentWidth;
  image.setWidth(Math.round(currentWidth * scale));
  image.setHeight(Math.round(currentHeight * scale));
}

function getManagedImageInfoFromBodyChild_(child) {
  if (!child || child.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    return null;
  }

  const paragraph = child.asParagraph();
  let marker = null;

  for (let index = 0; index < paragraph.getNumChildren(); index += 1) {
    const paragraphChild = paragraph.getChild(index);

    if (paragraphChild.getType() === DocumentApp.ElementType.TEXT) {
      const text = paragraphChild.asText().getText();
      if (text.trim()) {
        return null;
      }
      continue;
    }

    if (paragraphChild.getType() !== DocumentApp.ElementType.INLINE_IMAGE) {
      return null;
    }

    const image = paragraphChild.asInlineImage();
    marker = parseManagedMarker_(image.getAltDescription()) || marker;
  }

  return marker;
}

function buildManagedMarker_(number, fileId) {
  return [
    PHOTO_REPORT_CONFIG.markerPrefix,
    String(number),
    fileId,
  ].join('|');
}

function parseManagedMarker_(text) {
  if (!text) {
    return null;
  }

  const parts = String(text).split('|');
  if (parts.length < 3 || parts[0] !== PHOTO_REPORT_CONFIG.markerPrefix) {
    return null;
  }

  return {
    number: Number(parts[1]),
    fileId: parts[2],
  };
}

function getConfiguredFolder_() {
  const folderId = PropertiesService.getDocumentProperties().getProperty(
      PHOTO_REPORT_CONFIG.folderPropertyKey,
  );

  if (!folderId) {
    throw new Error(
        'No Drive folder is configured yet. Use Photo Report > Set Image Folder first.',
    );
  }

  try {
    return DriveApp.getFolderById(folderId);
  } catch (error) {
    throw new Error(
        'The saved Drive folder could not be opened. Use Photo Report > Set Image Folder again.',
    );
  }
}

function getConfiguredFolderInfo_() {
  const documentProperties = PropertiesService.getDocumentProperties();
  const folderId = PropertiesService.getDocumentProperties().getProperty(
      PHOTO_REPORT_CONFIG.folderPropertyKey,
  );
  const savedFolderName = String(
      documentProperties.getProperty(PHOTO_REPORT_CONFIG.folderNamePropertyKey) || '',
  );

  if (!folderId) {
    return null;
  }

  try {
    return buildFolderInfo_(DriveApp.getFolderById(folderId));
  } catch (error) {
    return {
      id: folderId,
      name: savedFolderName || 'Configured folder unavailable',
      invalid: true,
    };
  }
}

function buildFolderInfo_(folder) {
  return {
    id: folder.getId(),
    name: folder.getName(),
  };
}

function getPickerConfig_() {
  const userProperties = PropertiesService.getUserProperties();
  const userDeveloperKey = String(
      userProperties.getProperty(PHOTO_REPORT_CONFIG.userPickerKeyPropertyKey) || '',
  ).trim();
  const userCloudProjectNumber = String(
      userProperties.getProperty(PHOTO_REPORT_CONFIG.userPickerProjectPropertyKey) || '',
  ).trim();

  if (userDeveloperKey && userCloudProjectNumber) {
    return {
      developerKey: userDeveloperKey,
      cloudProjectNumber: userCloudProjectNumber,
      source: 'user',
    };
  }

  const defaultDeveloperKey = String(PHOTO_REPORT_CONFIG.defaultPickerDeveloperKey || '').trim();
  const defaultCloudProjectNumber = String(
      PHOTO_REPORT_CONFIG.defaultPickerCloudProjectNumber || '',
  ).trim();

  if (defaultDeveloperKey && defaultCloudProjectNumber) {
    return {
      developerKey: defaultDeveloperKey,
      cloudProjectNumber: defaultCloudProjectNumber,
      source: 'repo',
    };
  }

  return null;
}

function buildPickerClientState_() {
  const pickerConfig = getPickerConfig_();

  return {
    configured: Boolean(pickerConfig),
    source: pickerConfig ? pickerConfig.source : '',
    developerKey: pickerConfig ? pickerConfig.developerKey : '',
    cloudProjectNumber: pickerConfig ? pickerConfig.cloudProjectNumber : '',
    hasUserCredentials: hasUserPickerCredentials_(),
  };
}

function hasUserPickerCredentials_() {
  const userProperties = PropertiesService.getUserProperties();
  return Boolean(
      userProperties.getProperty(PHOTO_REPORT_CONFIG.userPickerKeyPropertyKey) &&
      userProperties.getProperty(PHOTO_REPORT_CONFIG.userPickerProjectPropertyKey),
  );
}

function saveConfiguredFolderById_(folderId, errorMessage, preferredFolderName) {
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (error) {
    throw new Error(errorMessage);
  }

  const documentProperties = PropertiesService.getDocumentProperties();
  documentProperties.setProperty(
      PHOTO_REPORT_CONFIG.folderPropertyKey,
      folder.getId(),
  );
  documentProperties.setProperty(
      PHOTO_REPORT_CONFIG.folderNamePropertyKey,
      String(preferredFolderName || folder.getName()),
  );
  clearStoredRunState_();

  return folder;
}

function buildSidebarModelAfterFolderSave_(folder) {
  return {
    requestedAction: '',
    folder: buildFolderInfo_(folder),
    picker: buildPickerClientState_(),
    run: buildClientRunState_(null),
  };
}

function extractDriveFolderId_(value) {
  const input = String(value || '').trim();
  if (!input) {
    throw new Error('Paste a Drive folder URL or folder ID.');
  }

  const folderPathMatch = input.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  if (folderPathMatch) {
    return folderPathMatch[1];
  }

  const idParamMatch = input.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (idParamMatch) {
    return idParamMatch[1];
  }

  if (/^[a-zA-Z0-9_-]{10,}$/.test(input)) {
    return input;
  }

  throw new Error('Could not find a Drive folder ID in that value.');
}

function getPhotoReportAction_(actionKey) {
  if (!PHOTO_REPORT_ACTIONS[actionKey]) {
    throw new Error('Unknown Photo Report action: ' + actionKey);
  }

  return PHOTO_REPORT_ACTIONS[actionKey];
}

function getStoredRunState_() {
  const rawState = PropertiesService.getDocumentProperties().getProperty(
      PHOTO_REPORT_CONFIG.runStatePropertyKey,
  );

  if (!rawState) {
    return null;
  }

  try {
    return JSON.parse(rawState);
  } catch (error) {
    clearStoredRunState_();
    return null;
  }
}

function saveRunState_(runState) {
  normalizeRunStateForStorage_(runState);
  PropertiesService.getDocumentProperties().setProperty(
      PHOTO_REPORT_CONFIG.runStatePropertyKey,
      JSON.stringify(runState),
  );
}

function clearStoredRunState_() {
  PropertiesService.getDocumentProperties().deleteProperty(
      PHOTO_REPORT_CONFIG.runStatePropertyKey,
  );
}

function normalizeRunStateForStorage_(runState) {
  runState.consoleLines = ensureRecentStringArray_(runState.consoleLines);
  runState.errors = ensureRecentStringArray_(runState.errors, PHOTO_REPORT_CONFIG.maxStoredErrors);
  runState.errorDetails = ensureRecentObjectArray_(
      runState.errorDetails,
      PHOTO_REPORT_CONFIG.maxStoredErrorDetails,
  );
}

function saveProgressSnapshot_(runState) {
  runState.updatedAt = new Date().toISOString();
  saveRunState_(runState);
}

function buildClientRunState_(runState) {
  if (!runState) {
    return {
      action: '',
      actionLabel: '',
      status: 'idle',
      phase: 'idle',
      folderName: '',
      totalCount: 0,
      processedCount: 0,
      remainingCount: 0,
      progressPercent: 0,
      insertedCount: 0,
      removedCount: 0,
      skippedCount: 0,
      missingNumbers: [],
      duplicateFolderNumbers: [],
      errors: [],
      consoleLines: [],
      limitReached: false,
      needsContinuation: false,
      canResume: false,
      message: 'Ready to choose a folder and start a report batch.',
      startedAt: '',
      updatedAt: '',
      diagnosticLogUrl: '',
    };
  }

  const processedCount = ensureNumberArray_(runState.completedNumbers).length;
  const totalCount = Number(runState.totalCount || 0);
  const remainingCount = Math.max(totalCount - processedCount, 0);
  const progressPercent = totalCount ?
    Math.max(0, Math.min(100, Math.round((processedCount / totalCount) * 100))) :
    0;

  return {
    action: String(runState.action || ''),
    actionLabel: String(runState.actionLabel || ''),
    status: String(runState.status || 'idle'),
    phase: String(runState.phase || 'idle'),
    folderName: String(runState.folderName || ''),
    totalCount: totalCount,
    processedCount: processedCount,
    remainingCount: remainingCount,
    progressPercent: progressPercent,
    insertedCount: Number(runState.insertedCount || 0),
    removedCount: Number(runState.removedCount || 0),
    skippedCount: Number(runState.skippedCount || 0),
    missingNumbers: sortedNumberCopy_(runState.missingNumbers),
    duplicateFolderNumbers: sortedNumberCopy_(runState.duplicateFolderNumbers),
    errors: ensureStringArray_(runState.errors),
    consoleLines: ensureRecentStringArray_(runState.consoleLines),
    limitReached: Boolean(runState.limitReached),
    needsContinuation: Boolean(runState.needsContinuation),
    canResume: Boolean(runState.needsContinuation && runState.limitReached),
    message: String(runState.message || ''),
    startedAt: String(runState.startedAt || ''),
    updatedAt: String(runState.updatedAt || ''),
    diagnosticLogUrl: String(runState.diagnosticLogUrl || ''),
  };
}

function buildRunPhaseMessage_(runState) {
  if (runState.phase === 'remove') {
    return 'Removing previously managed images...';
  }

  return buildProcessingMessage_(runState);
}

function buildProcessingMessage_(runState) {
  const totalCount = Number(runState.totalCount || 0);
  if (!totalCount) {
    return 'Preparing report batch...';
  }

  const nextIndex = Math.min(runState.completedNumbers.length + 1, totalCount);
  return 'Processing image ' + nextIndex + ' of ' + totalCount + '...';
}

function buildPauseMessage_(runState, fallbackMessage) {
  if (runState.phase === 'remove') {
    return (
      fallbackMessage +
      ' Removed ' +
      String(runState.removedCount || 0) +
      ' managed image paragraph(s) so far.'
    );
  }

  return fallbackMessage + ' ' + buildProcessingMessage_(runState);
}

function buildCompletionMessage_(runState) {
  const missingCount = ensureNumberArray_(runState.missingNumbers).length;
  const errorCount = ensureStringArray_(runState.errors).length;

  return [
    'Complete.',
    'Inserted ' + String(runState.insertedCount || 0) + '.',
    'Skipped ' + String(runState.skippedCount || 0) + '.',
    'Missing ' + String(missingCount) + '.',
    'Errors ' + String(errorCount) + '.',
  ].join(' ');
}

function formatDiagnosticError_(detail) {
  return [
    'Caption ' + String(detail.number || 'n/a'),
    'file "' + String(detail.fileName || 'Unknown file') + '"',
    'ID ' + String(detail.fileId || 'n/a'),
    String(detail.message || 'Unknown error'),
  ].join(' | ');
}

function ensureNumberArray_(values) {
  if (!Array.isArray(values)) {
    return [];
  }

  return values.map(function(value) {
    return Number(value);
  }).filter(function(value) {
    return !Number.isNaN(value);
  });
}

function ensureStringArray_(values) {
  if (!Array.isArray(values)) {
    return [];
  }

  return values.map(function(value) {
    return String(value);
  });
}

function ensureRecentStringArray_(values, maxItems) {
  const limit = Number(maxItems || PHOTO_REPORT_CONFIG.maxConsoleLines);
  const normalized = ensureStringArray_(values);
  return normalized.slice(Math.max(normalized.length - limit, 0));
}

function ensureObjectArray_(values) {
  if (!Array.isArray(values)) {
    return [];
  }

  return values.filter(function(value) {
    return value && typeof value === 'object';
  });
}

function ensureRecentObjectArray_(values, maxItems) {
  const normalized = ensureObjectArray_(values);
  const limit = Number(maxItems || PHOTO_REPORT_CONFIG.maxStoredErrorDetails);
  return normalized.slice(Math.max(normalized.length - limit, 0));
}

function filterNumbersToKnownSet_(values, knownNumbers) {
  const knownSet = {};
  for (const number of knownNumbers) {
    knownSet[number] = true;
  }

  const filtered = [];
  for (const value of ensureNumberArray_(values)) {
    if (!knownSet[value]) {
      continue;
    }

    pushUniqueNumber_(filtered, value);
  }

  return filtered;
}

function sortedNumberCopy_(values) {
  return sortNumbers_(ensureNumberArray_(values).slice());
}

function appendRunLog_(runState, message) {
  const trimmedMessage = String(message || '').trim();
  if (!trimmedMessage) {
    return;
  }

  runState.consoleLines = ensureRecentStringArray_(runState.consoleLines);
  const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'HH:mm:ss',
  );
  const nextLine = '[' + timestamp + '] ' + trimmedMessage;
  const lastLine = runState.consoleLines.length ?
    runState.consoleLines[runState.consoleLines.length - 1] :
    '';

  if (lastLine === nextLine) {
    return;
  }

  runState.consoleLines.push(nextLine);
  runState.consoleLines = ensureRecentStringArray_(runState.consoleLines);
}

function uniqueNumbers_(values) {
  const seen = {};
  const uniqueValues = [];

  for (const value of values) {
    if (seen[value]) {
      continue;
    }

    seen[value] = true;
    uniqueValues.push(value);
  }

  return sortNumbers_(uniqueValues);
}

function pushUniqueNumber_(values, value) {
  if (values.indexOf(value) !== -1) {
    return;
  }

  values.push(value);
}

function sortNumbers_(values) {
  return values.sort(function(left, right) {
    return left - right;
  });
}
