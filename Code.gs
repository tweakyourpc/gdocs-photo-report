'use strict';

const PHOTO_REPORT_CONFIG = Object.freeze({
  folderPropertyKey: 'PHOTO_REPORT_FOLDER_ID',
  markerPrefix: 'PHOTO_REPORT',
  captionPattern: /^\s*(\d+)\s*[-.):]\s*(.+\S)\s*$/,
  maxImageWidthPx: 520,
  pickerDeveloperKey: 'REPLACE_WITH_PICKER_API_KEY',
  pickerCloudProjectNumber: 'REPLACE_WITH_CLOUD_PROJECT_NUMBER',
  pickerDialogWidth: 600,
  pickerDialogHeight: 425,
});

function onOpen() {
  DocumentApp.getUi()
    .createMenu('Photo Report')
    .addItem('Set Image Folder', 'showPicker')
    .addSeparator()
    .addItem('Insert Missing Images', 'insertMissingImages')
    .addItem('Rebuild Images', 'rebuildImages')
    .addToUi();
}

function showPicker() {
  if (!isPickerConfigured_()) {
    setImageFolderPrompt_();
    return;
  }

  const pickerConfig = getPickerConfig_();
  const template = HtmlService.createTemplateFromFile('Picker');
  template.pickerDeveloperKey = pickerConfig.developerKey;
  template.pickerCloudProjectNumber = pickerConfig.cloudProjectNumber;

  const html = template.evaluate()
    .setWidth(PHOTO_REPORT_CONFIG.pickerDialogWidth)
    .setHeight(PHOTO_REPORT_CONFIG.pickerDialogHeight)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);

  DocumentApp.getUi().showModalDialog(html, 'Select image folder');
}

function getOAuthToken() {
  return ScriptApp.getOAuthToken();
}

function handlePickerResponse(id) {
  if (!id) {
    throw new Error('No Drive folder was selected.');
  }

  const folder = saveConfiguredFolderById_(String(id), 'The selected Drive folder could not be opened.');

  return {
    id: folder.getId(),
    name: folder.getName(),
  };
}

function setImageFolderPrompt_() {
  const ui = DocumentApp.getUi();
  const response = ui.prompt(
    'Set Google Drive folder',
    'Google Picker is not configured for this copy of the script yet. Paste the Drive folder URL or folder ID that holds Image 1, Image 2, etc.',
    ui.ButtonSet.OK_CANCEL
  );

  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }

  try {
    const folderId = extractDriveFolderId_(response.getResponseText());
    const folder = saveConfiguredFolderById_(
      folderId,
      'Could not save that Drive folder. Make sure the folder URL or ID is valid and accessible.'
    );

    ui.alert(
      'Folder saved',
      'Using "' + folder.getName() + '" for future photo imports in this document.',
      ui.ButtonSet.OK
    );
  } catch (error) {
    ui.alert('Could not save folder.\n\n' + error.message);
  }
}

function insertMissingImages() {
  runPhotoReportAction_('Insert Missing Images', {
    rebuild: false,
  });
}

function rebuildImages() {
  runPhotoReportAction_('Rebuild Images', {
    rebuild: true,
  });
}

function runPhotoReportAction_(title, options) {
  const ui = DocumentApp.getUi();

  try {
    const summary = syncPhotoReport_(options);
    ui.alert(title + ' complete.\n\n' + formatSummary_(summary));
  } catch (error) {
    ui.alert(title + ' failed.\n\n' + error.message);
  }
}

function syncPhotoReport_(options) {
  const startTime = Date.now();
  const MAX_EXECUTION_MS = 270000;
  const folder = getConfiguredFolder_();
  const document = DocumentApp.getActiveDocument();
  const body = document.getBody();
  const summary = {
    folderName: folder.getName(),
    captionsFound: 0,
    insertedCount: 0,
    removedCount: 0,
    skippedCount: 0,
    missingNumbers: [],
    duplicateFolderNumbers: [],
    limitReached: false,
    errors: [],
  };

  const captionScan = collectCaptionEntries_(body);
  const captions = captionScan.entries;
  summary.captionsFound = captions.length;
  validateCaptionEntries_(captionScan);

  if (options.rebuild) {
    const removalSummary = removeManagedImageParagraphs_(body, startTime, MAX_EXECUTION_MS);
    summary.removedCount = removalSummary.removedCount;

    if (removalSummary.limitReached) {
      summary.limitReached = true;
      return summary;
    }
  }

  const folderImages = collectFolderImages_(folder);
  summary.duplicateFolderNumbers = folderImages.duplicateNumbers;

  const captionsDescending = captions
    .slice()
    .sort(function (left, right) {
      return right.childIndex - left.childIndex;
    });

  for (const captionEntry of captionsDescending) {
    if (hasReachedExecutionLimit_(startTime, MAX_EXECUTION_MS)) {
      summary.limitReached = true;
      break;
    }

    const imageRecord = folderImages.byNumber[captionEntry.number];

    if (!imageRecord) {
      pushUniqueNumber_(summary.missingNumbers, captionEntry.number);
      continue;
    }

    if (!options.rebuild && hasManagedImageAbove_(body, captionEntry.childIndex, captionEntry.number)) {
      summary.skippedCount += 1;
      continue;
    }

    try {
      insertManagedImageAboveCaption_(body, captionEntry.childIndex, captionEntry.number, imageRecord.file);
      summary.insertedCount += 1;
    } catch (error) {
      summary.errors.push('Image ' + captionEntry.number + ': ' + error.message);
    }

    if (hasReachedExecutionLimit_(startTime, MAX_EXECUTION_MS)) {
      summary.limitReached = true;
      break;
    }
  }

  sortNumbers_(summary.missingNumbers);
  return summary;
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
      'No caption lines were found. Use plain text lines like "1 - Front view of house" in the body of the document.'
    );
  }

  if (captionScan.duplicateNumbers.length) {
    throw new Error(
      'Duplicate caption numbers in document: ' +
        captionScan.duplicateNumbers.join(', ') +
        '. Fix the duplicate caption lines, then run the report again.'
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

  imageRecords.sort(function (left, right) {
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

function removeManagedImageParagraphs_(body, startTime, maxExecutionMs) {
  let removedCount = 0;

  for (let index = body.getNumChildren() - 1; index >= 0; index -= 1) {
    if (hasReachedExecutionLimit_(startTime, maxExecutionMs)) {
      return {
        removedCount: removedCount,
        limitReached: true,
      };
    }

    const child = body.getChild(index);
    if (!getManagedImageInfoFromBodyChild_(child)) {
      continue;
    }

    body.removeChild(child);
    removedCount += 1;

    if (hasReachedExecutionLimit_(startTime, maxExecutionMs)) {
      return {
        removedCount: removedCount,
        limitReached: true,
      };
    }
  }

  return {
    removedCount: removedCount,
    limitReached: false,
  };
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

function hasReachedExecutionLimit_(startTime, maxExecutionMs) {
  return Date.now() - startTime > maxExecutionMs;
}

function getConfiguredFolder_() {
  const folderId = PropertiesService.getDocumentProperties().getProperty(
    PHOTO_REPORT_CONFIG.folderPropertyKey
  );

  if (!folderId) {
    throw new Error(
      'No Drive folder is configured yet. Use Photo Report > Set Image Folder first.'
    );
  }

  try {
    return DriveApp.getFolderById(folderId);
  } catch (error) {
    throw new Error(
      'The saved Drive folder could not be opened. Use Photo Report > Set Image Folder again.'
    );
  }
}

function getPickerConfig_() {
  return {
    developerKey: PHOTO_REPORT_CONFIG.pickerDeveloperKey,
    cloudProjectNumber: PHOTO_REPORT_CONFIG.pickerCloudProjectNumber,
  };
}

function isPickerConfigured_() {
  return Boolean(
    PHOTO_REPORT_CONFIG.pickerDeveloperKey &&
      PHOTO_REPORT_CONFIG.pickerDeveloperKey !== 'REPLACE_WITH_PICKER_API_KEY' &&
      PHOTO_REPORT_CONFIG.pickerCloudProjectNumber &&
      PHOTO_REPORT_CONFIG.pickerCloudProjectNumber !== 'REPLACE_WITH_CLOUD_PROJECT_NUMBER'
  );
}

function saveConfiguredFolderById_(folderId, errorMessage) {
  let folder;
  try {
    folder = DriveApp.getFolderById(folderId);
  } catch (error) {
    throw new Error(errorMessage);
  }

  PropertiesService.getDocumentProperties().setProperty(
    PHOTO_REPORT_CONFIG.folderPropertyKey,
    folder.getId()
  );

  return folder;
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
  return values.sort(function (left, right) {
    return left - right;
  });
}

function formatSummary_(summary) {
  const lines = [
    'Folder: ' + summary.folderName,
    'Caption lines found: ' + summary.captionsFound,
    'Images inserted: ' + summary.insertedCount,
  ];

  if (summary.limitReached) {
    lines.unshift(
      "⚠️ EXECUTION TIME LIMIT REACHED. The script paused to prevent timing out. Please run 'Insert Missing Images' again to continue processing the remaining photos."
    );
  }

  if (summary.removedCount) {
    lines.push('Old managed images removed: ' + summary.removedCount);
  }

  if (summary.skippedCount) {
    lines.push('Captions already matched to an image: ' + summary.skippedCount);
  }

  if (summary.missingNumbers.length) {
    lines.push('Missing image numbers: ' + summary.missingNumbers.join(', '));
  }

  if (summary.duplicateFolderNumbers.length) {
    lines.push('Duplicate image numbers in folder: ' + summary.duplicateFolderNumbers.join(', '));
  }

  if (summary.errors.length) {
    lines.push('Problems: ' + summary.errors.join(' | '));
  }

  return lines.join('\n');
}
