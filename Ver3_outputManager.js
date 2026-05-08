/****************************************************
 * Ver3_outputManager.js — JREC-01 月次出力フォルダ共通管理
 *
 * 目的:
 *   申請書 / 施術録 の「月次出力」と「再生成旧版退避」を同一ルールへ統一し、
 *   テンプレート正本と成果物の保存場所を混同しないようにする。
 *
 * フォルダ構造:
 *   JREC-01_月次出力/
 *     YYYY-MM/
 *       01_申請書/
 *       02_施術録/
 *       90_再生成旧版/
 ****************************************************/

var V3OUT = V3OUT || {};

V3OUT.ROOT_FOLDER_NAME = 'JREC-01_月次出力';
V3OUT.ARCHIVE_FOLDER_NAME = '90_再生成旧版';

V3OUT.DOC_TYPE_FOLDER = {
  application: '01_申請書',
  shuroku: '02_施術録'
};

V3OUT.SETTING_KEY = {
  outputFolderId: '出力フォルダID'
};

V3OUT.normalizeYearMonth_ = function(yearMonth) {
  var ym = String(yearMonth || '').trim();
  if (!/^\d{4}-\d{2}$/.test(ym)) {
    throw new Error('年月は YYYY-MM 形式で指定してください: ' + ym);
  }
  return ym;
};

V3OUT.getSettingValue_ = function(ss, keyName) {
  if (!ss || !keyName) return '';
  var sh = ss.getSheetByName('設定');
  if (!sh || sh.getLastRow() < 1) return '';

  var values = sh.getDataRange().getValues();
  for (var r = 0; r < values.length; r++) {
    if (String(values[r][0] || '').trim() === keyName) {
      return String(values[r][1] || '').trim();
    }
  }
  return '';
};

V3OUT.getFolderByIdSafe_ = function(folderId) {
  var id = String(folderId || '').trim();
  if (!id) return null;
  try {
    return DriveApp.getFolderById(id);
  } catch (e) {
    Logger.log('[V3OUT] フォルダID無効: ' + id + ' - ' + e.message);
    return null;
  }
};

V3OUT.getSpreadsheetParentFolderSafe_ = function(ss) {
  if (!ss) return null;
  try {
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    return parents.hasNext() ? parents.next() : null;
  } catch (e) {
    Logger.log('[V3OUT] スプレッドシート親フォルダ取得失敗: ' + e.message);
    return null;
  }
};

V3OUT.getOrCreateChildFolder_ = function(parentFolder, childName) {
  if (!parentFolder) {
    throw new Error('親フォルダが未指定です: ' + childName);
  }
  var name = String(childName || '').trim();
  if (!name) {
    throw new Error('子フォルダ名が未指定です');
  }
  var iter = parentFolder.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : parentFolder.createFolder(name);
};

/**
 * 月次出力ルートを取得する。
 * 優先順:
 *   1. 設定!A:B の「出力フォルダID」
 *   2. 呼び出し側の legacyFallbackFolderId
 *   3. このスプレッドシートの親フォルダ
 *   4. Drive ルート
 *
 * 取得したフォルダが `JREC-01_月次出力` 自体でなければ、
 * その配下に `JREC-01_月次出力` を作って成果物棚とする。
 */
V3OUT.getOrCreateMonthlyOutputRootFolder_ = function(ss, legacyFallbackFolderId) {
  var configuredFolderId = V3OUT.getSettingValue_(ss, V3OUT.SETTING_KEY.outputFolderId);
  var baseFolder =
    V3OUT.getFolderByIdSafe_(configuredFolderId) ||
    V3OUT.getFolderByIdSafe_(legacyFallbackFolderId) ||
    V3OUT.getSpreadsheetParentFolderSafe_(ss) ||
    DriveApp.getRootFolder();

  if (baseFolder.getName() === V3OUT.ROOT_FOLDER_NAME) {
    return baseFolder;
  }
  return V3OUT.getOrCreateChildFolder_(baseFolder, V3OUT.ROOT_FOLDER_NAME);
};

V3OUT.getOrCreateMonthlyOutputFolder_ = function(ss, yearMonth, legacyFallbackFolderId) {
  var ym = V3OUT.normalizeYearMonth_(yearMonth);
  var root = V3OUT.getOrCreateMonthlyOutputRootFolder_(ss, legacyFallbackFolderId);
  return V3OUT.getOrCreateChildFolder_(root, ym);
};

V3OUT.getOrCreateDocTypeFolder_ = function(ss, yearMonth, docTypeKey, legacyFallbackFolderId) {
  var monthFolder = V3OUT.getOrCreateMonthlyOutputFolder_(ss, yearMonth, legacyFallbackFolderId);
  var childName = V3OUT.DOC_TYPE_FOLDER[String(docTypeKey || '').trim()];
  if (!childName) {
    throw new Error('未知の出力種別です: ' + docTypeKey);
  }
  return V3OUT.getOrCreateChildFolder_(monthFolder, childName);
};

V3OUT.getOrCreateArchiveFolder_ = function(ss, yearMonth, legacyFallbackFolderId) {
  var monthFolder = V3OUT.getOrCreateMonthlyOutputFolder_(ss, yearMonth, legacyFallbackFolderId);
  return V3OUT.getOrCreateChildFolder_(monthFolder, V3OUT.ARCHIVE_FOLDER_NAME);
};

V3OUT.buildArchiveName_ = function(fileName, fileId) {
  var name = String(fileName || '').trim();
  if (!name) name = 'unnamed';

  var dot = name.lastIndexOf('.');
  var base = dot > 0 ? name.substring(0, dot) : name;
  var ext = dot > 0 ? name.substring(dot) : '';
  var idSuffix = String(fileId || '').trim();
  if (idSuffix.length > 8) idSuffix = idSuffix.substring(0, 8);
  var timestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyyMMdd-HHmmss');

  return base + '_旧_' + timestamp + (idSuffix ? '_' + idSuffix : '') + ext;
};

V3OUT.archiveFileObject_ = function(fileObj, archiveFolder) {
  if (!fileObj || !archiveFolder) return;
  var archivedName = V3OUT.buildArchiveName_(fileObj.getName(), fileObj.getId());
  fileObj.setName(archivedName);
  fileObj.moveTo(archiveFolder);
  Logger.log('[V3OUT] 旧版退避: ' + archivedName + ' -> ' + archiveFolder.getName());
};

V3OUT.archiveFilesByExactName_ = function(sourceFolder, archiveFolder, fileName) {
  if (!sourceFolder || !archiveFolder) return 0;
  var targetName = String(fileName || '').trim();
  if (!targetName) return 0;

  var count = 0;
  var iter = sourceFolder.getFilesByName(targetName);
  while (iter.hasNext()) {
    V3OUT.archiveFileObject_(iter.next(), archiveFolder);
    count++;
  }
  return count;
};

V3OUT.archiveFilesByPrefix_ = function(sourceFolder, archiveFolder, namePrefix) {
  if (!sourceFolder || !archiveFolder) return 0;
  var prefix = String(namePrefix || '').trim();
  if (!prefix) return 0;

  var count = 0;
  var iter = sourceFolder.getFiles();
  while (iter.hasNext()) {
    var f = iter.next();
    if (f.getName().indexOf(prefix) === 0) {
      V3OUT.archiveFileObject_(f, archiveFolder);
      count++;
    }
  }
  return count;
};
