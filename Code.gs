/**
 * 避難所サイネージ更新監視 PoC - Google Apps Script版
 *
 * 目的:
 * - 複数自治体ページを1時間ごとに監視
 * - 前回との差分から避難生活に関係する更新候補だけ抽出
 * - 原文からのみドラフトを作成（生成AIなし）
 * - A4横1ページPDFをGoogle Sheets経由で生成
 * - メール / Slack通知
 * - Web Appで確認・修正・承認・対象外
 *
 * 初回:
 * 1. setupPoc() を実行して承認
 * 2. Webアプリとしてデプロイ（実行ユーザー: 自分、アクセス: Googleアカウントを持つ全員）
 * 3. showPocInfo() を実行し、レビューURLを確認
 */

const APP = {
  NAME: '避難所サイネージ更新監視PoC',
  TZ: 'Asia/Tokyo',
  DB_NAME: '避難所サイネージPoC_DB',
  ROOT_FOLDER: '避難所サイネージPoC',
  PDF_FOLDER: 'PDF',
  SNAPSHOT_FOLDER: 'Snapshots',
  SOURCE_SHEET: 'Sources',
  CANDIDATE_SHEET: 'Candidates',
  LOG_SHEET: 'Logs',
  MONITOR_HANDLER: 'monitorAllSources',
  MONITOR_HOURS: 1,
  MAX_DIFF_LINES: 80,
  MAX_LINKS: 180,
  MAX_BULLETS: 7,
  MAX_BLOCK_CHARS: 170,
  PDF_PUBLIC_LINK: true,
  ERROR_NOTIFY_AFTER: 2,
};

const DEFAULT_SOURCES = [
  {
    id: 'yachiyo_bousai',
    name: '八千代市 防災',
    url: 'https://www.city.yachiyo.lg.jp/life/1/9/index.html',
    category: '防災',
  },
  {
    id: 'yachiyo_hinan',
    name: '八千代市 避難所',
    url: 'https://www.city.yachiyo.lg.jp/life/1/9/56/',
    category: '避難所',
  },
  {
    id: 'yachiyo_water',
    name: '八千代市 上水道課',
    url: 'https://www.city.yachiyo.lg.jp/soshiki/92/',
    category: '給水・断水',
  },
  {
    id: 'chiba_bousai',
    name: '千葉県 防災ポータル',
    url: 'https://www.bousai.pref.chiba.lg.jp/portal/X_PUB_VF_Top',
    category: '県防災',
  },
  {
    id: 'yachiyo_top',
    name: '八千代市 トップ',
    url: 'https://www.city.yachiyo.lg.jp/',
    category: '総合',
  },
];

const SOURCE_HEADERS = [
  'id', 'name', 'url', 'category', 'active',
  'last_checked', 'last_success', 'http_status',
  'etag', 'last_modified', 'content_hash',
  'snapshot_file_id', 'links_json',
  'consecutive_errors', 'last_error', 'error_notified_at', 'recovered_at'
];

const CANDIDATE_HEADERS = [
  'id', 'source_id', 'source_name', 'status',
  'detected_at', 'notified_at', 'first_viewed_at', 'reviewed_at',
  'reviewer', 'title', 'source_url', 'diff_summary',
  'old_excerpt', 'new_excerpt', 'draft_json',
  'pdf_file_id', 'pdf_url', 'review_note',
  'old_hash', 'new_hash', 'updated_at'
];

const LOG_HEADERS = ['timestamp', 'level', 'event', 'source_id', 'candidate_id', 'message'];

// 避難生活・公的支援に関係する語。広めに拾い、最終確認は人が行う。
const POSITIVE_KEYWORDS = [
  '給水', '断水', '濁り水', '応急給水', '飲料水', '水道',
  '避難所', '避難場所', '避難',
  '入浴', '風呂', 'シャワー',
  '食料', '食品', '炊き出し', '弁当', '物資', '配布', '支援物資',
  '医療', '診療', '病院', '救護', '薬', '保健',
  '通行止', '通行規制', '交通', '道路', '鉄道', 'バス',
  '停電', '電気', 'ガス', 'ライフライン', '通信',
  '罹災', 'り災', '証明書', '行政手続', '手続き', '申請',
  '支援金', '給付', '補助', '減免', '貸付', '災害援護',
  '住宅', '住まい', '応急仮設', 'みなし仮設', '被災者',
  'ごみ', '災害ごみ', '廃棄物', 'トイレ', '洗濯',
  '充電', '携帯電話', 'Wi-Fi', 'WiFi',
  'ボランティア', '生活再建', '相談窓口'
];

const HIGH_PRIORITY_KEYWORDS = [
  '給水', '断水', '避難所', '入浴',
  '食料', '物資', '医療', '通行止', '停電', 'ガス', '支援金', '給付'
];

const NEGATIVE_KEYWORDS = [
  '採用', '職員募集', '入札', '議会', '選挙', '観光', 'スポーツ',
  '講座', 'イベント', '文化祭', '募集', '広報紙', 'ふるさと納税'
];

const BOILERPLATE_PATTERNS = [
  /^本文へ$/,
  /^メニュー$/,
  /^トップページ$/,
  /^このページを見ている人はこんなページも見ています$/,
  /^ページの先頭へ$/,
  /^サイトマップ$/,
  /^お問い合わせ$/,
  /^Copyright/i,
  /^All Rights Reserved/i,
  /^文字サイズ/,
  /^検索$/,
  /^閉じる$/,
  /^前のページへ$/,
];

/** -------------------- 初期セットアップ -------------------- */

function setupPoc() {
  const props = PropertiesService.getScriptProperties();

  let rootFolderId = props.getProperty('ROOT_FOLDER_ID');
  let pdfFolderId = props.getProperty('PDF_FOLDER_ID');
  let snapshotFolderId = props.getProperty('SNAPSHOT_FOLDER_ID');
  let dbId = props.getProperty('DB_ID');

  if (!rootFolderId) {
    const root = DriveApp.createFolder(APP.ROOT_FOLDER);
    rootFolderId = root.getId();
    props.setProperty('ROOT_FOLDER_ID', rootFolderId);
  }
  const root = DriveApp.getFolderById(rootFolderId);

  if (!pdfFolderId) {
    const folder = root.createFolder(APP.PDF_FOLDER);
    pdfFolderId = folder.getId();
    props.setProperty('PDF_FOLDER_ID', pdfFolderId);
  }
  if (!snapshotFolderId) {
    const folder = root.createFolder(APP.SNAPSHOT_FOLDER);
    snapshotFolderId = folder.getId();
    props.setProperty('SNAPSHOT_FOLDER_ID', snapshotFolderId);
  }

  if (!dbId) {
    const db = SpreadsheetApp.create(APP.DB_NAME);
    dbId = db.getId();
    props.setProperty('DB_ID', dbId);
    moveFileToFolder_(dbId, rootFolderId);
  }

  ensureDbStructure_();
  seedDefaultSources_();

  if (!props.getProperty('REVIEW_TOKEN')) {
    props.setProperty('REVIEW_TOKEN', randomToken_(24));
  }

  if (!props.getProperty('NOTIFY_EMAIL')) {
    const ownerEmail = Session.getEffectiveUser().getEmail();
    if (ownerEmail) props.setProperty('NOTIFY_EMAIL', ownerEmail);
  }

  installMonitorTrigger_();

  // 初回はbaselineとして保存される。既存情報を更新候補にはしない。
  monitorAllSources();

  log_('INFO', 'setup', '', '', '初期セットアップ完了');
  showPocInfo();
}

function ensureDbStructure_() {
  const db = getDb_();
  ensureSheet_(db, APP.SOURCE_SHEET, SOURCE_HEADERS);
  ensureSheet_(db, APP.CANDIDATE_SHEET, CANDIDATE_HEADERS);
  ensureSheet_(db, APP.LOG_SHEET, LOG_HEADERS);
}

function ensureSheet_(db, name, headers) {
  let sh = db.getSheetByName(name);
  if (!sh) {
    sh = db.insertSheet(name);
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8eef7');
  } else {
    const existing = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), headers.length)).getValues()[0];
    headers.forEach((h, i) => {
      if (existing[i] !== h) sh.getRange(1, i + 1).setValue(h);
    });
  }
  return sh;
}

function seedDefaultSources_() {
  const rows = getObjects_(APP.SOURCE_SHEET);
  const ids = new Set(rows.map(r => String(r.id)));
  DEFAULT_SOURCES.forEach(src => {
    if (!ids.has(src.id)) {
      appendObject_(APP.SOURCE_SHEET, SOURCE_HEADERS, {
        ...src,
        active: true,
        consecutive_errors: 0,
      });
    }
  });
}

function installMonitorTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === APP.MONITOR_HANDLER)
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger(APP.MONITOR_HANDLER)
    .timeBased()
    .everyHours(APP.MONITOR_HOURS)
    .create();
}

function showPocInfo() {
  const props = PropertiesService.getScriptProperties();
  const token = props.getProperty('REVIEW_TOKEN') || '';
  const webUrl = ScriptApp.getService().getUrl();
  const reviewUrl = webUrl ? `${webUrl}?token=${encodeURIComponent(token)}` : '(Webアプリをデプロイ後に再実行してください)';
  const info = [
    `DB: https://docs.google.com/spreadsheets/d/${props.getProperty('DB_ID')}/edit`,
    `Drive: https://drive.google.com/drive/folders/${props.getProperty('ROOT_FOLDER_ID')}`,
    `Web App: ${webUrl || '(未デプロイ)'}`,
    `チーム共有URL: ${reviewUrl}`,
    `通知メール: ${props.getProperty('NOTIFY_EMAIL') || '(未設定)'}`,
    `Slack: ${props.getProperty('SLACK_WEBHOOK_URL') ? '設定済み' : '未設定'}`,
  ].join('\n');
  console.log(info);
  Logger.log(info);
  return info;
}

function setNotificationEmailForPoc(email) {
  PropertiesService.getScriptProperties().setProperty('NOTIFY_EMAIL', String(email || '').trim());
  return '通知メールを更新しました。';
}

function setSlackWebhookForPoc(url) {
  const value = String(url || '').trim();
  const props = PropertiesService.getScriptProperties();
  if (value) props.setProperty('SLACK_WEBHOOK_URL', value);
  else props.deleteProperty('SLACK_WEBHOOK_URL');
  return value ? 'Slack Webhookを設定しました。' : 'Slack Webhookを削除しました。';
}

function rotateReviewTokenForPoc() {
  const token = randomToken_(24);
  PropertiesService.getScriptProperties().setProperty('REVIEW_TOKEN', token);
  showPocInfo();
  return token;
}

/** -------------------- 監視 -------------------- */

function monitorAllSources() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;

  try {
    const sources = getObjects_(APP.SOURCE_SHEET).filter(r => truthy_(r.active));
    sources.forEach(source => {
      try {
        monitorOneSource_(source);
      } catch (err) {
        handleSourceError_(source, err);
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function monitorOneSource_(source) {
  const headers = {
    'User-Agent': 'Mozilla/5.0 (compatible; ShelterSignagePoC/1.0; GoogleAppsScript)',
    'Accept-Language': 'ja,en;q=0.7',
  };
  if (source.etag) headers['If-None-Match'] = String(source.etag);
  if (source.last_modified) headers['If-Modified-Since'] = String(source.last_modified);

  const now = nowIso_();
  const resp = UrlFetchApp.fetch(String(source.url), {
    method: 'get',
    headers,
    followRedirects: true,
    muteHttpExceptions: true,
    validateHttpsCertificates: true,
  });

  const status = resp.getResponseCode();
  const responseHeaders = resp.getAllHeaders();
  const etag = headerValue_(responseHeaders, 'ETag');
  const lastModified = headerValue_(responseHeaders, 'Last-Modified');

  if (status === 304) {
    updateObjectById_(APP.SOURCE_SHEET, 'id', source.id, {
      last_checked: now,
      last_success: now,
      http_status: status,
      etag: etag || source.etag,
      last_modified: lastModified || source.last_modified,
      consecutive_errors: 0,
      last_error: '',
    });
    maybeNotifyRecovery_(source);
    return;
  }

  if (status < 200 || status >= 300) {
    throw new Error(`HTTP ${status}`);
  }

  const html = resp.getContentText('UTF-8');
  const blocks = extractBlocks_(html);
  if (blocks.length < 3) {
    throw new Error(`本文抽出結果が少なすぎます (${blocks.length} blocks)`);
  }

  const normalized = blocks.join('\n');
  const newHash = sha256_(normalized);
  const links = extractLinks_(html, String(source.url)).slice(0, APP.MAX_LINKS);

  const patchBase = {
    last_checked: now,
    last_success: now,
    http_status: status,
    etag: etag || '',
    last_modified: lastModified || '',
    consecutive_errors: 0,
    last_error: '',
  };

  // 初回はbaselineだけ保存。
  if (!source.content_hash || !source.snapshot_file_id) {
    const fileId = saveSnapshot_(source, normalized);
    updateObjectById_(APP.SOURCE_SHEET, 'id', source.id, {
      ...patchBase,
      content_hash: newHash,
      snapshot_file_id: fileId,
      links_json: safeJson_(links),
    });
    log_('INFO', 'baseline', source.id, '', `${source.name} baseline保存`);
    maybeNotifyRecovery_(source);
    return;
  }

  if (String(source.content_hash) === newHash) {
    updateObjectById_(APP.SOURCE_SHEET, 'id', source.id, patchBase);
    maybeNotifyRecovery_(source);
    return;
  }

  const oldText = readSnapshot_(source.snapshot_file_id);
  const oldLinks = parseJson_(source.links_json, []);
  const diff = calculateDiff_(oldText, normalized);
  const relevant = assessRelevance_(diff);

  // snapshotは候補対象外でも更新する。そうしないと同じ差分を何度も拾うため。
  const fileId = saveSnapshot_(source, normalized, source.snapshot_file_id);
  updateObjectById_(APP.SOURCE_SHEET, 'id', source.id, {
    ...patchBase,
    content_hash: newHash,
    snapshot_file_id: fileId,
    links_json: safeJson_(links),
  });
  maybeNotifyRecovery_(source);

  if (!relevant.isRelevant) {
    log_('INFO', 'ignored_change', source.id, '', `対象外変更 score=${relevant.score}`);
    return;
  }

  const bestLink = findBestRelevantNewLink_(oldLinks, links, diff, source.url);
  let candidatePage = {
    url: String(source.url),
    html,
    blocks,
    title: extractTitle_(html) || String(source.name),
  };

  if (bestLink && bestLink.href !== source.url) {
    try {
      const detailResp = UrlFetchApp.fetch(bestLink.href, {
        method: 'get',
        headers: {
          'User-Agent': 'Mozilla/5.0 (compatible; ShelterSignagePoC/1.0; GoogleAppsScript)',
          'Accept-Language': 'ja,en;q=0.7',
        },
        followRedirects: true,
        muteHttpExceptions: true,
      });
      if (detailResp.getResponseCode() >= 200 && detailResp.getResponseCode() < 300) {
        const detailHtml = detailResp.getContentText('UTF-8');
        const detailBlocks = extractBlocks_(detailHtml);
        if (detailBlocks.length >= 3 && scoreText_(detailBlocks.join(' ')) > 0) {
          candidatePage = {
            url: bestLink.href,
            html: detailHtml,
            blocks: detailBlocks,
            title: extractTitle_(detailHtml) || bestLink.text || String(source.name),
          };
        }
      }
    } catch (e) {
      log_('WARN', 'detail_fetch_failed', source.id, '', `${bestLink.href}: ${e.message}`);
    }
  }

  const draft = buildDraft_(candidatePage, source, diff);
  const candidateId = Utilities.getUuid();
  const detectedAt = nowIso_();

  appendObject_(APP.CANDIDATE_SHEET, CANDIDATE_HEADERS, {
    id: candidateId,
    source_id: source.id,
    source_name: source.name,
    status: 'REVIEW_REQUIRED',
    detected_at: detectedAt,
    notified_at: '',
    first_viewed_at: '',
    reviewed_at: '',
    reviewer: '',
    title: draft.title,
    source_url: draft.sourceUrl,
    diff_summary: relevant.summary,
    old_excerpt: diff.removed.slice(0, 12).join('\n'),
    new_excerpt: diff.added.slice(0, 12).join('\n'),
    draft_json: safeJson_(draft),
    pdf_file_id: '',
    pdf_url: '',
    review_note: '',
    old_hash: String(source.content_hash),
    new_hash: newHash,
    updated_at: detectedAt,
  });

  try {
    const pdf = generateCandidatePdf_(candidateId, draft);
    updateObjectById_(APP.CANDIDATE_SHEET, 'id', candidateId, {
      pdf_file_id: pdf.fileId,
      pdf_url: pdf.url,
      updated_at: nowIso_(),
    });
  } catch (e) {
    log_('ERROR', 'pdf_failed', source.id, candidateId, e.message);
  }

  const notified = notifyCandidate_(candidateId);
  if (notified) {
    updateObjectById_(APP.CANDIDATE_SHEET, 'id', candidateId, {
      notified_at: nowIso_(),
      updated_at: nowIso_(),
    });
  }
  log_('INFO', 'candidate_created', source.id, candidateId, draft.title);
}

function handleSourceError_(source, err) {
  const currentErrors = Number(source.consecutive_errors || 0) + 1;
  updateObjectById_(APP.SOURCE_SHEET, 'id', source.id, {
    last_checked: nowIso_(),
    http_status: 'ERROR',
    consecutive_errors: currentErrors,
    last_error: String(err && err.message ? err.message : err),
  });
  log_('ERROR', 'monitor_failed', source.id, '', String(err && err.stack ? err.stack : err));

  if (currentErrors >= APP.ERROR_NOTIFY_AFTER && !source.error_notified_at) {
    const sent = notifySystem_(
      `⚠️ 監視エラー: ${source.name}`,
      `${source.name}\n${source.url}\n${currentErrors}回連続で取得に失敗しました。\n${String(err.message || err)}`
    );
    if (sent) {
      updateObjectById_(APP.SOURCE_SHEET, 'id', source.id, {
        error_notified_at: nowIso_(),
      });
    }
  }
}

function maybeNotifyRecovery_(source) {
  if (Number(source.consecutive_errors || 0) > 0 || source.error_notified_at) {
    notifySystem_(
      `✅ 監視復旧: ${source.name}`,
      `${source.name}\n${source.url}\n取得が再び成功しました。`
    );
    updateObjectById_(APP.SOURCE_SHEET, 'id', source.id, {
      error_notified_at: '',
      recovered_at: nowIso_(),
      consecutive_errors: 0,
      last_error: '',
    });
  }
}

/** -------------------- 差分・抽出 -------------------- */

function extractBlocks_(html) {
  let s = String(html || '');
  s = s.replace(/<!--([\s\S]*?)-->/g, ' ');
  s = s.replace(/<(script|style|svg|noscript|template|iframe)[^>]*>[\s\S]*?<\/\1>/gi, ' ');
  s = s.replace(/<br\s*\/?\s*>/gi, '\n');
  s = s.replace(/<\/(h[1-6]|p|li|div|section|article|tr|td|th|dt|dd)>/gi, '\n');
  s = s.replace(/<[^>]+>/g, ' ');
  s = decodeEntities_(s);
  s = s.replace(/\r/g, '\n');
  s = s.replace(/[\t\f\v]+/g, ' ');
  s = s.replace(/ +/g, ' ');

  const raw = s.split(/\n+/).map(x => normalizeWhitespace_(x)).filter(Boolean);
  const out = [];
  const seen = new Set();
  raw.forEach(line => {
    if (line.length < 2) return;
    if (BOILERPLATE_PATTERNS.some(re => re.test(line))) return;
    if (/^(ホーム|Home|English|やさしい日本語)$/.test(line)) return;
    if (line.length > 1200) {
      splitSentences_(line).forEach(x => pushUniqueBlock_(out, seen, x));
    } else {
      pushUniqueBlock_(out, seen, line);
    }
  });
  return out;
}

function pushUniqueBlock_(out, seen, line) {
  const v = normalizeWhitespace_(line);
  if (!v || v.length < 2) return;
  const key = v.toLowerCase();
  if (seen.has(key)) return;
  seen.add(key);
  out.push(v);
}

function extractTitle_(html) {
  const s = String(html || '');
  const h1 = s.match(/<h1[^>]*>([\s\S]*?)<\/h1>/i);
  if (h1) return cleanInlineHtml_(h1[1]);
  const og = s.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)["']/i) ||
             s.match(/<meta[^>]+content=["']([^"']+)["'][^>]+property=["']og:title["']/i);
  if (og) return decodeEntities_(og[1]).trim();
  const title = s.match(/<title[^>]*>([\s\S]*?)<\/title>/i);
  return title ? cleanInlineHtml_(title[1]) : '';
}

function extractLinks_(html, baseUrl) {
  const out = [];
  const seen = new Set();
  const re = /<a\b[^>]*href\s*=\s*["']([^"'#]+)["'][^>]*>([\s\S]*?)<\/a>/gi;
  let m;
  while ((m = re.exec(String(html || ''))) !== null) {
    const href = resolveUrl_(baseUrl, decodeEntities_(m[1]).trim());
    const text = cleanInlineHtml_(m[2]);
    if (!href || !/^https?:\/\//i.test(href)) continue;
    if (!text || text.length > 240) continue;
    const key = href + '|' + text;
    if (seen.has(key)) continue;
    seen.add(key);
    out.push({ href, text });
    if (out.length >= APP.MAX_LINKS) break;
  }
  return out;
}

function calculateDiff_(oldText, newText) {
  const oldLines = String(oldText || '').split('\n').map(normalizeWhitespace_).filter(Boolean);
  const newLines = String(newText || '').split('\n').map(normalizeWhitespace_).filter(Boolean);
  const oldCount = countMap_(oldLines);
  const newCount = countMap_(newLines);
  const added = [];
  const removed = [];

  newLines.forEach(line => {
    const n = newCount[line] || 0;
    const o = oldCount[line] || 0;
    if (n > o && added.filter(x => x === line).length < (n - o)) added.push(line);
  });
  oldLines.forEach(line => {
    const o = oldCount[line] || 0;
    const n = newCount[line] || 0;
    if (o > n && removed.filter(x => x === line).length < (o - n)) removed.push(line);
  });

  return {
    added: added.slice(0, APP.MAX_DIFF_LINES),
    removed: removed.slice(0, APP.MAX_DIFF_LINES),
  };
}

function assessRelevance_(diff) {
  const addedText = diff.added.join(' ');
  const removedText = diff.removed.join(' ');
  const combined = `${addedText} ${removedText}`;
  let score = scoreText_(combined);

  const negativeHits = NEGATIVE_KEYWORDS.filter(k => combined.includes(k)).length;
  const positiveHits = POSITIVE_KEYWORDS.filter(k => combined.includes(k)).length;
  if (negativeHits > 0 && positiveHits === 0) score -= negativeHits * 3;

  const relevantAdded = diff.added.filter(x => scoreText_(x) > 0).slice(0, 6);
  const relevantRemoved = diff.removed.filter(x => scoreText_(x) > 0).slice(0, 4);
  const summaryParts = [];
  if (relevantAdded.length) summaryParts.push('追加: ' + relevantAdded.join(' / '));
  if (relevantRemoved.length) summaryParts.push('削除・変更前: ' + relevantRemoved.join(' / '));
  const summary = summaryParts.join('\n').slice(0, 3000);

  return {
    isRelevant: score >= 2 && (relevantAdded.length > 0 || relevantRemoved.length > 0),
    score,
    summary: summary || `関連語スコア ${score}`,
  };
}

function scoreText_(text) {
  const s = String(text || '');
  let score = 0;
  POSITIVE_KEYWORDS.forEach(k => { if (s.includes(k)) score += 1; });
  HIGH_PRIORITY_KEYWORDS.forEach(k => { if (s.includes(k)) score += 1; });
  if (/\d{1,2}[:：]\d{2}/.test(s)) score += 1;
  if (/\d{1,2}月\d{1,2}日/.test(s)) score += 1;
  if (/0\d{1,4}-\d{1,4}-\d{3,4}/.test(s)) score += 1;
  return score;
}

function findBestRelevantNewLink_(oldLinks, newLinks, diff, sourceUrl) {
  const oldHref = new Set((oldLinks || []).map(x => String(x.href || '')));
  const addedText = diff.added.join(' ');
  const candidates = (newLinks || [])
    .filter(x => !oldHref.has(String(x.href || '')) && sameHost_(x.href, sourceUrl))
    .map(x => {
      let score = scoreText_(x.text || '');
      if (x.text && addedText.includes(x.text)) score += 3;
      return { ...x, score };
    })
    .filter(x => x.score > 0)
    .sort((a, b) => b.score - a.score);
  return candidates[0] || null;
}

function buildDraft_(page, source, diff) {
  const blocks = page.blocks || [];
  const title = normalizeWhitespace_(page.title || source.name || 'お知らせ').slice(0, 100);

  const candidates = [];
  blocks.forEach((block, idx) => {
    splitSentences_(block).forEach(sentence => {
      const text = normalizeWhitespace_(sentence);
      if (!text || text === title || text.length < 4 || text.length > APP.MAX_BLOCK_CHARS) return;
      const score = scoreText_(text) + factScore_(text) + (idx < 12 ? 1 : 0);
      if (score > 0) candidates.push({ text, score, idx });
    });
  });

  candidates.sort((a, b) => b.score - a.score || a.idx - b.idx);
  const bullets = [];
  const seen = new Set();
  candidates.forEach(c => {
    const key = c.text.replace(/\s/g, '');
    if (seen.has(key)) return;
    seen.add(key);
    bullets.push(c.text);
  });

  // 内容が少ない場合は原文先頭から補完。ただし原文そのものを使用する。
  if (bullets.length < 3) {
    blocks.slice(0, 30).forEach(block => {
      splitSentences_(block).forEach(sentence => {
        const text = normalizeWhitespace_(sentence);
        if (!text || text === title || text.length < 5 || text.length > APP.MAX_BLOCK_CHARS) return;
        const key = text.replace(/\s/g, '');
        if (seen.has(key)) return;
        seen.add(key);
        bullets.push(text);
      });
    });
  }

  const finalBullets = bullets.slice(0, APP.MAX_BULLETS);
  const lead = finalBullets.shift() || '更新内容を原文で確認してください。';

  const contact = findContact_(blocks);
  return {
    title,
    lead,
    bullets: finalBullets,
    contact,
    sourceName: String(source.name || ''),
    sourceUrl: String(page.url || source.url || ''),
    checkedAt: formatJst_(new Date()),
    note: '原文から抽出した自動生成ドラフトです。配信前に必ず確認してください。',
    diffAdded: diff.added.slice(0, 8),
    diffRemoved: diff.removed.slice(0, 6),
  };
}

function factScore_(text) {
  let score = 0;
  const s = String(text || '');
  if (/\d{1,2}[:：]\d{2}/.test(s)) score += 3;
  if (/\d{1,2}月\d{1,2}日|令和\d+年/.test(s)) score += 2;
  if (/場所|会場|住所|所在地|丁目|番地|市役所|避難所|公民館|学校|港/.test(s)) score += 2;
  if (/持参|必要|対象|期限|終了|中止|休止|再開/.test(s)) score += 2;
  if (/問い合わせ|問合せ|電話|TEL|☎|0\d{1,4}-\d{1,4}-\d{3,4}/i.test(s)) score += 2;
  return score;
}

function findContact_(blocks) {
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i];
    if (/問い合わせ|問合せ|お問い合わせ|電話|TEL|☎|0\d{1,4}-\d{1,4}-\d{3,4}/i.test(b)) {
      const parts = [b];
      if (i + 1 < blocks.length && blocks[i + 1].length < 160) parts.push(blocks[i + 1]);
      return parts.join(' / ').slice(0, 260);
    }
  }
  return '';
}

/** -------------------- PDF -------------------- */

function generateCandidatePdf_(candidateId, draft) {
  const temp = SpreadsheetApp.create(`tmp_signage_${candidateId.slice(0, 8)}`);
  try {
    const sh = temp.getSheets()[0];
    sh.setName('Signage');
    sh.setHiddenGridlines(true);

    // A4横1ページを意識した固定レイアウト。export時はscale=4 (fit to page)。
    sh.setColumnWidths(1, 12, 72);
    for (let r = 1; r <= 16; r++) sh.setRowHeight(r, 34);
    sh.setRowHeights(1, 2, 42);
    sh.setRowHeights(3, 2, 36);
    sh.setRowHeights(14, 3, 28);

    sh.getRange('A1:L2').merge()
      .setValue(draft.title || 'お知らせ')
      .setFontSize(26).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle')
      .setWrap(true);

    sh.getRange('A3:L4').merge()
      .setValue(draft.lead || '')
      .setFontSize(19).setFontWeight('bold')
      .setVerticalAlignment('middle').setWrap(true);

    const bullets = (draft.bullets || []).slice(0, APP.MAX_BULLETS);
    let row = 5;
    bullets.forEach(b => {
      sh.getRange(row, 1, 1, 12).merge()
        .setValue('● ' + b)
        .setFontSize(16)
        .setVerticalAlignment('middle')
        .setWrap(true);
      row += 1;
    });

    while (row <= 11) {
      sh.getRange(row, 1, 1, 12).merge().setValue('');
      row++;
    }

    if (draft.contact) {
      sh.getRange('A12:L12').merge()
        .setValue('問い合わせ: ' + draft.contact)
        .setFontSize(11).setWrap(true);
    } else {
      sh.getRange('A12:L12').merge().setValue('');
    }

    sh.getRange('A13:L13').merge()
      .setValue(draft.note || '')
      .setFontSize(10).setFontWeight('bold').setWrap(true);

    sh.getRange('A14:L14').merge()
      .setValue(`情報提供元: ${draft.sourceName || ''}`)
      .setFontSize(9).setWrap(true);
    sh.getRange('A15:L15').merge()
      .setValue(`確認日時: ${draft.checkedAt || formatJst_(new Date())}`)
      .setFontSize(9).setWrap(true);
    sh.getRange('A16:L16').merge()
      .setValue(`原文: ${draft.sourceUrl || ''}`)
      .setFontSize(8).setWrap(true);

    sh.getRange('A1:L16').setFontFamily('Arial').setVerticalAlignment('middle');
    sh.getRange('A1:L16').setBorder(false, false, false, false, false, false);
    SpreadsheetApp.flush();
    Utilities.sleep(800);

    const ssId = temp.getId();
    const gid = sh.getSheetId();
    const url = `https://docs.google.com/spreadsheets/d/${ssId}/export` +
      `?format=pdf&size=7&portrait=false&scale=4&fitw=true` +
      `&sheetnames=false&printtitle=false&pagenum=UNDEFINED&gridlines=false&fzr=false` +
      `&top_margin=0.25&bottom_margin=0.25&left_margin=0.25&right_margin=0.25` +
      `&gid=${gid}&r1=0&c1=0&r2=16&c2=12`;

    const response = UrlFetchApp.fetch(url, {
      headers: { Authorization: `Bearer ${ScriptApp.getOAuthToken()}` },
      muteHttpExceptions: false,
    });
    const blob = response.getBlob().setName(`${dateStamp_()}_${safeFileName_(draft.title || 'signage')}.pdf`);
    const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('PDF_FOLDER_ID'));
    const file = folder.createFile(blob);

    if (APP.PDF_PUBLIC_LINK) {
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {
        log_('WARN', 'pdf_share_failed', '', candidateId, e.message);
      }
    }

    return { fileId: file.getId(), url: file.getUrl(), size: blob.getBytes().length };
  } finally {
    try { DriveApp.getFileById(temp.getId()).setTrashed(true); } catch (e) {}
  }
}

function regenerateCandidatePdf(token, candidateId, edited, reviewer) {
  assertToken_(token);
  const candidate = getObjectById_(APP.CANDIDATE_SHEET, 'id', candidateId);
  if (!candidate) throw new Error('候補が見つかりません。');
  const draft = parseJson_(candidate.draft_json, {});

  if (edited && typeof edited === 'object') {
    if (edited.title) draft.title = normalizeWhitespace_(edited.title).slice(0, 100);
    if (edited.lead) draft.lead = normalizeWhitespace_(edited.lead).slice(0, 220);
    if (Array.isArray(edited.bullets)) {
      draft.bullets = edited.bullets.map(normalizeWhitespace_).filter(Boolean).slice(0, APP.MAX_BULLETS);
    }
    draft.checkedAt = formatJst_(new Date());
  }

  const pdf = generateCandidatePdf_(candidateId, draft);
  updateObjectById_(APP.CANDIDATE_SHEET, 'id', candidateId, {
    title: draft.title,
    draft_json: safeJson_(draft),
    pdf_file_id: pdf.fileId,
    pdf_url: pdf.url,
    reviewer: String(reviewer || candidate.reviewer || ''),
    updated_at: nowIso_(),
  });
  log_('INFO', 'pdf_regenerated', candidate.source_id, candidateId, `by ${reviewer || ''}`);
  return getCandidateDetails(token, candidateId);
}

/** -------------------- 通知 -------------------- */

function notifyCandidate_(candidateId) {
  const candidate = getObjectById_(APP.CANDIDATE_SHEET, 'id', candidateId);
  if (!candidate) return false;
  const reviewUrl = reviewUrl_(candidateId);
  const subject = `【要確認】サイネージ更新候補: ${candidate.title}`;
  const text = [
    '🔴 サイネージ更新候補',
    candidate.title,
    `情報元: ${candidate.source_name}`,
    `検知: ${candidate.detected_at}`,
    '',
    candidate.diff_summary || '',
    '',
    `確認: ${reviewUrl || '(Webアプリ未デプロイ)'}`,
    `原文: ${candidate.source_url}`,
  ].join('\n');
  return notifySystem_(subject, text);
}

function notifySystem_(subject, text) {
  const props = PropertiesService.getScriptProperties();
  let sent = false;

  const email = props.getProperty('NOTIFY_EMAIL');
  if (email) {
    try {
      MailApp.sendEmail({
        to: email,
        subject,
        body: text,
        htmlBody: escapeHtml_(text).replace(/\n/g, '<br>'),
      });
      sent = true;
    } catch (e) {
      log_('ERROR', 'email_failed', '', '', e.message);
    }
  }

  const webhook = props.getProperty('SLACK_WEBHOOK_URL');
  if (webhook) {
    try {
      const resp = UrlFetchApp.fetch(webhook, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ text: `${subject}\n${text}` }),
        muteHttpExceptions: true,
      });
      if (resp.getResponseCode() >= 200 && resp.getResponseCode() < 300) sent = true;
      else log_('WARN', 'slack_failed', '', '', `HTTP ${resp.getResponseCode()}`);
    } catch (e) {
      log_('ERROR', 'slack_failed', '', '', e.message);
    }
  }
  return sent;
}

/** -------------------- Web App -------------------- */

function doGet(e) {
  const token = e && e.parameter ? String(e.parameter.token || '') : '';
  if (!validToken_(token)) {
    return HtmlService.createHtmlOutput(`
      <html><head><meta name="viewport" content="width=device-width,initial-scale=1"></head>
      <body style="font-family:Arial,sans-serif;padding:32px;max-width:720px;margin:auto">
        <h2>アクセスできません</h2>
        <p>チーム共有URLを確認してください。</p>
      </body></html>`).setTitle(APP.NAME);
  }

  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.token = token;
  return tpl.evaluate()
    .setTitle(APP.NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getDashboard(token) {
  assertToken_(token);
  const sources = getObjects_(APP.SOURCE_SHEET).map(r => ({
    id: r.id,
    name: r.name,
    url: r.url,
    category: r.category,
    active: truthy_(r.active),
    last_checked: r.last_checked,
    last_success: r.last_success,
    http_status: r.http_status,
    consecutive_errors: Number(r.consecutive_errors || 0),
    last_error: r.last_error,
  }));

  const candidates = getObjects_(APP.CANDIDATE_SHEET)
    .sort((a, b) => String(b.detected_at || '').localeCompare(String(a.detected_at || '')))
    .slice(0, 60)
    .map(r => ({
      id: r.id,
      status: r.status,
      detected_at: r.detected_at,
      source_name: r.source_name,
      title: r.title,
      source_url: r.source_url,
      pdf_url: r.pdf_url,
      reviewer: r.reviewer,
      reviewed_at: r.reviewed_at,
    }));

  return {
    appName: APP.NAME,
    now: formatJst_(new Date()),
    sources,
    candidates,
    reviewRequired: candidates.filter(x => x.status === 'REVIEW_REQUIRED').length,
  };
}

function getCandidateDetails(token, candidateId) {
  assertToken_(token);
  const candidate = getObjectById_(APP.CANDIDATE_SHEET, 'id', candidateId);
  if (!candidate) throw new Error('候補が見つかりません。');

  if (!candidate.first_viewed_at) {
    updateObjectById_(APP.CANDIDATE_SHEET, 'id', candidateId, {
      first_viewed_at: nowIso_(),
      updated_at: nowIso_(),
    });
    candidate.first_viewed_at = nowIso_();
  }

  return {
    ...candidate,
    draft: parseJson_(candidate.draft_json, {}),
  };
}

function reviewCandidate(token, candidateId, action, reviewer, note) {
  assertToken_(token);
  const allowed = ['APPROVED', 'REJECTED'];
  if (!allowed.includes(action)) throw new Error('不正な操作です。');
  const candidate = getObjectById_(APP.CANDIDATE_SHEET, 'id', candidateId);
  if (!candidate) throw new Error('候補が見つかりません。');

  updateObjectById_(APP.CANDIDATE_SHEET, 'id', candidateId, {
    status: action,
    reviewer: String(reviewer || '').trim(),
    review_note: String(note || '').trim(),
    reviewed_at: nowIso_(),
    updated_at: nowIso_(),
  });
  log_('INFO', action === 'APPROVED' ? 'approved' : 'rejected', candidate.source_id, candidateId, `by ${reviewer || ''}`);
  return getCandidateDetails(token, candidateId);
}

function runMonitorNow(token) {
  assertToken_(token);
  monitorAllSources();
  return getDashboard(token);
}

function createTestCandidateFromWeb(token) {
  assertToken_(token);
  return createTestCandidate_();
}

function createTestCandidate() {
  return createTestCandidate_();
}

function createTestCandidate_() {
  const id = Utilities.getUuid();
  const draft = {
    title: '【テスト】応急給水のお知らせ',
    lead: '8月23日 11:00から応急給水を実施します。',
    bullets: [
      '場所：八代港内（テスト表示）',
      '終了時刻：17:00予定',
      '給水容器をご持参ください。',
      'この候補はPoCの通知・PDF・承認動作確認用です。',
    ],
    contact: 'テスト問い合わせ先',
    sourceName: 'PoCテストデータ',
    sourceUrl: 'https://example.com/',
    checkedAt: formatJst_(new Date()),
    note: 'テストデータです。実際の災害情報ではありません。',
    diffAdded: ['テスト用更新'],
    diffRemoved: [],
  };
  appendObject_(APP.CANDIDATE_SHEET, CANDIDATE_HEADERS, {
    id,
    source_id: 'TEST',
    source_name: 'PoCテスト',
    status: 'REVIEW_REQUIRED',
    detected_at: nowIso_(),
    title: draft.title,
    source_url: draft.sourceUrl,
    diff_summary: 'テスト候補を手動作成しました。',
    draft_json: safeJson_(draft),
    updated_at: nowIso_(),
  });

  const pdf = generateCandidatePdf_(id, draft);
  updateObjectById_(APP.CANDIDATE_SHEET, 'id', id, {
    pdf_file_id: pdf.fileId,
    pdf_url: pdf.url,
    updated_at: nowIso_(),
  });
  const notified = notifyCandidate_(id);
  if (notified) updateObjectById_(APP.CANDIDATE_SHEET, 'id', id, { notified_at: nowIso_() });
  log_('INFO', 'test_candidate', 'TEST', id, 'テスト候補作成');
  return { id, reviewUrl: reviewUrl_(id), pdfUrl: pdf.url };
}

/** -------------------- Snapshot -------------------- */

function saveSnapshot_(source, text, existingFileId) {
  if (existingFileId) {
    try {
      const f = DriveApp.getFileById(String(existingFileId));
      f.setContent(String(text));
      return f.getId();
    } catch (e) {}
  }
  const folder = DriveApp.getFolderById(PropertiesService.getScriptProperties().getProperty('SNAPSHOT_FOLDER_ID'));
  const file = folder.createFile(`${source.id}.txt`, String(text), MimeType.PLAIN_TEXT);
  return file.getId();
}

function readSnapshot_(fileId) {
  try {
    return DriveApp.getFileById(String(fileId)).getBlob().getDataAsString('UTF-8');
  } catch (e) {
    return '';
  }
}

/** -------------------- Spreadsheet helpers -------------------- */

function getDb_() {
  const id = PropertiesService.getScriptProperties().getProperty('DB_ID');
  if (!id) throw new Error('DB_IDがありません。setupPoc()を実行してください。');
  return SpreadsheetApp.openById(id);
}

function getObjects_(sheetName) {
  const sh = getDb_().getSheetByName(sheetName);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const headers = values.shift().map(String);
  return values.map((row, idx) => {
    const obj = { __row: idx + 2 };
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function getObjectById_(sheetName, key, value) {
  const rows = getObjects_(sheetName);
  return rows.find(r => String(r[key]) === String(value)) || null;
}

function appendObject_(sheetName, headers, obj) {
  const sh = getDb_().getSheetByName(sheetName);
  sh.appendRow(headers.map(h => obj[h] !== undefined ? obj[h] : ''));
}

function updateObjectById_(sheetName, key, value, patch) {
  const sh = getDb_().getSheetByName(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return false;
  const headers = data[0].map(String);
  const keyCol = headers.indexOf(key);
  if (keyCol < 0) throw new Error(`列がありません: ${key}`);

  for (let r = 1; r < data.length; r++) {
    if (String(data[r][keyCol]) === String(value)) {
      Object.keys(patch).forEach(k => {
        const c = headers.indexOf(k);
        if (c >= 0) sh.getRange(r + 1, c + 1).setValue(patch[k]);
      });
      return true;
    }
  }
  return false;
}

function log_(level, event, sourceId, candidateId, message) {
  try {
    const sh = getDb_().getSheetByName(APP.LOG_SHEET);
    sh.appendRow([nowIso_(), level, event, sourceId || '', candidateId || '', String(message || '').slice(0, 5000)]);
  } catch (e) {
    console.log(`${level} ${event}: ${message}`);
  }
}

/** -------------------- URL / HTML / Utility -------------------- */

function reviewUrl_(candidateId) {
  const base = ScriptApp.getService().getUrl();
  const token = PropertiesService.getScriptProperties().getProperty('REVIEW_TOKEN');
  if (!base || !token) return '';
  return `${base}?token=${encodeURIComponent(token)}#candidate=${encodeURIComponent(candidateId || '')}`;
}

function validToken_(token) {
  const expected = PropertiesService.getScriptProperties().getProperty('REVIEW_TOKEN');
  return !!expected && String(token || '') === expected;
}

function assertToken_(token) {
  if (!validToken_(token)) throw new Error('認証に失敗しました。');
}

function moveFileToFolder_(fileId, folderId) {
  const file = DriveApp.getFileById(fileId);
  const folder = DriveApp.getFolderById(folderId);
  file.moveTo(folder);
}

function resolveUrl_(base, href) {
  const h = String(href || '').trim();
  if (!h || /^javascript:|^mailto:|^tel:/i.test(h)) return '';
  if (/^https?:\/\//i.test(h)) return h;
  const bm = String(base || '').match(/^(https?):\/\/([^\/]+)(\/.*)?$/i);
  if (!bm) return '';
  const origin = `${bm[1]}://${bm[2]}`;
  if (h.startsWith('//')) return `${bm[1]}:${h}`;
  if (h.startsWith('/')) return origin + h;
  const path = (bm[3] || '/').split('?')[0].split('#')[0];
  const dir = path.endsWith('/') ? path : path.substring(0, path.lastIndexOf('/') + 1);
  const parts = (dir + h).split('/');
  const clean = [];
  parts.forEach(p => {
    if (!p || p === '.') return;
    if (p === '..') clean.pop();
    else clean.push(p);
  });
  return origin + '/' + clean.join('/');
}


function sameHost_(urlA, urlB) {
  const a = String(urlA || '').match(/^https?:\/\/([^\/]+)/i);
  const b = String(urlB || '').match(/^https?:\/\/([^\/]+)/i);
  return !!a && !!b && a[1].toLowerCase() === b[1].toLowerCase();
}

function cleanInlineHtml_(s) {
  return normalizeWhitespace_(decodeEntities_(String(s || '').replace(/<[^>]+>/g, ' ')));
}

function decodeEntities_(s) {
  let x = String(s || '');
  const map = {
    '&nbsp;': ' ', '&amp;': '&', '&lt;': '<', '&gt;': '>', '&quot;': '"',
    '&#39;': "'", '&apos;': "'", '&yen;': '¥', '&copy;': '©'
  };
  Object.keys(map).forEach(k => { x = x.split(k).join(map[k]); });
  x = x.replace(/&#(\d+);/g, (_, n) => String.fromCharCode(Number(n)));
  x = x.replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCharCode(parseInt(n, 16)));
  return x;
}

function normalizeWhitespace_(s) {
  return String(s || '').replace(/\u3000/g, ' ').replace(/\s+/g, ' ').trim();
}

function splitSentences_(s) {
  const text = normalizeWhitespace_(s);
  if (!text) return [];
  const parts = text.match(/[^。！？!?]+[。！？!?]?/g) || [text];
  return parts.map(normalizeWhitespace_).filter(Boolean);
}

function countMap_(arr) {
  const m = {};
  arr.forEach(x => m[x] = (m[x] || 0) + 1);
  return m;
}

function sha256_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(text), Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function headerValue_(headers, name) {
  const target = String(name).toLowerCase();
  for (const k in headers) {
    if (String(k).toLowerCase() === target) return String(headers[k]);
  }
  return '';
}

function parseJson_(s, fallback) {
  try { return JSON.parse(String(s || '')); } catch (e) { return fallback; }
}

function safeJson_(obj) {
  let value = obj;
  let s = JSON.stringify(value === undefined ? {} : value);
  if (s.length <= 48000) return s;

  // Sheetsの1セル上限対策。リンク配列などは末尾を落として必ず正しいJSONで保存する。
  if (Array.isArray(value)) {
    let arr = value.slice();
    while (arr.length > 0) {
      arr = arr.slice(0, Math.floor(arr.length * 0.75));
      s = JSON.stringify(arr);
      if (s.length <= 48000) return s;
    }
    return '[]';
  }

  // 通常のdraftはこのサイズに達しない。万一の場合も壊れたJSONを保存しない。
  return JSON.stringify({ truncated: true, message: 'データが長すぎるため省略しました。' });
}

function randomToken_(length) {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '').slice(0, Math.max(0, length - 32));
}

function nowIso_() {
  return Utilities.formatDate(new Date(), APP.TZ, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function formatJst_(d) {
  return Utilities.formatDate(d || new Date(), APP.TZ, 'yyyy/MM/dd HH:mm:ss');
}

function dateStamp_() {
  return Utilities.formatDate(new Date(), APP.TZ, 'yyyyMMdd_HHmm');
}

function safeFileName_(s) {
  return String(s || 'signage').replace(/[\\\/:*?"<>|]/g, '').replace(/\s+/g, '_').slice(0, 70) || 'signage';
}

function truthy_(v) {
  if (v === true || v === 1) return true;
  const s = String(v || '').toLowerCase();
  return ['true', '1', 'yes', 'on', '有効'].includes(s);
}

function escapeHtml_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
