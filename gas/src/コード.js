// ================================================================
// おまんぼさんイラストゲーム バックエンド v1.6
//
// 【セットアップ手順】
// 1. このコードをすべて貼り替えて「デプロイ」→「新しいデプロイ」（毎回新しいデプロイが必要）
//    ※URLが変わる場合はgame.htmlのWEBHOOK_URLも更新
// 2. 日次まとめトリガー: sendDailySummary / 時間主導型 / 日タイマー
// 3. 月次リセットトリガー: archiveMonthlyRanking / 時間主導型 / 月タイマー / 毎月1日
// ================================================================

const CONFIG = {
  sheetId:       '1nkGoMoNKc4opnD-Z_Z41I_4re8Brz7lFyBdNXJfpyIw',
  ownerEmail:    'omanbosan.lv@gmail.com',
  driveFolderId: '1gXwYsxpKBqZ6tMKdNyOPGVvtQTcV1QgN'
};

// ----------------------------------------------------------------
// GET
// ----------------------------------------------------------------
function doGet(e) {
  try {
    const type = e && e.parameter && e.parameter.type;
    if (type === 'ranking')      return buildResponse(getRanking(e.parameter.difficulty));
    if (type === 'archive')      return buildResponse(getLastMonthRanking(e.parameter.difficulty));
    if (type === 'zukan')        return buildResponse(getZukan());
    if (type === 'zukanScoring') return buildResponse(getZukanScoring());
    if (type === 'pending')        return buildResponse(getPending());
    if (type === 'pendingScores')  return buildResponse(getPendingScores());
    if (type === 'voteRanking')    return buildResponse(getVoteRanking());
    if (type === 'config')         return buildResponse(getConfig());
    if (type === 'stats')                return buildResponse(getStats());
    if (type === 'unlinkedScoreImages')  return buildResponse(getUnlinkedScoreImages());
    return buildResponse({ ok: true, message: 'おまんぼさんイラストゲームAPI v1.9' });
  } catch(err) {
    return buildResponse({ error: err.message });
  }
}

// ----------------------------------------------------------------
// POST
// ----------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.type === 'score')      return buildResponse(saveScore(data));
    if (data.type === 'drawing')    return buildResponse(saveDrawing(data));
    if (data.type === 'setApproval') return buildResponse(setApproval(data));
    if (data.type === 'vote')       return buildResponse(castVote(data));
    if (data.type === 'unvote')     return buildResponse(castVote({ ...data, delta: -1 }));
    if (data.type === 'setConfig')  return buildResponse(setConfig(data));
    if (data.type === 'play')             return buildResponse(logPlay(data));
    if (data.type === 'recoverScoreImage') return buildResponse(recoverScoreImage(data));
    return buildResponse({ ok: true });
  } catch(err) {
    return buildResponse({ error: err.message });
  }
}

function buildResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ----------------------------------------------------------------
// 今月のランキング
// ----------------------------------------------------------------
function getRanking(difficulty) {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('scores');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues()
    .map(r => ({ name: r[0], score: Number(r[1]), date: r[2], instagram: r[3] || '', difficulty: r[5] || 'normal', approved: r[6] || '' }))
    .filter(r => r.name && r.score > 0 && r.approved !== '却下')
    .filter(r => !difficulty || r.difficulty === difficulty)
    .sort((a, b) => b.score - a.score)
    .slice(0, 20);
}

// ----------------------------------------------------------------
// 先月のランキング（アーカイブから取得）
// ----------------------------------------------------------------
function getLastMonthRanking(difficulty) {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('scores_archive');
  if (!sheet || sheet.getLastRow() < 2) return [];

  // 最新の月ラベルを取得
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const latestMonth = rows[rows.length - 1][0];

  return rows
    .filter(r => r[0] === latestMonth)
    .filter(r => !difficulty || (r[5] || 'normal') === difficulty)
    .map(r => ({ month: r[0], name: r[2], score: Number(r[3]), instagram: r[4] || '', difficulty: r[5] || 'normal' }))
    .sort((a, b) => b.score - a.score);
}

// ----------------------------------------------------------------
// スコア保存（採点チャレンジ図鑑対応）
// scores列: 名前, スコア, 日付, Instagram, imageUrl, 難易度, 承認
// ----------------------------------------------------------------
function saveScore(data) {
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('scores');
  if (!sheet) {
    sheet = ss.insertSheet('scores');
    sheet.appendRow(['名前', 'スコア', '日付', 'Instagram', '画像URL', '難易度', '承認']);
    sheet.setFrozenRows(1);
  }
  // ヘッダー列拡張（既存シート対応）
  const colCount = sheet.getLastColumn();
  if (colCount < 5) sheet.getRange(1, 5).setValue('画像URL');
  if (colCount < 6) sheet.getRange(1, 6).setValue('難易度');
  if (colCount < 7) sheet.getRange(1, 7).setValue('承認');
  if (colCount < 8) sheet.getRange(1, 8).setValue('照合コード');
  if (colCount < 9) sheet.getRange(1, 9).setValue('採点レベル');

  const name       = (data.name || '名無し').slice(0, 20);
  const score      = Number(data.score) || 0;
  const instagram  = (data.instagram || '').replace('@', '');
  const difficulty = data.difficulty || 'normal';
  const entryId    = String(data.entryId || '');
  const luckLevel  = Number(data.luckLevel) || 2; // 🔥の数(1〜4)
  const rowIndex   = sheet.getLastRow() + 1;

  sheet.appendRow([name, score, new Date().toLocaleDateString('ja-JP'), instagram,
                   'Drive保存中...', difficulty, '審査中', entryId, luckLevel]);

  // Drive に画像を保存
  if (data.image) {
    try {
      const base64 = data.image.replace(/^data:image\/\w+;base64,/, '');
      const folder = DriveApp.getFolderById(CONFIG.driveFolderId);
      const blob   = Utilities.newBlob(
        Utilities.base64Decode(base64), 'image/png',
        `score_${name}_${Date.now()}.png`
      );
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      sheet.getRange(rowIndex, 5).setValue(file.getUrl());
    } catch(err) {
      sheet.getRange(rowIndex, 5).setValue('Drive保存失敗: ' + err.message);
    }
  } else {
    sheet.getRange(rowIndex, 5).setValue('');
  }

  return { ok: true, entryId };
}

// ----------------------------------------------------------------
// 採点チャレンジ図鑑取得（承認済みのみ）
// 月次アーカイブ（scores_zukan_archive・過去分）+ scoresシート（今月分）を
// 古い順に結合して返す。月が変わってscoresがクリアされても画像が消えないようにする
// ----------------------------------------------------------------
function getZukanScoring() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const results = [];

  // 過去分（永久保存・承認済みのみ）
  const archSheet = ss.getSheetByName('scores_zukan_archive');
  if (archSheet && archSheet.getLastRow() >= 2) {
    archSheet.getRange(2, 1, archSheet.getLastRow() - 1, 7).getValues()
      .filter(r => r[0])
      .forEach(r => results.push({
        name:       r[0] || '名無し',
        score:      Number(r[1]) || 0,
        date:       r[2] ? String(r[2]).slice(0, 10) : '',
        imageUrl:   driveUrlToThumb(r[3]),
        difficulty: r[4] || 'normal',
        luckLevel:  Number(r[5]) || 2
      }));
  }

  // 今月分（scoresシート、承認済みのみ）
  const sheet = ss.getSheetByName('scores');
  if (sheet && sheet.getLastRow() >= 2) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues()
      .filter(r => r[0] && (r[6] || '') === '承認済み')
      .forEach(r => results.push({
        name:       r[0] || '名無し',
        score:      Number(r[1]) || 0,
        date:       r[2] ? String(r[2]).slice(0, 10) : '',
        imageUrl:   driveUrlToThumb(r[4]),
        difficulty: r[5] || 'normal',
        luckLevel:  Number(r[8]) || 2
      }));
  }

  return results.map((r, i) => Object.assign({ no: i + 1 }, r)); // 登録順（古い順 = No.1が最初）
}

// ----------------------------------------------------------------
// 審査中一覧取得（管理画面用）
// ----------------------------------------------------------------
function getPending() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues()
    .map((r, i) => ({
      row:         i + 2,
      instagram:   (r[0] || '').replace('@', ''),
      date:        r[1] ? String(r[1]).slice(0, 10) : '',
      imageUrl:    driveUrlToThumb(r[2]),
      villageName: r[4] || '',
      comment:     r[5] || '',
      approved:    r[7] || '審査中',
      penname:     r[8] || '',
      igShow:      r[9] === true || r[9] === 'TRUE',
      entryId:     r[11] || ''
    }))
    .filter(r => r.villageName);
}

// scores シートの審査待ちエントリ（管理画面用）
// scores列: 名前(1), スコア(2), 日付(3), Instagram(4), imageUrl(5), 難易度(6), 承認(7), 参加番号(8)
function getPendingScores() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('scores');
  if (!sheet || sheet.getLastRow() < 2) return [];

  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues()
    .map((r, i) => ({
      row:        i + 2,
      name:       r[0] || '名無し',
      score:      Number(r[1]) || 0,
      date:       r[2] ? String(r[2]).slice(0, 10) : '',
      instagram:  (r[3] || '').replace('@', ''),
      imageUrl:   driveUrlToThumb(r[4]),
      difficulty: r[5] || 'normal',
      approved:   r[6] || '審査中',
      entryId:    r[7] || '',
      luckLevel:  Number(r[8]) || 2
    }))
    .filter(r => r.name);
}

// ----------------------------------------------------------------
// 承認・却下（管理画面用）
// sheetName: 'drawings'（デフォルト）または 'scores'
// ----------------------------------------------------------------
function setApproval(data) {
  const ss        = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheetName = data.sheetName || 'drawings';
  const sheet     = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'no sheet: ' + sheetName };
  const row    = Number(data.row);
  const status = data.status === '承認済み' ? '承認済み' : '却下';
  // drawings: 承認はcol8, scores: 承認はcol7
  const approvalCol = sheetName === 'scores' ? 7 : 8;
  sheet.getRange(row, approvalCol).setValue(status);
  return { ok: true, row, status };
}

// ----------------------------------------------------------------
// DriveのURLをサムネイル表示用URLに変換
// ----------------------------------------------------------------
function driveUrlToThumb(url) {
  if (!url || !url.startsWith('http')) return '';
  const m = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return 'https://drive.google.com/thumbnail?id=' + m[1] + '&sz=w300';
  return '';
}

// ----------------------------------------------------------------
// 図鑑取得（承認済みのみ）
// ----------------------------------------------------------------
function getZukan() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet || sheet.getLastRow() < 2) return [];

  const approved = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues()
    .filter(r => (r[4] || '') && (r[7] || '') === '承認済み');

  // 承認済みのみを連番付与（却下分で番号が飛ばないように）
  return approved.map((r, i) => ({
    no:          i + 1,                          // 連番（承認済みのみカウント）
    instagram:   (r[0] || '').replace('@', ''),
    date:        r[1] ? String(r[1]).slice(0, 10) : '',
    imageUrl:    driveUrlToThumb(r[2]),
    villageName: r[4] || '',
    comment:     r[5] || '',
    penname:     r[8] || '',
    igShow:      r[9] === true || r[9] === 'TRUE'
  })); // 登録順（古い順 = No.1が最初）
}

// ----------------------------------------------------------------
// お絵描き保存（シート記録→Drive保存の順で確実に残す）
// ----------------------------------------------------------------
function saveDrawing(data) {
  if (!data.image) return { error: 'no image data' };
  const ig          = String(data.instagram   || '').replace('@', '');
  const villageName = String(data.villageName || '').slice(0, 10);
  const penname     = String(data.penname     || '').slice(0, 15);
  const comment     = String(data.comment     || '').slice(0, 20);
  const igShow      = data.igShow === true || data.igShow === 'true';

  // 1. シートに仮記録（Drive保存前でも記録が残るように）
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('drawings');
  if (!sheet) {
    sheet = ss.insertSheet('drawings');
    sheet.appendRow(['Instagram', '投稿日時', 'ファイルURL', '送信状態', '村人名', 'ひとこと', '図鑑No', '承認']);
    sheet.setFrozenRows(1);
  }
  // ヘッダー列が足りない場合は拡張
  const colCount = sheet.getLastColumn();
  if (colCount < 5) sheet.getRange(1, 5).setValue('村人名');
  if (colCount < 6) sheet.getRange(1, 6).setValue('ひとこと');
  if (colCount < 7) sheet.getRange(1, 7).setValue('図鑑No');
  if (colCount < 8) sheet.getRange(1, 8).setValue('承認');
  if (colCount < 9)  sheet.getRange(1, 9).setValue('ペンネーム');
  if (colCount < 10) sheet.getRange(1, 10).setValue('IG表示');
  if (colCount < 11) sheet.getRange(1, 11).setValue('著作権同意日時');
  if (colCount < 12) sheet.getRange(1, 12).setValue('照合コード');

  const consentAt  = data.consentAt || '';
  const entryId    = String(data.entryId || '');
  const zukanNo    = Math.max(sheet.getLastRow(), 1);
  const rowIndex   = sheet.getLastRow() + 1;
  sheet.appendRow([
    ig ? '@' + ig : '',
    new Date().toLocaleString('ja-JP'),
    'Drive保存中...',
    '未送信',
    villageName,
    comment,
    zukanNo,
    '審査中',
    penname,
    igShow,
    consentAt,
    entryId
  ]);

  // 2. Drive に保存
  try {
    const base64 = data.image.replace(/^data:image\/\w+;base64,/, '');
    const folder = DriveApp.getFolderById(CONFIG.driveFolderId);
    const blob   = Utilities.newBlob(
      Utilities.base64Decode(base64), 'image/png', `@${ig}_${Date.now()}.png`
    );
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    const fileUrl = file.getUrl();

    // 3. シートのURLを更新
    sheet.getRange(rowIndex, 3).setValue(fileUrl);
    return { ok: true, fileUrl, entryId };
  } catch(err) {
    sheet.getRange(rowIndex, 3).setValue('Drive保存失敗: ' + err.message);
    sheet.getRange(rowIndex, 4).setValue('エラー');
    return { error: err.message, entryId };
  }
}

// ----------------------------------------------------------------
// 投票（図鑑の投票数をvotesシートに記録）
// votes列: no, villageName, count, lastVote
// ----------------------------------------------------------------
function castVote(data) {
  const no          = Number(data.no) || 0;
  const villageName = String(data.villageName || '');
  const delta       = Number(data.delta) || 1; // +1 or -1
  if (!no) return { error: 'no entry no' };

  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('votes');
  if (!sheet) {
    sheet = ss.insertSheet('votes');
    sheet.appendRow(['no', '村人名', '票数', '最終投票日']);
    sheet.setFrozenRows(1);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const rows = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (Number(rows[i][0]) === no) {
        const newCount = Math.max(0, Number(rows[i][2]) + delta);
        sheet.getRange(i + 2, 3).setValue(newCount);
        sheet.getRange(i + 2, 4).setValue(new Date().toLocaleDateString('ja-JP'));
        return { ok: true, no, count: newCount };
      }
    }
  }
  // 新規（取り消しの場合は登録しない）
  if (delta > 0) {
    sheet.appendRow([no, villageName, 1, new Date().toLocaleDateString('ja-JP')]);
    return { ok: true, no, count: 1 };
  }
  return { ok: true, no, count: 0 };
}

// 投票ランキング取得
function getVoteRanking() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('votes');
  if (!sheet || sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues()
    .map(r => ({ no: Number(r[0]), villageName: r[1] || '', count: Number(r[2]) || 0, lastVote: String(r[3] || '') }))
    .filter(r => r.no > 0)
    .sort((a, b) => b.count - a.count);
}

// ----------------------------------------------------------------
// 日次まとめメール（毎日1回トリガーで実行）
// ----------------------------------------------------------------
function sendDailySummary() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet || sheet.getLastRow() < 2) return;

  const rows   = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
  const unsent = rows
    .map((r, i) => ({ ig: r[0], date: r[1], url: r[2], status: r[3], name: r[4], comment: r[5], approved: r[7], row: i + 2 }))
    .filter(r => r.status === '未送信');

  if (unsent.length === 0) return;

  const today = new Date().toLocaleDateString('ja-JP');
  const sheetUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.sheetId}/edit`;
  MailApp.sendEmail({
    to: CONFIG.ownerEmail,
    subject: `【おまんぼさんイラスト】${today} の投稿まとめ（${unsent.length}件）`,
    htmlBody: `
      <h2 style="color:#27ae60;">🎨 お絵描き投稿 日次まとめ【${today}】</h2>
      <p>本日の新着：<strong>${unsent.length}件</strong></p>
      <p style="background:#fff3cd;padding:10px;border-radius:6px;font-size:0.9em;">
        📋 <b>図鑑に載せる場合は「承認」列を <span style="color:green">承認済み</span> に変更してください</b><br>
        問題がある絵は <span style="color:red">却下</span> と入力すると図鑑に表示されません。
      </p>
      <table style="border-collapse:collapse;width:100%;">
        <tr style="background:#27ae60;color:white;">
          <th style="padding:8px;">村人名</th>
          <th style="padding:8px;">ひとこと</th>
          <th style="padding:8px;">Instagram</th>
          <th style="padding:8px;">画像</th>
          <th style="padding:8px;">承認状態</th>
        </tr>
        ${unsent.map((r, i) => `
          <tr style="background:${i%2===0?'#f9f9f9':'white'};">
            <td style="padding:8px;font-weight:bold;">${r.name || '（未入力）'}</td>
            <td style="padding:8px;color:#666;">${r.comment ? '「'+r.comment+'」' : '-'}</td>
            <td style="padding:8px;">${r.ig}</td>
            <td style="padding:8px;"><a href="${r.url}" style="color:#3498db;">画像を見る🔍</a></td>
            <td style="padding:8px;color:#e67e22;font-weight:bold;">${r.approved || '審査中'}</td>
          </tr>
        `).join('')}
      </table>
      <div style="margin-top:16px;display:flex;gap:10px;flex-wrap:wrap;">
        <a href="${sheetUrl}" style="background:#27ae60;color:white;padding:10px 18px;border-radius:6px;text-decoration:none;display:inline-block;">
          📊 スプレッドシートで承認する
        </a>
        <a href="https://drive.google.com/drive/folders/${CONFIG.driveFolderId}" style="background:#3498db;color:white;padding:10px 18px;border-radius:6px;text-decoration:none;display:inline-block;">
          📂 Driveフォルダを確認
        </a>
      </div>
    `
  });

  unsent.forEach(r => sheet.getRange(r.row, 4).setValue('送信済み'));
}

// ----------------------------------------------------------------
// 月次ランキングアーカイブ（毎月1日トリガーで実行）
// ----------------------------------------------------------------
function archiveMonthlyRanking() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('scores');
  if (!sheet || sheet.getLastRow() < 2) return;

  const lastMonth = new Date();
  lastMonth.setMonth(lastMonth.getMonth() - 1);
  const label = `${lastMonth.getFullYear()}年${lastMonth.getMonth() + 1}月`;

  const allRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  // 上位ランキングアーカイブ（メール通知・先月ランキング表示用）
  let archSheet = ss.getSheetByName('scores_archive');
  if (!archSheet) {
    archSheet = ss.insertSheet('scores_archive');
    archSheet.appendRow(['月', '順位', '名前', 'スコア', 'Instagram', '難易度']);
    archSheet.setFrozenRows(1);
  }

  const rows = allRows
    .map(r => ({ name: r[0], score: Number(r[1]), instagram: r[3] || '', difficulty: r[5] || 'normal' }))
    .filter(r => r.name && r.score > 0)
    .sort((a, b) => b.score - a.score)
    .slice(0, 10);

  rows.forEach((r, i) => {
    archSheet.appendRow([label, i + 1, r.name, r.score, r.instagram, r.difficulty]);
  });

  // 採点図鑑用の永久アーカイブ：承認済みの画像・難易度は月をまたいでも消えないよう保存
  let zukanArchSheet = ss.getSheetByName('scores_zukan_archive');
  if (!zukanArchSheet) {
    zukanArchSheet = ss.insertSheet('scores_zukan_archive');
    zukanArchSheet.appendRow(['名前', 'スコア', '日付', '画像URL', '難易度', '採点レベル', '参加番号']);
    zukanArchSheet.setFrozenRows(1);
  }
  allRows
    .filter(r => r[0] && (r[6] || '') === '承認済み')
    .forEach(r => {
      zukanArchSheet.appendRow([r[0], Number(r[1]) || 0, r[2], r[4] || '', r[5] || 'normal', Number(r[8]) || 2, r[7] || '']);
    });

  // 先月結果をメールで送信
  const medals = ['🥇', '🥈', '🥉'];
  MailApp.sendEmail({
    to: CONFIG.ownerEmail,
    subject: `【${label} ランキング確定】おまんぼさんイラストゲーム`,
    htmlBody: `
      <h2>🏆 ${label} ランキング確定結果</h2>
      <table style="border-collapse:collapse;width:100%;">
        <tr style="background:#9b59b6;color:white;">
          <th style="padding:8px;">順位</th><th style="padding:8px;">名前</th>
          <th style="padding:8px;">スコア</th><th style="padding:8px;">Instagram</th>
        </tr>
        ${rows.map((r, i) => `
          <tr style="background:${i%2===0?'#f9f9f9':'white'};">
            <td style="padding:8px;text-align:center;">${medals[i] || i+1}</td>
            <td style="padding:8px;">${r.name}</td>
            <td style="padding:8px;text-align:center;">${r.score}点</td>
            <td style="padding:8px;">${r.instagram ? '@'+r.instagram : '-'}</td>
          </tr>
        `).join('')}
      </table>
      <p style="color:#888;">※ ランキングは本日リセットされました</p>
    `
  });

  // scoresシートをクリア（ヘッダー以外）
  if (sheet.getLastRow() > 1) {
    sheet.deleteRows(2, sheet.getLastRow() - 1);
  }
}

// ----------------------------------------------------------------
// 動作テスト用
// ----------------------------------------------------------------
function testDrive() {
  const folder = DriveApp.getFolderById(CONFIG.driveFolderId);
  Logger.log('フォルダ名: ' + folder.getName());
  Logger.log('✅ Drive接続OK');
}

function testMail() {
  MailApp.sendEmail({
    to: CONFIG.ownerEmail,
    subject: '【テスト】おまんぼさんイラストゲーム GAS接続確認',
    body: 'GASのメール送信テストです。このメールが届いていれば設定完了です！'
  });
  Logger.log('✅ メール送信OK');
}

// ----------------------------------------------------------------
// 【初回のみ実行】承認列をセットアップ
// drawingsシートにH列「承認」を追加し、既存データを全て「承認済み」にする
// ----------------------------------------------------------------
function setupApprovalColumn() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('drawings');
  if (!sheet) { Logger.log('❌ drawingsシートが見つかりません'); return; }

  const lastRow = sheet.getLastRow();

  // H1にヘッダー追加
  sheet.getRange(1, 8).setValue('承認');

  // 既存データ（2行目以降）をすべて「承認済み」に設定
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 8, lastRow - 1, 1);
    const values = Array(lastRow - 1).fill(['承認済み']);
    range.setValues(values);
  }

  Logger.log(`✅ 承認列セットアップ完了！${lastRow - 1}件を「承認済み」に設定しました`);
}

// ----------------------------------------------------------------
// 採点設定（config シート）
// ----------------------------------------------------------------
function getConfig() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('config');
  const autoTheme = getCurrentWeeklyTheme();
  if (!sheet || sheet.getLastRow() < 2) {
    return { hardMult: 2.0, normalMult: 4.5, easyMult: 3.0, hellMult: 2.0, theme: autoTheme, themeOverride: '', autoTheme: autoTheme }; // デフォルト値
  }
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const cfg = {};
  rows.forEach(r => { if (r[0]) cfg[String(r[0])] = r[1]; });
  const themeOverride = cfg['theme'] ? String(cfg['theme']) : '';
  return {
    hardMult:   parseFloat(cfg['hardMult'])   || 2.0,
    normalMult: parseFloat(cfg['normalMult']) || 4.5,
    easyMult:   parseFloat(cfg['easyMult'])   || 3.0,
    hellMult:   parseFloat(cfg['hellMult'])   || 2.0,
    theme:         themeOverride || autoTheme, // 実際に画面へ出す値（手動上書き優先）
    themeOverride: themeOverride,              // 管理画面の入力欄用（生の上書き値）
    autoTheme:     autoTheme                   // 週替わりで自動計算された今週のお題
  };
}

// ----------------------------------------------------------------
// お絵描きモードの週替わりお題（月×週の自動計算・themesシート参照）
// 1〜7日=第1週, 8〜14日=第2週, 15〜21日=第3週, 22日以降=第4週
// ----------------------------------------------------------------
function getCurrentWeeklyTheme() {
  const ss    = SpreadsheetApp.openById(CONFIG.sheetId);
  const sheet = ss.getSheetByName('themes');
  if (!sheet || sheet.getLastRow() < 2) return '';

  const tz    = Session.getScriptTimeZone() || 'Asia/Tokyo';
  const now   = new Date();
  const month = Number(Utilities.formatDate(now, tz, 'M'));
  const day   = Number(Utilities.formatDate(now, tz, 'd'));
  const week  = Math.min(4, Math.ceil(day / 7));

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const match = rows.find(r => Number(r[0]) === month && Number(r[1]) === week);
  return match ? String(match[2]) : '';
}

// ----------------------------------------------------------------
// お絵描きモードのお題データ（キャラクター縛りなし・自由ジャンル）
// setupThemesSheet()（初回構築）とresetThemesToFreeThemes()（既存データの
// 総入れ替え）の両方から参照する単一の定義元。
// ----------------------------------------------------------------
function getFreeThemesData() {
  return {
    1:  ['好きな動物を1匹描いてみよう', '電車や列車を描いてみよう', 'あったかい食べ物を描いてみよう', '雪が積もった景色を描いてみよう'],
    2:  ['魔法使いの杖を持ったキャラクターを描いてみよう', 'お気に入りの部屋の一角を描いてみよう', 'うさぎを描いてみよう', '空を飛ぶ乗り物を描いてみよう'],
    3:  ['甘いお菓子を描いてみよう', 'お花畑を描いてみよう', '小さな妖精を描いてみよう', '宝物にしたいアイテムを描いてみよう'],
    4:  ['鳥を1羽描いてみよう', '船や潜水艦を描いてみよう', 'お弁当を描いてみよう', '虹を描いてみよう'],
    5:  ['ドラゴンを描いてみよう', '新生活で使いたい文房具を描いてみよう', 'ねこを描いてみよう', '自転車を描いてみよう'],
    6:  ['フルーツを描いてみよう', '雨上がりの景色を描いてみよう', '魔法の本を描いてみよう', '読んでみたい本の表紙を想像して描いてみよう'],
    7:  ['かえるを描いてみよう', '宇宙船を描いてみよう', 'かき氷を描いてみよう', '夜空の花火を描いてみよう'],
    8:  ['小さなモンスターを描いてみよう', '星空を見上げているシーンを描いてみよう', '水の中の生き物を描いてみよう', '気球を描いてみよう'],
    9:  ['アイスクリームを描いてみよう', '満月を眺めているシーンを描いてみよう', 'かぼちゃのモンスターを描いてみよう', '自由研究にしたいテーマを描いてみよう'],
    10: ['のんびり歩く動物を描いてみよう', '未来の乗り物を想像して描いてみよう', '秋の味覚を描いてみよう', '紅葉した木を描いてみよう'],
    11: ['空想上の生き物を描いてみよう', '読書の秋、好きな物語のワンシーンを描いてみよう', '大きな動物を描いてみよう', 'お気に入りの乗り物に乗っているシーンを描いてみよう'],
    12: ['世界に一つだけのオリジナルスイーツを描いてみよう', '行ってみたい景色を描いてみよう', '魔法の世界の1コマを描いてみよう', 'あったかい部屋でくつろぐシーンを描いてみよう']
  };
}

function writeThemesData(sheet, themesObj) {
  Object.keys(themesObj).forEach(month => {
    themesObj[month].forEach((theme, i) => {
      sheet.appendRow([Number(month), i + 1, theme]);
    });
  });
}

// ----------------------------------------------------------------
// 【初回のみ実行】themesシートを作成し、12ヶ月×4週分のお題を投入する
// ----------------------------------------------------------------
function setupThemesSheet() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('themes');
  if (sheet) { Logger.log('⚠️ themesシートは既に存在します。処理をスキップしました。'); return; }

  sheet = ss.insertSheet('themes');
  sheet.appendRow(['月', '週', 'お題']);
  sheet.setFrozenRows(1);

  writeThemesData(sheet, getFreeThemesData());

  Logger.log('✅ themesシートに48件のお題を投入しました！');
}

// ----------------------------------------------------------------
// 【総入れ替え用・GASエディタから手動で1回だけ実行すること】
// 既存の48件（キャラクター固定お題）をクリアしてから自由お題に差し替える。
// doGet/doPostには公開しない（HTTP経由で誤って実行されるリスクを避けるため）。
// 実行前に themes シートをGoogle Sheets上で手動複製してバックアップしておくこと。
// ----------------------------------------------------------------
function resetThemesToFreeThemes() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('themes');
  if (!sheet) {
    sheet = ss.insertSheet('themes');
    sheet.appendRow(['月', '週', 'お題']);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).clearContent();
  }
  writeThemesData(sheet, getFreeThemesData());
  Logger.log('✅ themesシートを自由お題48件に総入れ替えしました！');
}

function setConfig(data) {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('config');
  if (!sheet) {
    sheet = ss.insertSheet('config');
    sheet.appendRow(['key', 'value']);
  }
  const lastRow = sheet.getLastRow();
  const rows = lastRow >= 2
    ? sheet.getRange(2, 1, lastRow - 1, 2).getValues()
    : [];

  const setValue = (key, val) => {
    const idx = rows.findIndex(r => String(r[0]) === key);
    if (idx >= 0) {
      sheet.getRange(idx + 2, 2).setValue(val);
    } else {
      sheet.appendRow([key, val]);
      rows.push([key, val]); // 同じリクエスト内で複数キーを追加する際の重複行を防ぐ
    }
  };

  const keys = ['hardMult', 'normalMult', 'easyMult', 'hellMult'];
  keys.forEach(key => {
    if (data[key] === undefined) return;
    const val = parseFloat(data[key]);
    if (isNaN(val)) return;
    setValue(key, val);
  });

  // お絵描きモードの週替わりお題の手動上書き（自由入力テキスト・空欄なら自動お題に戻る）
  if (data.theme !== undefined) {
    setValue('theme', String(data.theme).slice(0, 60));
  }

  return { ok: true, hardMult: data.hardMult, normalMult: data.normalMult, easyMult: data.easyMult, hellMult: data.hellMult, theme: data.theme };
}

// ----------------------------------------------------------------
// プレイ統計（軽量ログ・集計）
// ----------------------------------------------------------------
function logPlay(data) {
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('plays');
  if (!sheet) {
    sheet = ss.insertSheet('plays');
    sheet.appendRow(['日時', 'イベント', '難易度', 'スコア']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([new Date(), String(data.event || ''), String(data.difficulty || ''), Number(data.score) || 0]);
  return { ok: true };
}

function getStats() {
  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  const tz  = Session.getScriptTimeZone() || 'Asia/Tokyo';

  // 難易度選択・採点完了の回数
  const select   = { hard: 0, hell: 0, normal: 0, easy: 0 };
  const complete = { hard: 0, hell: 0, normal: 0, easy: 0 };
  const dailyMap = {}; // 'yyyy-MM-dd' -> 選択回数

  const playsSheet = ss.getSheetByName('plays');
  if (playsSheet && playsSheet.getLastRow() >= 2) {
    playsSheet.getRange(2, 1, playsSheet.getLastRow() - 1, 4).getValues().forEach(r => {
      const date  = r[0];
      const event = r[1];
      const diff  = r[2] || 'normal';
      if (event === 'select') {
        if (select[diff] !== undefined) select[diff]++;
        if (date instanceof Date) {
          const key = Utilities.formatDate(date, tz, 'yyyy-MM-dd');
          dailyMap[key] = (dailyMap[key] || 0) + 1;
        }
      } else if (event === 'complete') {
        if (complete[diff] !== undefined) complete[diff]++;
      }
    });
  }

  // 直近14日間（データが無い日は0件で埋める）
  const daily = [];
  for (let i = 13; i >= 0; i--) {
    const d = new Date();
    d.setDate(d.getDate() - i);
    const key = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    daily.push({ date: key, count: dailyMap[key] || 0 });
  }

  // 図鑑登録数（お友達・採点＝今月分+永久アーカイブ）
  let zukanFriend = 0;
  const drawingsSheet = ss.getSheetByName('drawings');
  if (drawingsSheet && drawingsSheet.getLastRow() >= 2) {
    zukanFriend = drawingsSheet.getRange(2, 8, drawingsSheet.getLastRow() - 1, 1).getValues()
      .filter(r => r[0] === '承認済み').length;
  }
  let zukanScoring = 0;
  const scoresSheet = ss.getSheetByName('scores');
  if (scoresSheet && scoresSheet.getLastRow() >= 2) {
    zukanScoring += scoresSheet.getRange(2, 7, scoresSheet.getLastRow() - 1, 1).getValues()
      .filter(r => r[0] === '承認済み').length;
  }
  const zukanArchSheet = ss.getSheetByName('scores_zukan_archive');
  if (zukanArchSheet && zukanArchSheet.getLastRow() >= 2) {
    zukanScoring += zukanArchSheet.getLastRow() - 1;
  }

  // 投票（総票数・投票された村人数）
  let voteTotal = 0, votedCharacters = 0;
  const votesSheet = ss.getSheetByName('votes');
  if (votesSheet && votesSheet.getLastRow() >= 2) {
    votesSheet.getRange(2, 3, votesSheet.getLastRow() - 1, 1).getValues().forEach(r => {
      const c = Number(r[0]) || 0;
      voteTotal += c;
      if (c > 0) votedCharacters++;
    });
  }

  return { select, complete, daily, zukanFriend, zukanScoring, voteTotal, votedCharacters };
}

// ----------------------------------------------------------------
// 過去に消えた鬼/地獄モード画像の復旧（管理画面用）
// scores・scores_zukan_archiveのどちらにも紐付いていないDrive内のscore_画像を一覧化し、
// 管理者が難易度を手動指定して図鑑に登録し直せるようにする
// ----------------------------------------------------------------
function getUnlinkedScoreImages() {
  const ss = SpreadsheetApp.openById(CONFIG.sheetId);
  const linkedIds = new Set();

  const extractId = (url) => {
    const m = String(url || '').match(/[-\w]{25,}/);
    return m ? m[0] : '';
  };

  const scoresSheet = ss.getSheetByName('scores');
  if (scoresSheet && scoresSheet.getLastRow() >= 2) {
    scoresSheet.getRange(2, 5, scoresSheet.getLastRow() - 1, 1).getValues()
      .forEach(r => { const id = extractId(r[0]); if (id) linkedIds.add(id); });
  }
  const zukanArchSheet = ss.getSheetByName('scores_zukan_archive');
  if (zukanArchSheet && zukanArchSheet.getLastRow() >= 2) {
    zukanArchSheet.getRange(2, 4, zukanArchSheet.getLastRow() - 1, 1).getValues()
      .forEach(r => { const id = extractId(r[0]); if (id) linkedIds.add(id); });
  }

  const folder = DriveApp.getFolderById(CONFIG.driveFolderId);
  const it = folder.getFiles();
  const found = [];
  while (it.hasNext()) {
    const f = it.next();
    const name = f.getName();
    if (!name.startsWith('score_')) continue;
    if (linkedIds.has(f.getId())) continue;
    const m = name.match(/^score_(.+)_(\d+)\.png$/);
    found.push({
      fileId:    f.getId(),
      name:      m ? m[1] : '',
      imageUrl:  'https://drive.google.com/thumbnail?id=' + f.getId() + '&sz=w300',
      createdAt: f.getDateCreated()
    });
  }
  found.sort((a, b) => b.createdAt - a.createdAt); // 新しい順（確認しやすい）
  return found.map(r => ({
    fileId:   r.fileId,
    name:     r.name,
    imageUrl: r.imageUrl,
    date:     r.createdAt.toLocaleDateString('ja-JP')
  }));
}

// 管理者が難易度・名前・スコアを指定して、孤立画像を図鑑（永久アーカイブ）に手動登録
function recoverScoreImage(data) {
  const fileId     = String(data.fileId || '');
  const name       = String(data.name || '名無し').slice(0, 20);
  const score      = Number(data.score) || 0;
  const difficulty = String(data.difficulty || '');
  if (!fileId)     return { error: 'fileId is required' };
  if (!difficulty) return { error: 'difficulty is required' };

  const file     = DriveApp.getFileById(fileId);
  const imageUrl = file.getUrl();

  const ss  = SpreadsheetApp.openById(CONFIG.sheetId);
  let sheet = ss.getSheetByName('scores_zukan_archive');
  if (!sheet) {
    sheet = ss.insertSheet('scores_zukan_archive');
    sheet.appendRow(['名前', 'スコア', '日付', '画像URL', '難易度', '採点レベル', '参加番号']);
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([name, score, file.getDateCreated().toLocaleDateString('ja-JP'), imageUrl, difficulty, 2, '']);
  return { ok: true };
}
