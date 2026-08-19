/**
 * =========================================================================
 * 劇団ハレトケ チケット予約・管理システム GAS (Google Apps Script)
 * =========================================================================
 * 
 * 【使い方】
 * 1. Googleスプレッドシートを作成し、「拡張機能」>「Apps Script」を開きます。
 * 2. このコードをすべて貼り付けて保存（Cmd+S / Ctrl+S）します。
 * 3. ツールバーの関数選択で「initialSetup」を選び「実行」をクリックします。
 *    （初回のみ権限の承認を求められます。「詳細」>「...に移動」で許可してください）
 *    → 必要なシート（設定、日程、券種、予約一覧）と初期データが自動生成されます！
 * 4. 右上の「デプロイ」>「新しいデプロイ」をクリックします。
 *    - 種類の選択：ウェブアプリ
 *    - 次のユーザーとして実行：自分 (your-account@gmail.com)
 *    - アクセスできるユーザー：全員 (Anyone)
 *    ※コードを更新した場合は、必ず「デプロイを管理」>「鉛筆アイコン（編集）」>「新バージョン」を選んで再デプロイしてください。
 * 5. 発行された「ウェブアプリURL」を ticket.html の GAS_ENDPOINT_URL に設定します。
 */

// ==========================================
// 1. 初回自動セットアップ関数
// ==========================================
function initialSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ① 設定シート
  let settingSheet = ss.getSheetByName('設定');
  if (!settingSheet) {
    settingSheet = ss.insertSheet('設定');
    settingSheet.appendRow(['設定項目', '設定値', '説明']);
    settingSheet.appendRow(['公演名', '劇団ハレトケ 2026年 秋公演', 'Webページの見出しに表示される公演名']);
    settingSheet.appendRow(['サブタイトル', '〜いま、ここ、こそ、パラダイス〜', '公演のキャッチコピーやサブタイトル']);
    settingSheet.appendRow(['会場名', '〇〇劇場', '会場名（例：JOY JOY THEATER、〇〇ホール等）']);
    settingSheet.appendRow(['会場URL', 'https://example.com/venue', '会場の公式サイト・Googleマップ等のURL']);
    settingSheet.appendRow(['会場住所・アクセス', '東京都〇〇区...（〇〇駅 徒歩5分）', 'アクセス情報']);
    settingSheet.appendRow(['チケットページURL', 'https://corgiii-w7.github.io/Haretoke--web/ticket.html', 'キャンセル用URL生成に使用するWebページのURL']);
    settingSheet.appendRow(['販売ステータス', '自動', '「自動」「販売中」「予約開始前」「販売停止/終了」のいずれか']);
    settingSheet.appendRow(['販売開始日時', '2026-09-01 12:00', '全公演回共通の予約開始日時 (YYYY-MM-DD HH:mm 形式)']);
    settingSheet.appendRow(['公演説明・あらすじ', '次回公演に向けて鋭意稽古中です！団員一同、劇場でお待ちしております。', '公演のあらすじや概要']);
    settingSheet.appendRow(['注意事項・備考', '・開場は各回開演の30分前です。\n・未就学児のご入場はご遠慮いただいております。\n・当日受付にてご予約名をお伝えいただき、代金をご精算ください。', 'チケットに関する注意事項']);
    settingSheet.appendRow(['自動返信メール送信', '有効', '「有効」または「無効」']);
    settingSheet.appendRow(['自動返信メール件名', '【劇団ハレトケ】チケットご予約完了のお知らせ（予約番号: {予約番号}）', '予約者へ自動送信するメールの件名']);
    settingSheet.appendRow(['自動返信メール本文', 
`{お名前} 様

この度は『{公演名}』のチケットをご予約いただき、誠にありがとうございます。
以下の内容でご予約を承りました。

━━━━━━━━━━━━━━━━━━━━━━━━
【予約番号】: {予約番号}
【お名前】: {お名前} 様
【公演日時】: {公演日時}
【券種】: {券種}
【予約枚数】: {枚数}枚
【合計金額】: {合計金額} (当日受付精算)
【会場】: {会場名}
━━━━━━━━━━━━━━━━━━━━━━━━

【ご案内・注意事項】
※予約時に振り分けられる番号は、チケット管理のものであり、それによって座席を指定するものではございません。
※公演時間の変更をご希望の場合は、一度チケットをキャンセルいただき、新しくご希望のチケットをご予約ください。
{注意事項}

━━━━━━━━━━━━━━━━━━━━━━━━
■ ご予約のキャンセル・枚数変更について
枚数の変更（例：5枚から4枚に減らす等）やキャンセルは、以下のURLよりいつでもお手続きいただけます。
{キャンセルURL}
━━━━━━━━━━━━━━━━━━━━━━━━

当日、劇場にてお会いできることを団員一同心より楽しみにしております！

━━━━━━━━━━━━━━━━━━━━━━━━
劇団ハレトケ 公式サイト
https://theatrical.net-menber.com/
━━━━━━━━━━━━━━━━━━━━━━━━`, '予約完了メールのテンプレート（{タグ}が自動置換されます）']);

    settingSheet.appendRow(['キャンセル通知メール件名', '【劇団ハレトケ】チケットキャンセル手続き完了のお知らせ（予約番号: {予約番号}）', 'キャンセル完了時に送信するメール件名']);
    settingSheet.appendRow(['キャンセル通知メール本文',
`{お名前} 様

『{公演名}』のチケットキャンセル手続きを承りました。

━━━━━━━━━━━━━━━━━━━━━━━━
【予約番号】: {予約番号}
【公演日時】: {公演日時}
【券種】: {券種}
【キャンセル枚数】: {キャンセル枚数}枚
【現在の有効枚数】: {有効枚数}枚
【合計金額】: {合計金額}
━━━━━━━━━━━━━━━━━━━━━━━━

※予約時に振り分けられる番号は、チケット管理のものであり、それによって座席を指定するものではございません。
※公演時間の変更をご希望の場合は、一度チケットをキャンセルいただき、新しくご希望のチケットをご予約ください。
※残りのチケットの確認・キャンセルは以下のURLよりいつでも可能です。
{キャンセルURL}

またのご来場を団員一同心よりお待ちしております。

━━━━━━━━━━━━━━━━━━━━━━━━
劇団ハレトケ 公式サイト
https://theatrical.net-menber.com/
━━━━━━━━━━━━━━━━━━━━━━━━`, 'キャンセル完了メールのテンプレート']);

    // ヘッダースタイル
    settingSheet.getRange('A1:C1').setBackground('#222222').setFontColor('#ffffff').setFontWeight('bold');
    settingSheet.setColumnWidth(1, 180);
    settingSheet.setColumnWidth(2, 400);
    settingSheet.setColumnWidth(3, 300);
  }

  // ② 日程シート（公演時間ごとの販売終了日時を設定可能）
  let scheduleSheet = ss.getSheetByName('日程');
  if (!scheduleSheet) {
    scheduleSheet = ss.insertSheet('日程');
    scheduleSheet.appendRow(['公演日時', '残席ステータス', '販売終了日時', '備考/開場時間']);
    scheduleSheet.appendRow(['2026年10月24日(土) 14:00', '受付中', '2026-10-24 12:00', '開場 13:30']);
    scheduleSheet.appendRow(['2026年10月24日(土) 18:30', '受付中', '2026-10-24 16:30', '開場 18:00']);
    scheduleSheet.appendRow(['2026年10月25日(日) 13:00', '残りわずか', '2026-10-25 11:00', '開場 12:30']);
    
    scheduleSheet.getRange('A1:D1').setBackground('#222222').setFontColor('#ffffff').setFontWeight('bold');
    scheduleSheet.setColumnWidth(1, 240);
    scheduleSheet.setColumnWidth(2, 120);
    scheduleSheet.setColumnWidth(3, 160);
    scheduleSheet.setColumnWidth(4, 200);
  }

  // ③ 券種シート
  let ticketTypeSheet = ss.getSheetByName('券種');
  if (!ticketTypeSheet) {
    ticketTypeSheet = ss.insertSheet('券種');
    ticketTypeSheet.appendRow(['券種名', '料金(円)', '説明・対象']);
    ticketTypeSheet.appendRow(['一般', 3000, '一般前売券']);
    ticketTypeSheet.appendRow(['U-25', 2000, '25歳以下対象（当日要身分証明書）']);
    ticketTypeSheet.appendRow(['応援プレミアム', 4500, '前方良席指定 ＋ オリジナル特典付き']);
    
    ticketTypeSheet.getRange('A1:C1').setBackground('#222222').setFontColor('#ffffff').setFontWeight('bold');
    ticketTypeSheet.setColumnWidth(1, 160);
    ticketTypeSheet.setColumnWidth(2, 120);
    ticketTypeSheet.setColumnWidth(3, 280);
  }

  // ④ 予約一覧シート（1行＝1チケット管理）
  let resSheet = ss.getSheetByName('予約一覧');
  if (!resSheet) {
    resSheet = ss.insertSheet('予約一覧');
    resSheet.appendRow([
      '予約日時',
      '予約番号',
      'チケット番号',
      'キャンセルキー',
      'ステータス',
      'お名前',
      'フリガナ',
      'メールアドレス',
      '公演日時',
      '券種',
      '金額(円)',
      '知ったきっかけ',
      '備考・メッセージ',
      'キャンセル日時'
    ]);
    resSheet.getRange('A1:N1').setBackground('#111111').setFontColor('#ffffff').setFontWeight('bold');
    resSheet.setColumnWidth(1, 160);
    resSheet.setColumnWidth(2, 140);
    resSheet.setColumnWidth(3, 150);
    resSheet.setColumnWidth(4, 150);
    resSheet.setColumnWidth(5, 100);
    resSheet.setColumnWidth(6, 120);
    resSheet.setColumnWidth(7, 120);
    resSheet.setColumnWidth(8, 200);
    resSheet.setColumnWidth(9, 200);
    resSheet.setColumnWidth(10, 120);
    resSheet.setColumnWidth(11, 90);
    resSheet.setColumnWidth(12, 150);
    resSheet.setColumnWidth(13, 220);
    resSheet.setColumnWidth(14, 160);
  }

  // 不要な初期シート「シート1」があれば削除
  const defaultSheet = ss.getSheetByName('シート1');
  if (defaultSheet && ss.getSheets().length > 1) {
    ss.deleteSheet(defaultSheet);
  }

  Logger.log('初期セットアップが完了しました！');
}

// ==========================================
// 2. GETリクエスト処理 (Webサイトへのデータ提供)
// ==========================================
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const action = e && e.parameter && e.parameter.action;

    // --- A. 予約照会（キャンセル画面用：複数枚チケット情報） ---
    if (action === 'getReservation') {
      const targetId = String(e.parameter.id || '').trim();
      const targetKey = String(e.parameter.key || '').trim();

      if (!targetId || !targetKey) {
        return sendJsonResponse({
          status: 'error',
          message: '予約番号またはキャンセルキーが指定されていません。'
        }, e);
      }

      const resSheet = ss.getSheetByName('予約一覧');
      if (!resSheet) {
        return sendJsonResponse({ status: 'error', message: '予約データが見つかりません。' }, e);
      }

      const resValues = resSheet.getDataRange().getValues();
      const matchedTickets = [];
      let customerName = '';
      let customerFurigana = '';
      let customerEmail = '';
      let scheduleDate = '';
      let ticketType = '';
      let unitPrice = 0;
      let source = '';
      let remarks = '';

      for (let i = 1; i < resValues.length; i++) {
        const rowId = String(resValues[i][1] || '').trim();
        const rowKey = String(resValues[i][3] || '').trim();

        if (rowId === targetId && rowKey === targetKey) {
          customerName = String(resValues[i][5] || '');
          customerFurigana = String(resValues[i][6] || '');
          customerEmail = String(resValues[i][7] || '');
          scheduleDate = String(resValues[i][8] || '');
          ticketType = String(resValues[i][9] || '');
          unitPrice = Number(resValues[i][10]) || 0;
          source = String(resValues[i][11] || '');
          remarks = String(resValues[i][12] || '');

          matchedTickets.push({
            ticketNumber: String(resValues[i][2] || ''),
            status: String(resValues[i][4] || '有効').trim(),
            cancelTime: String(resValues[i][13] || '')
          });
        }
      }

      if (matchedTickets.length === 0) {
        return sendJsonResponse({
          status: 'error',
          message: 'ご指定の予約情報が見つからないか、認証キーが一致しません。'
        }, e);
      }

      const activeTickets = matchedTickets.filter(t => t.status === '有効');
      const cancelledTickets = matchedTickets.filter(t => t.status === 'キャンセル済み');

      const settings = readSettings(ss);

      const reservationData = {
        id: targetId,
        key: targetKey,
        name: customerName,
        furigana: customerFurigana,
        email: customerEmail,
        date: scheduleDate,
        ticketType: ticketType,
        unitPrice: unitPrice,
        totalCount: matchedTickets.length,
        activeCount: activeTickets.length,
        cancelledCount: cancelledTickets.length,
        activeTotalPrice: `¥${(activeTickets.length * unitPrice).toLocaleString()}`,
        isFullyCancelled: activeTickets.length === 0,
        tickets: matchedTickets,
        source: source,
        remarks: remarks,
        eventTitle: settings['公演名'] || '劇団ハレトケ 公演',
        venue: settings['会場名'] || ''
      };

      return sendJsonResponse({
        status: 'success',
        reservation: reservationData
      }, e);
    }

    // --- B. 通常の公演・チケット情報取得 ---
    const now = new Date();
    const settings = readSettings(ss);

    // 日程シートの読み取り
    const scheduleSheet = ss.getSheetByName('日程');
    const schedules = [];
    let availableScheduleCount = 0;

    if (scheduleSheet) {
      const schValues = scheduleSheet.getDataRange().getValues();
      if (schValues.length > 1) {
        const headerRow = schValues[0].map(h => String(h || '').trim());
        let dtCol = 0;
        let stCol = 1;
        let endCol = -1;
        let noteCol = -1;

        // ヘッダー名から列位置を厳密に特定（「終了」を優先判定）
        headerRow.forEach((h, idx) => {
          if (h.includes('終了') || h.includes('締切')) {
            endCol = idx;
          } else if (h.includes('公演') || (h.includes('日時') && !h.includes('終了'))) {
            dtCol = idx;
          } else if (h.includes('ステータス') || h.includes('残席')) {
            stCol = idx;
          } else if (h.includes('開場') || h.includes('備考') || h.includes('注')) {
            noteCol = idx;
          }
        });

        // 未設定時のデフォルト列インデックス
        if (endCol === -1) endCol = 2;
        if (noteCol === -1) noteCol = 3;

        for (let i = 1; i < schValues.length; i++) {
          let rawDatetime = schValues[i][dtCol];
          let datetime = '';
          if (rawDatetime instanceof Date) {
            datetime = formatDate(rawDatetime);
          } else {
            datetime = formatDateString(String(rawDatetime || '').trim());
          }

          let status = String(schValues[i][stCol] || '受付中').trim();
          
          let valEnd = schValues[i][endCol];
          let valNote = schValues[i][noteCol];

          let slotEndDate = null;
          let note = '';

          // どちらが日付でどちらが備考かをスマートに判定
          if (parseDateValue(valEnd)) {
            slotEndDate = parseDateValue(valEnd);
            note = (valNote instanceof Date) ? formatDate(valNote) : String(valNote || '').trim();
          } else if (parseDateValue(valNote)) {
            slotEndDate = parseDateValue(valNote);
            note = (valEnd instanceof Date) ? formatDate(valEnd) : String(valEnd || '').trim();
          } else {
            note = String(valNote || valEnd || '').trim();
          }

          // noteに英語の日時文字列が入っていた場合は日本語化または除外
          if (note.includes('GMT') || note.includes('日本標準時') || note.includes('Standard Time')) {
            const parsedNoteDate = parseDateValue(note);
            if (parsedNoteDate) {
              note = ''; // 販売終了日時と重複している場合はクリア
            }
          }

          if (datetime) {
            if (slotEndDate && now.getTime() > slotEndDate.getTime()) {
              status = '受付終了';
            }

            if (status !== '完売' && status !== '受付終了' && status !== '販売終了') {
              availableScheduleCount++;
            }

            schedules.push({
              datetime: datetime,
              status: status,
              salesEndDate: slotEndDate ? formatDate(slotEndDate) : (valEnd ? formatDateString(String(valEnd).trim()) : ''),
              note: note
            });
          }
        }
      }
    }

    // 券種シートの読み取り
    const ticketSheet = ss.getSheetByName('券種');
    const ticketTypes = [];
    if (ticketSheet) {
      const tValues = ticketSheet.getDataRange().getValues();
      for (let i = 1; i < tValues.length; i++) {
        const name = String(tValues[i][0] || '').trim();
        const price = Number(tValues[i][1]) || 0;
        const description = String(tValues[i][2] || '').trim();
        if (name) {
          ticketTypes.push({ name, price, description });
        }
      }
    }

    // 販売ステータスの自動判定
    const mode = String(settings['販売ステータス'] || '自動').trim();
    let currentStatus = 'closed';

    const rawStartDate = settings['販売開始日時'];
    const startDate = parseDateValue(rawStartDate);

    if (mode === '販売中' || mode === '受付中') {
      currentStatus = 'active';
    } else if (mode === '予約開始前') {
      currentStatus = 'scheduled';
    } else if (mode === '販売停止/終了' || mode === '売出なし' || mode === '受付停止') {
      currentStatus = 'closed';
    } else {
      if (startDate && now.getTime() < startDate.getTime()) {
        currentStatus = 'scheduled';
      } else if (schedules.length === 0 || availableScheduleCount === 0) {
        currentStatus = 'closed';
      } else {
        currentStatus = 'active';
      }
    }

    const payload = {
      status: 'success',
      data: {
        salesStatus: currentStatus,
        title: settings['公演名'] ? String(settings['公演名']) : '劇団ハレトケ 公演',
        subtitle: settings['サブタイトル'] ? String(settings['サブタイトル']) : '',
        venue: settings['会場名'] ? String(settings['会場名']) : '',
        venueUrl: settings['会場URL'] ? String(settings['会場URL']) : (settings['会場詳細URL'] ? String(settings['会場詳細URL']) : (settings['会場詳細'] ? String(settings['会場詳細']) : '')),
        venueAccess: settings['会場住所・アクセス'] ? String(settings['会場住所・アクセス']) : '',
        salesStartDate: startDate ? formatDate(startDate) : (rawStartDate ? String(rawStartDate).trim() : ''),
        description: settings['公演説明・あらすじ'] ? String(settings['公演説明・あらすじ']) : '',
        notes: settings['注意事項・備考'] ? String(settings['注意事項・備考']) : '',
        schedules: schedules,
        ticketTypes: ticketTypes
      }
    };

    return sendJsonResponse(payload, e);

  } catch (error) {
    return sendJsonResponse({
      status: 'error',
      message: error.toString()
    }, e);
  }
}

// ==========================================
// 3. POSTリクエスト処理 (予約受付 & n枚キャンセル処理)
// ==========================================
function doPost(e) {
  try {
    let data = {};
    if (e.postData && e.postData.contents) {
      data = JSON.parse(e.postData.contents);
    } else {
      data = e.parameter || {};
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let resSheet = ss.getSheetByName('予約一覧');
    if (!resSheet) {
      initialSetup();
      resSheet = ss.getSheetByName('予約一覧');
    }

    const now = new Date();
    const settings = readSettings(ss);
    const eventTitle = settings['公演名'] || '劇団ハレトケ 公演';
    const venue = settings['会場名'] || '';
    const venueUrl = settings['会場URL'] || settings['会場詳細URL'] || settings['会場詳細'] || '';
    const notes = settings['注意事項・備考'] || '';
    const emailEnabled = String(settings['自動返信メール送信'] || '有効').trim() !== '無効';

    // =========================================================
    // A. 予約キャンセル実行 (action: 'cancel', n枚キャンセル対応)
    // =========================================================
    if (data.action === 'cancel') {
      const targetId = String(data.id || '').trim();
      const targetKey = String(data.key || '').trim();
      // キャンセル枚数（指定がない場合は全有効枚数）
      let cancelCountRequested = Number(data.cancelCount) || 0;

      if (!targetId || !targetKey) {
        return sendJsonResponse({ status: 'error', message: '予約番号またはキャンセルキーが不足しています。' }, e);
      }

      const resValues = resSheet.getDataRange().getValues();
      const activeRowIndexes = [];
      let resData = null;

      for (let i = 1; i < resValues.length; i++) {
        const rowId = String(resValues[i][1] || '').trim();
        const rowKey = String(resValues[i][3] || '').trim();
        const rowStatus = String(resValues[i][4] || '').trim();

        if (rowId === targetId && rowKey === targetKey) {
          if (!resData) {
            resData = {
              id: rowId,
              name: String(resValues[i][5] || ''),
              email: String(resValues[i][7] || ''),
              date: String(resValues[i][8] || ''),
              ticketType: String(resValues[i][9] || ''),
              unitPrice: Number(resValues[i][10]) || 0
            };
          }

          if (rowStatus === '有効') {
            activeRowIndexes.push(i + 1); // 1-indexed sheet row
          }
        }
      }

      if (!resData) {
        return sendJsonResponse({ status: 'error', message: '該当する予約が見つかりませんでした。' }, e);
      }

      if (activeRowIndexes.length === 0) {
        return sendJsonResponse({ status: 'already_cancelled', message: 'この予約のチケットはすべて既にキャンセルされています。' }, e);
      }

      // キャンセル枚数の決定（指定が0または有効数以上の場合は全有効枚数）
      let actualCancelCount = cancelCountRequested;
      if (actualCancelCount <= 0 || actualCancelCount > activeRowIndexes.length) {
        actualCancelCount = activeRowIndexes.length;
      }

      // 後ろのチケットから順に actualCancelCount 枚をキャンセル
      const cancelTimeStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
      for (let k = 0; k < actualCancelCount; k++) {
        const targetRow = activeRowIndexes[activeRowIndexes.length - 1 - k];
        resSheet.getRange(targetRow, 5).setValue('キャンセル済み'); // Column 5: ステータス
        resSheet.getRange(targetRow, 14).setValue(cancelTimeStr);  // Column 14: キャンセル日時
      }

      const remainingActiveCount = activeRowIndexes.length - actualCancelCount;
      const remainingTotalPrice = `¥${(remainingActiveCount * resData.unitPrice).toLocaleString()}`;

      // キャンセルURL
      let siteUrl = String(settings['チケットページURL'] || 'https://corgiii-w7.github.io/Haretoke--web/ticket.html').trim();
      const sep = siteUrl.includes('?') ? '&' : '?';
      const cancelUrl = `${siteUrl}${sep}action=cancel&id=${encodeURIComponent(targetId)}&key=${encodeURIComponent(targetKey)}`;

      // キャンセル完了メールの送信
      if (emailEnabled && resData.email && resData.email.includes('@')) {
        let cancelSubject = settings['キャンセル通知メール件名'] || '【劇団ハレトケ】チケットキャンセル手続き完了のお知らせ（予約番号: {予約番号}）';
        let cancelBody = settings['キャンセル通知メール本文'] || `{お名前} 様\n\n『{公演名}』のチケットキャンセル手続きを承りました。\n\n【予約番号】: {予約番号}\n【公演日時】: {公演日時}\n【券種】: {券種}\n【キャンセル枚数】: {キャンセル枚数}枚\n【現在の有効枚数】: {有効枚数}枚\n【合計金額】: {合計金額}\n\n※公演時間の変更をご希望の場合は、一度チケットをキャンセルいただき、新しくご希望のチケットをご予約ください。\n※予約時に振り分けられる番号は、チケット管理のものであり、座席を指定するものではございません。\n\n{キャンセルURL}`;

        const placeholders = {
          お名前: resData.name,
          予約番号: resData.id,
          キャンセルキー: targetKey,
          キャンセルURL: cancelUrl,
          公演名: eventTitle,
          公演日時: resData.date,
          券種: resData.ticketType,
          キャンセル枚数: actualCancelCount,
          有効枚数: remainingActiveCount,
          枚数: remainingActiveCount,
          合計金額: remainingTotalPrice,
          会場名: venue,
          注意事項: notes
        };

        cancelSubject = replacePlaceholders(cancelSubject, placeholders);
        cancelBody = replacePlaceholders(cancelBody, placeholders);

        try {
          GmailApp.sendEmail(resData.email, cancelSubject, cancelBody, {
            name: '劇団ハレトケ チケット窓口'
          });
        } catch (mailErr) {
          Logger.log('キャンセルメール送信エラー: ' + mailErr);
        }
      }

      return sendJsonResponse({
        status: 'success',
        cancelledCount: actualCancelCount,
        remainingActiveCount: remainingActiveCount,
        remainingTotalPrice: remainingTotalPrice,
        isFullyCancelled: remainingActiveCount === 0,
        message: `${actualCancelCount}枚のチケットキャンセルが完了いたしました。`
      }, e);
    }

    // =========================================================
    // B. 新規チケット予約受付 (1枚につき1行作成)
    // =========================================================
    const dateCode = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
    const randomCode = Math.floor(1000 + Math.random() * 9000);
    const reservationId = `HT-${dateCode}-${randomCode}`;
    const cancelKey = 'ck_' + Utilities.getUuid().replace(/-/g, '').substring(0, 16);

    const reservationTime = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const name = String(data.name || '').trim();
    const furigana = String(data.furigana || '').trim();
    const email = String(data.email || '').trim();
    const scheduleDate = String(data.date || '').trim();
    const ticketType = String(data.ticketType || '').trim();
    const count = Math.max(1, Number(data.count) || 1);
    const unitPrice = Number(data.unitPrice) || 0;
    const totalPriceText = data.totalPrice || `¥${(count * unitPrice).toLocaleString()}`;
    const source = String(data.source || data.hearing || '').trim();
    const remarks = String(data.remarks || '').trim();

    // キャンセルURL
    let siteUrl = String(settings['チケットページURL'] || 'https://corgiii-w7.github.io/Haretoke--web/ticket.html').trim();
    const sep = siteUrl.includes('?') ? '&' : '?';
    const cancelUrl = `${siteUrl}${sep}action=cancel&id=${encodeURIComponent(reservationId)}&key=${encodeURIComponent(cancelKey)}`;

    // 予約枚数分、各チケットを1行ずつスプレッドシートに追加
    for (let c = 1; c <= count; c++) {
      const ticketNumber = `${reservationId}-${c}`;
      resSheet.appendRow([
        reservationTime,
        reservationId,
        ticketNumber,
        cancelKey,
        '有効',
        name,
        furigana,
        email,
        scheduleDate,
        ticketType,
        unitPrice,
        source,
        remarks,
        ''
      ]);
    }

    // 自動返信メールの送信
    if (emailEnabled && email && email.includes('@')) {
      let emailSubject = settings['自動返信メール件名'] || '【劇団ハレトケ】チケットご予約完了のお知らせ（予約番号: {予約番号}）';
      let emailBody = settings['自動返信メール本文'] || `{お名前} 様\n\nこの度は『{公演名}』のチケットをご予約いただき、誠にありがとうございます。\n\n【予約番号】: {予約番号}\n【公演日時】: {公演日時}\n【券種】: {券種}\n【枚数】: {枚数}枚\n【合計金額】: {合計金額}\n\n※予約時に振り分けられる番号は、チケット管理のものであり、座席を指定するものではございません。\n※公演時間の変更をご希望の場合は、一度チケットをキャンセルいただき、新しくご希望のチケットをご予約ください。\n\n■ キャンセル・枚数変更URL:\n{キャンセルURL}`;

      const placeholders = {
        お名前: name,
        フリガナ: furigana,
        予約番号: reservationId,
        キャンセルキー: cancelKey,
        キャンセルURL: cancelUrl,
        公演名: eventTitle,
        公演日時: scheduleDate,
        券種: ticketType,
        枚数: count,
        合計金額: totalPriceText,
        知ったきっかけ: source,
        会場名: venue,
        会場URL: venueUrl,
        注意事項: notes,
        備考: remarks
      };

      emailSubject = replacePlaceholders(emailSubject, placeholders);
      emailBody = replacePlaceholders(emailBody, placeholders);

      try {
        GmailApp.sendEmail(email, emailSubject, emailBody, {
          name: '劇団ハレトケ チケット窓口'
        });
      } catch (mailErr) {
        Logger.log('メール送信エラー: ' + mailErr);
      }
    }

    return sendJsonResponse({
      status: 'success',
      reservationId: reservationId,
      cancelKey: cancelKey,
      cancelUrl: cancelUrl,
      count: count,
      message: '予約が正常に完了しました'
    }, e);

  } catch (error) {
    return sendJsonResponse({
      status: 'error',
      message: error.toString()
    }, e);
  }
}

// ==========================================
// ヘルパー関数群
// ==========================================

// 設定シートを一括読込してオブジェクト化
function readSettings(ss) {
  const settingSheet = ss.getSheetByName('設定');
  const settings = {};
  if (settingSheet) {
    const settingValues = settingSheet.getDataRange().getValues();
    for (let i = 1; i < settingValues.length; i++) {
      const key = String(settingValues[i][0] || '').trim();
      const val = settingValues[i][1];
      if (key) {
        settings[key] = val;
      }
    }
  }
  return settings;
}

// メール本文中の {タグ} を置換
function replacePlaceholders(text, map) {
  if (!text) return '';
  let result = String(text);
  for (const [k, v] of Object.entries(map)) {
    const reg = new RegExp(`\\{${k}\\}`, 'g');
    result = result.replace(reg, v !== undefined && v !== null ? String(v) : '');
  }
  return result;
}

// JSON / JSONP レスポンス生成
function sendJsonResponse(payload, e) {
  const jsonString = JSON.stringify(payload);
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + jsonString + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonString)
    .setMimeType(ContentService.MimeType.JSON);
}

// 日時値パーサー（Date型オブジェクトと文字列の両方に対応）
function parseDateValue(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return isNaN(val.getTime()) ? null : val;
  }
  if (typeof val === 'string') {
    const str = val.trim();
    if (!str) return null;
    // YYYY/MM/DD や YYYY-MM-DD
    const d = new Date(str.replace(/\//g, '-'));
    if (!isNaN(d.getTime())) return d;
    // JSのtoString()形式（例: Sat Sep 26 2026...）
    const d2 = new Date(str);
    if (!isNaN(d2.getTime())) return d2;
  }
  return null;
}

// 日時フォーマット用ヘルパー（日本語表記）
function formatDate(d) {
  if (!d) return '';
  let dateObj = d;
  if (typeof d === 'string') {
    dateObj = parseDateValue(d);
    if (!dateObj) return d;
  }
  if (dateObj instanceof Date && !isNaN(dateObj.getTime())) {
    const year = dateObj.getFullYear();
    const month = dateObj.getMonth() + 1;
    const date = dateObj.getDate();
    const hours = ('0' + dateObj.getHours()).slice(-2);
    const minutes = ('0' + dateObj.getMinutes()).slice(-2);
    const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][dateObj.getDay()];
    return `${year}年${month}月${date}日(${dayOfWeek}) ${hours}:${minutes}`;
  }
  return String(d);
}

// 文字列が英語日時だった場合の日本語フォーマット
function formatDateString(str) {
  if (!str) return '';
  if (str.includes('GMT') || str.includes('日本標準時') || str.includes('Standard Time') || /^[A-Za-z]{3}\s+[A-Za-z]{3}/.test(str)) {
    const parsed = parseDateValue(str);
    if (parsed) {
      return formatDate(parsed);
    }
  }
  return str;
}
