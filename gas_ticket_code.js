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
    settingSheet.appendRow(['会場住所・アクセス', '東京都〇〇区...（〇〇駅 徒歩5分）', 'アクセス情報']);
    settingSheet.appendRow(['販売ステータス', '自動', '「自動」「販売中」「予約開始前」「販売停止/終了」のいずれか']);
    settingSheet.appendRow(['販売開始日時', '2026-09-01 12:00', '全公演回共通の予約開始日時 (YYYY-MM-DD HH:mm 形式)']);
    settingSheet.appendRow(['公演説明・あらすじ', '次回公演に向けて鋭意稽古中です！団員一同、劇場でお待ちしております。', '公演のあらすじや概要']);
    settingSheet.appendRow(['注意事項・備考', '・開場は各回開演の30分前です。\n・未就学児のご入場はご遠慮いただいております。\n・当日受付にてご予約名をお伝えいただき、代金をご精算ください。', 'チケットに関する注意事項']);
    settingSheet.appendRow(['自動返信メール件名', '【劇団ハレトケ】チケットご予約完了のお知らせ', '予約者へ自動送信するメールの件名']);
    settingSheet.appendRow(['自動返信メール送信', '有効', '「有効」または「無効」']);
    
    // ヘッダースタイル
    settingSheet.getRange('A1:C1').setBackground('#222222').setFontColor('#ffffff').setFontWeight('bold');
    settingSheet.setColumnWidth(1, 160);
    settingSheet.setColumnWidth(2, 320);
    settingSheet.setColumnWidth(3, 280);
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

  // ④ 予約一覧シート
  let resSheet = ss.getSheetByName('予約一覧');
  if (!resSheet) {
    resSheet = ss.insertSheet('予約一覧');
    resSheet.appendRow([
      '予約日時',
      '予約番号',
      'お名前',
      'フリガナ',
      'メールアドレス',
      '電話番号',
      '公演日時',
      '券種',
      '枚数',
      '合計金額',
      '知ったきっかけ',
      '備考・メッセージ'
    ]);
    resSheet.getRange('A1:L1').setBackground('#111111').setFontColor('#ffffff').setFontWeight('bold');
    resSheet.setColumnWidth(1, 160);
    resSheet.setColumnWidth(2, 140);
    resSheet.setColumnWidth(3, 120);
    resSheet.setColumnWidth(4, 120);
    resSheet.setColumnWidth(5, 200);
    resSheet.setColumnWidth(6, 130);
    resSheet.setColumnWidth(7, 200);
    resSheet.setColumnWidth(8, 120);
    resSheet.setColumnWidth(9, 80);
    resSheet.setColumnWidth(10, 100);
    resSheet.setColumnWidth(11, 150);
    resSheet.setColumnWidth(12, 220);
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
    const now = new Date();
    
    // 設定シートの読み取り
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

    // 日程シートの読み取り（公演ごとの販売終了日時判定）
    const scheduleSheet = ss.getSheetByName('日程');
    const schedules = [];
    let availableScheduleCount = 0;

    if (scheduleSheet) {
      const schValues = scheduleSheet.getDataRange().getValues();
      for (let i = 1; i < schValues.length; i++) {
        const datetime = String(schValues[i][0] || '').trim();
        let status = String(schValues[i][1] || '受付中').trim();
        const rawEndDate = schValues[i][2];
        const slotEndDate = parseDateValue(rawEndDate);
        const note = String(schValues[i][3] || '').trim();

        if (datetime) {
          // 公演ごとの販売終了日時の判定
          if (slotEndDate && now.getTime() > slotEndDate.getTime()) {
            status = '受付終了';
          }

          if (status !== '完売' && status !== '受付終了' && status !== '販売終了') {
            availableScheduleCount++;
          }

          schedules.push({
            datetime: datetime,
            status: status,
            salesEndDate: slotEndDate ? formatDate(slotEndDate) : (rawEndDate ? String(rawEndDate).trim() : ''),
            note: note
          });
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
    let currentStatus = 'closed'; // 'active' (販売中), 'scheduled' (予約開始前), 'closed' (売出なし/終了)

    const rawStartDate = settings['販売開始日時'];
    const startDate = parseDateValue(rawStartDate);

    if (mode === '販売中' || mode === '受付中') {
      currentStatus = 'active';
    } else if (mode === '予約開始前') {
      currentStatus = 'scheduled';
    } else if (mode === '販売停止/終了' || mode === '売出なし' || mode === '受付停止') {
      currentStatus = 'closed';
    } else {
      // 「自動」判定
      if (startDate && now.getTime() < startDate.getTime()) {
        currentStatus = 'scheduled';
      } else if (schedules.length === 0 || availableScheduleCount === 0) {
        currentStatus = 'closed'; // 有効な日程がすべて受付終了または完売
      } else {
        currentStatus = 'active';
      }
    }

    const payload = {
      status: 'success',
      data: {
        salesStatus: currentStatus, // 'active' | 'scheduled' | 'closed'
        title: settings['公演名'] ? String(settings['公演名']) : '劇団ハレトケ 公演',
        subtitle: settings['サブタイトル'] ? String(settings['サブタイトル']) : '',
        venue: settings['会場名'] ? String(settings['会場名']) : '',
        venueAccess: settings['会場住所・アクセス'] ? String(settings['会場住所・アクセス']) : '',
        salesStartDate: startDate ? formatDate(startDate) : (rawStartDate ? String(rawStartDate).trim() : ''),
        description: settings['公演説明・あらすじ'] ? String(settings['公演説明・あらすじ']) : '',
        notes: settings['注意事項・備考'] ? String(settings['注意事項・備考']) : '',
        schedules: schedules,
        ticketTypes: ticketTypes
      }
    };

    const jsonString = JSON.stringify(payload);
    
    // JSONP callbackサポート
    const callback = e && e.parameter && e.parameter.callback;
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + jsonString + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    return ContentService.createTextOutput(jsonString)
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    const errorPayload = {
      status: 'error',
      message: error.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(errorPayload))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================
// 3. POSTリクエスト処理 (予約受付と記録)
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

    // 予約番号の自動生成 (例: HT-20261024-1234)
    const now = new Date();
    const dateCode = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyyMMdd');
    const randomCode = Math.floor(1000 + Math.random() * 9000);
    const reservationId = `HT-${dateCode}-${randomCode}`;

    const reservationTime = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
    const name = data.name || '';
    const furigana = data.furigana || '';
    const email = data.email || '';
    const phone = data.phone || '';
    const scheduleDate = data.date || '';
    const ticketType = data.ticketType || '';
    const count = Number(data.count) || 1;
    const totalPrice = data.totalPrice || '';
    const source = data.source || data.hearing || data.cast || '';
    const remarks = data.remarks || '';

    // スプレッドシートに追記
    resSheet.appendRow([
      reservationTime,
      reservationId,
      name,
      furigana,
      email,
      phone,
      scheduleDate,
      ticketType,
      count,
      totalPrice,
      source,
      remarks
    ]);

    // 設定シートからメール設定を取得
    const settingSheet = ss.getSheetByName('設定');
    let emailSubject = '【劇団ハレトケ】チケットご予約完了のお知らせ';
    let emailEnabled = true;
    let eventTitle = '劇団ハレトケ 公演';
    let venue = '';
    let notes = '';

    if (settingSheet) {
      const settingValues = settingSheet.getDataRange().getValues();
      for (let i = 1; i < settingValues.length; i++) {
        const k = String(settingValues[i][0] || '').trim();
        const v = settingValues[i][1];
        if (k === '自動返信メール件名' && v) emailSubject = String(v);
        if (k === '自動返信メール送信' && String(v).trim() === '無効') emailEnabled = false;
        if (k === '公演名' && v) eventTitle = String(v);
        if (k === '会場名' && v) venue = String(v);
        if (k === '注意事項・備考' && v) notes = String(v);
      }
    }

    // 自動返信メールの送信
    if (emailEnabled && email && email.includes('@')) {
      const emailBody = `${name} 様\n\n`
        + `この度は『${eventTitle}』のチケットをご予約いただき、誠にありがとうございます。\n`
        + `以下の内容でご予約を承りました。\n\n`
        + `----------------------------------------\n`
        + `【予約番号】: ${reservationId}\n`
        + `【お名前】: ${name} 様\n`
        + `【公演日時】: ${scheduleDate}\n`
        + `【券種】: ${ticketType}\n`
        + `【枚数】: ${count}枚\n`
        + (totalPrice ? `【合計金額】: ${totalPrice}\n` : '')
        + (venue ? `【会場】: ${venue}\n` : '')
        + (source ? `【知ったきっかけ】: ${source}\n` : '')
        + `----------------------------------------\n\n`
        + `【ご来場の案内・注意事項】\n`
        + (notes ? `${notes}\n\n` : `・開場は開演の30分前となります。\n・当日受付にて上記「予約番号」または「お名前」をお伝えください。\n\n`)
        + `当日、劇場にてお会いできることを団員一同楽しみにしております！\n\n`
        + `━━━━━━━━━━━━━━━━━━━━━━━━\n`
        + `劇団ハレトケ 公式サイト\n`
        + `https://theatrical.net-menber.com/\n`
        + `お問い合わせフォームよりお気軽にご連絡ください。\n`
        + `━━━━━━━━━━━━━━━━━━━━━━━━`;

      try {
        GmailApp.sendEmail(email, emailSubject, emailBody, {
          name: '劇団ハレトケ チケット窓口'
        });
      } catch (mailErr) {
        Logger.log('メール送信スキップまたはエラー: ' + mailErr);
      }
    }

    const responsePayload = {
      status: 'success',
      reservationId: reservationId,
      message: '予約が正常に完了しました'
    };

    return ContentService.createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    const errorPayload = {
      status: 'error',
      message: error.toString()
    };
    return ContentService.createTextOutput(JSON.stringify(errorPayload))
      .setMimeType(ContentService.MimeType.JSON);
  }
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
    const d = new Date(str.replace(/\//g, '-'));
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

// 日時フォーマット用ヘルパー
function formatDate(d) {
  if (!d || isNaN(d.getTime())) return '';
  const year = d.getFullYear();
  const month = d.getMonth() + 1;
  const date = d.getDate();
  const hours = ('0' + d.getHours()).slice(-2);
  const minutes = ('0' + d.getMinutes()).slice(-2);
  const dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
  return `${year}年${month}月${date}日(${dayOfWeek}) ${hours}:${minutes}`;
}
