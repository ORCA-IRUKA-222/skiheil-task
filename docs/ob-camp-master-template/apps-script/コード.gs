const SHEETS = Object.freeze({
  SETTINGS: '年度設定',
  EXCEPTIONS: '料金例外',
  LIFT_PLAN: 'リフト設定',
  REGISTRATIONS: '参加者',
  DAILY: '日別集計',
});

const SETTINGS_HEADER_ROW = 4;
const LIFT_PLAN_HEADER_ROW = 4;
const REGISTRATION_HEADER_ROW = 4;
const DAILY_HEADER_ROW = 4;

const LIFT_RESORTS = Object.freeze(['未定', '片品', '尾瀬岩鞍', '丸沼', '購入なし']);
const LIFT_PURCHASABLE_RESORTS = Object.freeze(['片品', '尾瀬岩鞍', '丸沼']);

const REGISTRATION_HEADERS = [
  '更新日時', '登録ID', '氏名', '宿泊日', '朝食日（自動）', '夕食日（自動）',
  'リフト購入依頼日', 'バス(行)', 'バス(帰)', '検定',
  '宿泊費', 'リフト費', '交通費', '検定費', '共通費', '合計金額',
  '状態', '入金状況', '備考', 'パスワードハッシュ', 'ソルト',
];

const DAILY_HEADERS = [
  '日付', '宿泊者数', '宿泊者', 'リフト購入依頼数', '指定スキー場',
  'リフト購入依頼者', '朝食数', '朝食', '夕食数', '夕食',
];

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('SKIHEIL 管理')
    .addItem('リフト日程を同期', 'syncLiftPlan')
    .addItem('設定をチェック', 'validateSettings')
    .addItem('日別集計を更新', 'refreshDailySummary')
    .addItem('登録済み全員の金額を再計算', 'recalculateAllCosts')
    .addItem('選択行の仮パスワードを再発行', 'resetSelectedPassword')
    .addSeparator()
    .addItem('新年度を開始（旧データを保存）', 'startNewYear')
    .addToUi();
  tplOnOpen();
}

function onEdit(event) {
  if (!event || !event.range) return;
  const range = event.range;
  const sheet = range.getSheet();
  const editedSheet = sheet.getName();
  if (
    editedSheet === SHEETS.SETTINGS ||
    editedSheet === SHEETS.EXCEPTIONS ||
    editedSheet === SHEETS.LIFT_PLAN
  ) {
    if (range.getRow() > SETTINGS_HEADER_ROW) autoRecalculateCosts_();
    return;
  }

  if (
    sheet.getName() !== SHEETS.REGISTRATIONS ||
    range.getRow() <= REGISTRATION_HEADER_ROW ||
    range.getColumn() !== 18
  ) return;

  const value = String(range.getValue() || '');
  if (value === '入金済') {
    range.setBackground('#EAF7EF').setFontColor('#17633B');
  } else if (value === '未確認') {
    range.setBackground('#FFF2CC').setFontColor('#8A5300');
  } else {
    range.setBackground('#FFFFFF').setFontColor('#17202A');
  }
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.adminEntry = Boolean(
    e && e.parameter && (e.parameter.admin === '1' || e.parameter.admin === 'true'),
  );
  return template
    .evaluate()
    .setTitle('SKIHEIL OB合宿 参加登録')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getAppData() {
  const config = getConfig_();
  return {
    config: publicConfig_(config),
    schedule: buildSchedule_(config),
  };
}

function loadRegistration(name, password) {
  const config = getConfig_();
  const normalizedName = normalizeName_(name);
  const cleanPassword = normalizePassword_(password, config.passwordMinLength);
  const sheet = getSheet_(SHEETS.REGISTRATIONS);
  const rowNumber = findRegistrationRow_(sheet, normalizedName);
  if (!rowNumber) return null;

  const row = sheet.getRange(rowNumber, 1, 1, REGISTRATION_HEADERS.length).getValues()[0];
  verifyRegistrationPassword_(normalizedName, cleanPassword, row[19], row[20]);

  return registrationFromRow_(row);
}

function registrationFromRow_(row) {
  return {
    registrationId: String(row[1] || ''),
    name: String(row[2] || ''),
    stayDates: splitDates_(row[3]),
    breakfast: splitDates_(row[4]),
    dinner: splitDates_(row[5]),
    liftDates: splitDates_(row[6]),
    busOut: row[7] === '利用する',
    busIn: row[8] === '利用する',
    certification: row[9] === '受験する',
    status: String(row[16] || ''),
  };
}

function listRegistrationsForAdmin(adminPassword) {
  const config = getConfig_();
  assertAdmin_(adminPassword, config);
  const sheet = getSheet_(SHEETS.REGISTRATIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= REGISTRATION_HEADER_ROW) return [];
  return sheet
    .getRange(
      REGISTRATION_HEADER_ROW + 1,
      1,
      lastRow - REGISTRATION_HEADER_ROW,
      REGISTRATION_HEADERS.length,
    )
    .getValues()
    .filter((row) => String(row[2] || '').trim())
    .map((row) => ({
      name: String(row[2]).trim(),
      status: String(row[16] || ''),
      payment: String(row[17] || ''),
    }))
    .sort((left, right) => left.name.localeCompare(right.name, 'ja'));
}

function loadRegistrationAsAdmin(adminPassword, name) {
  const config = getConfig_();
  assertAdmin_(adminPassword, config);
  const normalizedName = normalizeName_(name);
  const sheet = getSheet_(SHEETS.REGISTRATIONS);
  const rowNumber = findRegistrationRow_(sheet, normalizedName);
  if (!rowNumber) throw new Error(`登録が見つかりません: ${normalizedName}`);
  const row = sheet.getRange(rowNumber, 1, 1, REGISTRATION_HEADERS.length).getValues()[0];
  return registrationFromRow_(row);
}

function saveRegistration(payload) {
  const config = getConfig_();
  const adminUnlocked = isAdminRequest_(payload, config);
  if (isEditingClosed_(config) && !adminUnlocked) {
    throw new Error(
      '変更受付期限（' + formatDeadline_(config) + '）を過ぎているため、'
      + 'Webからの登録・変更はできません。幹事へご連絡ください。',
    );
  }
  const clean = validatePayload_(payload, config, adminUnlocked);
  const calculation = calculateCost_(clean, config);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const sheet = getSheet_(SHEETS.REGISTRATIONS);
    ensureHeaders_(sheet, REGISTRATION_HEADER_ROW, REGISTRATION_HEADERS);

    const existing = findRegistrationRow_(sheet, clean.name);
    let registrationId = Utilities.getUuid();
    let paymentStatus = '未確認';
    let note = '';
    let salt = createSalt_();
    let passwordHash = hashPassword_(clean.name, clean.password, salt);

    if (existing) {
      const current = sheet
        .getRange(existing, 1, 1, REGISTRATION_HEADERS.length)
        .getValues()[0];
      registrationId = String(current[1] || registrationId);
      paymentStatus = String(current[17] || '未確認');
      note = String(current[18] || '');
      if (current[19] && current[20]) {
        if (!adminUnlocked) {
          verifyRegistrationPassword_(clean.name, clean.password, current[19], current[20]);
        }
        passwordHash = String(current[19]);
        salt = String(current[20]);
      }
    } else if (adminUnlocked) {
      throw new Error(
        '管理者モードでは新規登録はできません。登録済みの人を一覧から選んでください。',
      );
    }

    const row = [
      new Date(),
      registrationId,
      clean.name,
      clean.stayDates.join(', '),
      clean.breakfast.join(', '),
      clean.dinner.join(', '),
      clean.liftDates.join(', '),
      clean.busOut ? '利用する' : '利用しない',
      clean.busIn ? '利用する' : '利用しない',
      clean.certification ? '受験する' : '受験しない',
      calculation.stay,
      calculation.lift,
      calculation.transport,
      calculation.certification,
      calculation.common,
      calculation.total,
      calculation.status,
      paymentStatus,
      note,
      passwordHash,
      salt,
    ];

    const targetRow = existing || sheet.getLastRow() + 1;
    sheet.getRange(targetRow, 1, 1, row.length).setValues([row]);
    formatRegistrationRow_(sheet, targetRow, calculation.status, paymentStatus);

    refreshDailySummary_();
    SpreadsheetApp.flush();

    return {
      registrationId,
      status: calculation.status,
      message: calculation.status === 'キャンセル'
        ? '参加予定をキャンセルとして更新しました。'
        : existing
          ? '登録内容を更新しました。'
          : '参加予定を登録しました。',
      schedule: buildSchedule_(config),
    };
  } finally {
    lock.releaseLock();
  }
}

function syncLiftPlan() {
  SpreadsheetApp.getUi();
  const config = getConfig_();
  syncLiftPlanWithConfig_(config);
  SpreadsheetApp.getActive().toast(
    '合宿日程に合わせてリフト設定を同期しました。',
    'SKIHEIL 管理',
    5,
  );
}

function validateSettings() {
  const config = getConfig_();
  const warnings = [];

  Object.keys(config.exceptions).forEach((date) => {
    if (!config.validDates.has(date)) {
      warnings.push(`料金例外 ${date} は合宿日程外です`);
    }
  });
  Object.keys(config.liftPlan).forEach((date) => {
    if (!config.validDates.has(date)) {
      warnings.push(`リフト設定 ${date} は合宿日程外です`);
    }
  });
  [...config.validDates].forEach((date) => {
    const resort = config.liftPlan[date] ? config.liftPlan[date].resort : '未定';
    if (resort === '未定') warnings.push(`リフト設定 ${date} のスキー場が未定です`);
  });

  const message = warnings.length
    ? `設定は読み込めました。\n確認事項:\n- ${warnings.join('\n- ')}`
    : '設定に問題はありません。';
  SpreadsheetApp.getUi().alert('設定チェック', message, SpreadsheetApp.getUi().ButtonSet.OK);
  return { ok: warnings.length === 0, warnings };
}

function refreshDailySummary() {
  SpreadsheetApp.getUi();
  refreshDailySummary_();
  SpreadsheetApp.getActive().toast('日別集計を更新しました。', 'SKIHEIL 管理', 5);
}

function resetSelectedPassword() {
  const ui = SpreadsheetApp.getUi();
  if (!isSpreadsheetOwner_()) {
    ui.alert('仮パスワードの再発行は、スプレッドシートの所有者だけが実行できます。');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveRange() ? sheet.getActiveRange().getRow() : 0;
  if (sheet.getName() !== SHEETS.REGISTRATIONS || row <= REGISTRATION_HEADER_ROW) {
    ui.alert('参加者シートの対象行を選択してください。');
    return;
  }

  const name = String(sheet.getRange(row, 3).getValue() || '').trim();
  if (!name) {
    ui.alert('選択行に氏名がありません。');
    return;
  }
  const response = ui.alert(
    '仮パスワードを再発行',
    `${name} さんの仮パスワードを再発行しますか？\n現在のパスワードは直ちに使えなくなります。`,
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const currentName = normalizeName_(sheet.getRange(row, 3).getValue());
    if (currentName !== normalizeName_(name)) {
      throw new Error('選択行の氏名が変更されました。対象行を確認してやり直してください。');
    }

    const temporaryPassword = createTemporaryPassword_();
    const salt = createSalt_();
    const passwordHash = hashPassword_(currentName, temporaryPassword, salt);
    sheet.getRange(row, 20, 1, 2).setValues([[passwordHash, salt]]);
    sheet.getRange(row, 1).setValue(new Date()).setNumberFormat('yyyy-mm-dd hh:mm');
    sheet.hideColumns(20, 2);
    SpreadsheetApp.flush();

    ui.alert(
      '仮パスワードを再発行しました',
      `${currentName} さんの仮パスワード:\n\n${temporaryPassword}\n\n` +
        'この画面を閉じると再表示できません。本人確認後、個別に伝えてください。',
      ui.ButtonSet.OK,
    );
  } finally {
    lock.releaseLock();
  }
}

function startNewYear() {
  const ui = SpreadsheetApp.getUi();
  const config = getConfig_();
  const response = ui.alert(
    '新年度を開始',
    '現在データを日時付きシートへアーカイブし、参加者と日別集計を空にします。\nこの操作は元に戻せません。続行しますか？',
    ui.ButtonSet.YES_NO,
  );
  if (response !== ui.Button.YES) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    const suffix = Utilities.formatDate(new Date(), config.timeZone, 'yyyyMMdd_HHmm');
    archiveSheet_(ss, SHEETS.REGISTRATIONS, `参加者_archive_${suffix}`);
    archiveSheet_(ss, SHEETS.DAILY, `日別集計_archive_${suffix}`);

    const registrationSheet = getSheet_(SHEETS.REGISTRATIONS);
    if (registrationSheet.getLastRow() > REGISTRATION_HEADER_ROW) {
      registrationSheet
        .getRange(
          REGISTRATION_HEADER_ROW + 1,
          1,
          registrationSheet.getLastRow() - REGISTRATION_HEADER_ROW,
          REGISTRATION_HEADERS.length,
        )
        .clearContent();
    }
    syncLiftPlanWithConfig_(config);
    refreshDailySummary_();
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  ui.alert(
    '完了',
    '旧データをアーカイブし、新年度用に初期化しました。同じWebアプリURLをそのまま使えます。',
    ui.ButtonSet.OK,
  );
}

function getConfig_() {
  const ss = SpreadsheetApp.getActive();
  const timeZone = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Asia/Tokyo';
  const settingsSheet = getSheet_(SHEETS.SETTINGS);
  const lastRow = settingsSheet.getLastRow();
  if (lastRow <= SETTINGS_HEADER_ROW) throw new Error('年度設定シートに設定がありません。');

  const values = settingsSheet
    .getRange(SETTINGS_HEADER_ROW + 1, 1, lastRow - SETTINGS_HEADER_ROW, 3)
    .getValues();
  const map = {};
  values.forEach((row) => {
    const key = String(row[0] || '').trim();
    if (key) map[key] = row[2];
  });

  const required = [
    'CAMP_NAME', 'EVENT_YEAR', 'START_DATE', 'END_DATE',
    'STAY_NORMAL', 'KATASHINA_NORMAL', 'OZE_NORMAL', 'MARUNUMA_NORMAL',
    'BUS_OUT', 'BUS_IN', 'CERTIFICATION', 'COMMON_FEE', 'PASSWORD_MIN_LENGTH',
  ];
  const missing = required.filter(
    (key) => map[key] === '' || map[key] === null || map[key] === undefined,
  );
  if (missing.length) throw new Error(`年度設定の値が不足しています: ${missing.join(', ')}`);

  const start = asDate_(map.START_DATE, '開始日');
  const end = asDate_(map.END_DATE, '終了日');
  if (start > end) throw new Error('開始日は終了日以前にしてください。');
  const days = Math.round((stripTime_(end) - stripTime_(start)) / 86400000) + 1;
  if (days < 1 || days > 45) throw new Error('合宿日程は1日以上45日以内にしてください。');

  const validDates = new Set();
  for (let date = stripTime_(start); date <= end; date = addDays_(date, 1)) {
    validDates.add(formatIso_(date, timeZone));
  }

  const passwordMinLength = Number(map.PASSWORD_MIN_LENGTH);
  if (!Number.isInteger(passwordMinLength) || passwordMinLength < 4 || passwordMinLength > 32) {
    throw new Error('パスワード最小文字数は4〜32の整数にしてください。');
  }

  return {
    campName: String(map.CAMP_NAME),
    year: Number(map.EVENT_YEAR),
    start,
    end,
    timeZone,
    passwordMinLength,
    prices: {
      stayNormal: asMoney_(map.STAY_NORMAL, '宿泊料金'),
      katashinaNormal: asMoney_(map.KATASHINA_NORMAL, '片品料金'),
      ozeNormal: asMoney_(map.OZE_NORMAL, '尾瀬岩鞍料金'),
      marunumaNormal: asMoney_(map.MARUNUMA_NORMAL, '丸沼料金'),
      busOut: asMoney_(map.BUS_OUT, 'バス往路料金'),
      busIn: asMoney_(map.BUS_IN, 'バス復路料金'),
      certification: asMoney_(map.CERTIFICATION, '検定料金'),
      common: asMoney_(map.COMMON_FEE, '共通費'),
    },
    editDeadline: parseDeadline_(map.EDIT_DEADLINE),
    adminPassword: String(
      map.ADMIN_PASSWORD === null || map.ADMIN_PASSWORD === undefined
        ? ''
        : map.ADMIN_PASSWORD,
    ).trim(),
    exceptions: readExceptions_(timeZone),
    liftPlan: readLiftPlan_(timeZone),
    validDates,
  };
}

function publicConfig_(config) {
  const days = [];
  for (let date = stripTime_(config.start); date <= config.end; date = addDays_(date, 1)) {
    const iso = formatIso_(date, config.timeZone);
    const plan = config.liftPlan[iso] || { resort: '未定', note: '' };
    days.push({
      iso,
      label: Utilities.formatDate(date, config.timeZone, 'M/d'),
      weekday: ['日', '月', '火', '水', '木', '金', '土'][date.getDay()],
      canStay: date < config.end,
      liftResort: plan.resort,
      liftEnabled: LIFT_PURCHASABLE_RESORTS.indexOf(plan.resort) >= 0,
      liftNote: String(plan.note || ''),
    });
  }
  return {
    campName: config.campName,
    year: config.year,
    startIso: formatIso_(config.start, config.timeZone),
    endIso: formatIso_(config.end, config.timeZone),
    passwordMinLength: config.passwordMinLength,
    editDeadlineLabel: formatDeadline_(config),
    editingClosed: isEditingClosed_(config),
    adminUnlockAvailable: Boolean(config.adminPassword),
    days,
  };
}

function readExceptions_(timeZone) {
  const sheet = getSheet_(SHEETS.EXCEPTIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= SETTINGS_HEADER_ROW) return {};

  const rows = sheet
    .getRange(SETTINGS_HEADER_ROW + 1, 1, lastRow - SETTINGS_HEADER_ROW, 5)
    .getValues();
  const result = {};
  rows.forEach((row) => {
    if (!row[0]) return;
    const date = asDate_(row[0], '料金例外の日付');
    const iso = formatIso_(date, timeZone);
    result[iso] = {
      stay: nullableMoney_(row[1], '料金例外の宿泊費'),
      katashina: nullableMoney_(row[2], '料金例外の片品料金'),
      oze: nullableMoney_(row[3], '料金例外の尾瀬岩鞍料金'),
      marunuma: nullableMoney_(row[4], '料金例外の丸沼料金'),
      note: String(row[5] || ''),
    };
  });
  return result;
}

function readLiftPlan_(timeZone) {
  const sheet = getSheet_(SHEETS.LIFT_PLAN);
  const lastRow = sheet.getLastRow();
  if (lastRow <= LIFT_PLAN_HEADER_ROW) return {};

  const rows = sheet
    .getRange(LIFT_PLAN_HEADER_ROW + 1, 1, lastRow - LIFT_PLAN_HEADER_ROW, 3)
    .getValues();
  const result = {};
  rows.forEach((row) => {
    if (!row[0]) return;
    const date = asDate_(row[0], 'リフト設定の日付');
    const iso = formatIso_(date, timeZone);
    const resort = String(row[1] || '未定').trim();
    if (!LIFT_RESORTS.includes(resort)) {
      throw new Error(`リフト設定 ${iso} のスキー場が不正です: ${resort}`);
    }
    result[iso] = {
      resort,
      note: String(row[2] || ''),
    };
  });
  return result;
}

function validatePayload_(payload, config, adminUnlocked) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('送信内容を読み取れませんでした。');
  }
  const name = normalizeName_(payload.name);
  const password = adminUnlocked
    ? ''
    : normalizePassword_(payload.password, config.passwordMinLength);
  const stayDates = validDateList_(payload.stayDates, config.validDates, config, true);
  const breakfast = validDateList_(payload.breakfast, config.validDates, config, false);
  const dinner = validDateList_(payload.dinner, config.validDates, config, false);
  const liftDates = validDateList_(payload.liftDates, config.validDates, config, false);

  liftDates.forEach((date) => {
    const resort = config.liftPlan[date] ? config.liftPlan[date].resort : '未定';
    if (LIFT_PURCHASABLE_RESORTS.indexOf(resort) < 0) {
      throw new Error(`${date} はリフト購入を依頼できません。管理者の指定を確認してください。`);
    }
  });

  return {
    name,
    password,
    stayDates,
    breakfast,
    dinner,
    liftDates,
    busOut: payload.busOut === true,
    busIn: payload.busIn === true,
    certification: payload.certification === true,
  };
}

function deriveMeals_(stayDates, config) {
  const dinner = [...stayDates];
  const breakfast = stayDates
    .map((iso) => formatIso_(addDays_(parseIso_(iso), 1), config.timeZone))
    .filter((iso) => config.validDates.has(iso));
  return {
    breakfast: [...new Set(breakfast)].sort(),
    dinner: [...new Set(dinner)].sort(),
  };
}

function calculateCost_(payload, config) {
  const priceFor = (iso, key, fallback) => {
    const override = config.exceptions[iso] || {};
    return numberOrDefault_(override[key], fallback);
  };
  const stay = payload.stayDates.reduce(
    (sum, date) => sum + priceFor(date, 'stay', config.prices.stayNormal),
    0,
  );
  const lift = payload.liftDates.reduce((sum, date) => {
    const plan = config.liftPlan[date] || { resort: '未定' };
    const resort = plan.resort;
    if (resort === '片品') {
      return sum + priceFor(date, 'katashina', config.prices.katashinaNormal);
    }
    if (resort === '丸沼') {
      return sum + priceFor(date, 'marunuma', config.prices.marunumaNormal);
    }
    if (resort === '尾瀬岩鞍') {
      return sum + priceFor(date, 'oze', config.prices.ozeNormal);
    }
    return sum;
  }, 0);
  const transport =
    (payload.busOut ? config.prices.busOut : 0) +
    (payload.busIn ? config.prices.busIn : 0);
  const certification = payload.certification ? config.prices.certification : 0;
  const subtotal = stay + lift + transport + certification;
  const common = subtotal > 0 ? config.prices.common : 0;
  const total = subtotal + common;

  return {
    stay,
    lift,
    transport,
    certification,
    common,
    total,
    status: total > 0 ? '受付済' : 'キャンセル',
  };
}

function buildSchedule_(config) {
  const schedule = {};
  for (let date = stripTime_(config.start); date <= config.end; date = addDays_(date, 1)) {
    const iso = formatIso_(date, config.timeZone);
    const plan = config.liftPlan[iso] || { resort: '未定' };
    schedule[iso] = {
      iso,
      label: Utilities.formatDate(date, config.timeZone, 'M/d'),
      resort: plan.resort,
      stay: [],
      lift: [],
      breakfast: [],
      dinner: [],
    };
  }

  const sheet = getSheet_(SHEETS.REGISTRATIONS);
  const lastRow = sheet.getLastRow();
  if (lastRow <= REGISTRATION_HEADER_ROW) return Object.values(schedule);
  const rows = sheet
    .getRange(
      REGISTRATION_HEADER_ROW + 1,
      1,
      lastRow - REGISTRATION_HEADER_ROW,
      REGISTRATION_HEADERS.length,
    )
    .getValues();

  rows.forEach((row) => {
    const name = String(row[2] || '').trim();
    if (!name || row[16] !== '受付済') return;
    addNameToSchedule_(schedule, splitDates_(row[3]), 'stay', name);
    addNameToSchedule_(schedule, splitDates_(row[6]), 'lift', name);
    addNameToSchedule_(schedule, splitDates_(row[4]), 'breakfast', name);
    addNameToSchedule_(schedule, splitDates_(row[5]), 'dinner', name);
  });
  return Object.values(schedule);
}

function refreshDailySummary_() {
  const config = getConfig_();
  const schedule = buildSchedule_(config);
  const sheet = getSheet_(SHEETS.DAILY);
  ensureHeaders_(sheet, DAILY_HEADER_ROW, DAILY_HEADERS);

  if (sheet.getLastRow() > DAILY_HEADER_ROW) {
    sheet
      .getRange(
        DAILY_HEADER_ROW + 1,
        1,
        sheet.getLastRow() - DAILY_HEADER_ROW,
        DAILY_HEADERS.length,
      )
      .clearContent();
  }
  if (!schedule.length) return;

  const rows = schedule.map((day) => [
    parseIso_(day.iso),
    day.stay.length,
    day.stay.join('\n'),
    day.lift.length,
    day.resort,
    day.lift.join('\n'),
    day.breakfast.length,
    day.breakfast.join('\n'),
    day.dinner.length,
    day.dinner.join('\n'),
  ]);
  sheet
    .getRange(DAILY_HEADER_ROW + 1, 1, rows.length, DAILY_HEADERS.length)
    .setValues(rows);
  sheet.getRange(DAILY_HEADER_ROW + 1, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(DAILY_HEADER_ROW + 1, 3, rows.length, 8).setWrap(true);
  sheet.autoResizeRows(DAILY_HEADER_ROW + 1, rows.length);
}

function syncLiftPlanWithConfig_(config) {
  const sheet = getSheet_(SHEETS.LIFT_PLAN);
  const existing = readLiftPlan_(config.timeZone);
  const rows = [];
  for (let date = stripTime_(config.start); date <= config.end; date = addDays_(date, 1)) {
    const iso = formatIso_(date, config.timeZone);
    const current = existing[iso] || { resort: '未定', note: '' };
    rows.push([parseIso_(iso), current.resort, current.note]);
  }

  if (sheet.getLastRow() > LIFT_PLAN_HEADER_ROW) {
    sheet
      .getRange(
        LIFT_PLAN_HEADER_ROW + 1,
        1,
        sheet.getLastRow() - LIFT_PLAN_HEADER_ROW,
        3,
      )
      .clearContent()
      .clearDataValidations();
  }
  if (!rows.length) return;
  sheet
    .getRange(LIFT_PLAN_HEADER_ROW + 1, 1, rows.length, 3)
    .setValues(rows)
    .setVerticalAlignment('middle');
  sheet
    .getRange(LIFT_PLAN_HEADER_ROW + 1, 1, rows.length, 1)
    .setNumberFormat('yyyy-mm-dd');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(LIFT_RESORTS, true)
    .setAllowInvalid(false)
    .build();
  sheet
    .getRange(LIFT_PLAN_HEADER_ROW + 1, 2, rows.length, 1)
    .setDataValidation(rule)
    .setBackground('#FFF2CC');
}

function findRegistrationRow_(sheet, name) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= REGISTRATION_HEADER_ROW) return null;
  const names = sheet
    .getRange(REGISTRATION_HEADER_ROW + 1, 3, lastRow - REGISTRATION_HEADER_ROW, 1)
    .getValues()
    .flat();
  const index = names.findIndex((value) => normalizeName_(value, false) === name);
  return index === -1 ? null : REGISTRATION_HEADER_ROW + 1 + index;
}

function archiveSheet_(ss, sourceName, archiveName) {
  if (ss.getSheetByName(archiveName)) {
    throw new Error(`アーカイブ名が重複しました: ${archiveName}`);
  }
  const source = getSheet_(sourceName);
  const archive = source.copyTo(ss).setName(archiveName);
  const range = archive.getDataRange();
  range.copyTo(range, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  archive.protect().setDescription('新年度開始時の自動アーカイブ');
}

function ensureHeaders_(sheet, headerRow, headers) {
  const actual = sheet.getRange(headerRow, 1, 1, headers.length).getValues()[0];
  const matches = headers.every((header, index) => actual[index] === header);
  if (!matches) throw new Error(`${sheet.getName()} シートの見出しが変更されています。`);
}

function formatRegistrationRow_(sheet, row, status, paymentStatus) {
  sheet.getRange(row, 1).setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange(row, 11, 1, 6).setNumberFormat('¥#,##0');
  sheet
    .getRange(row, 1, 1, REGISTRATION_HEADERS.length)
    .setVerticalAlignment('middle')
    .setWrap(true)
    .setBorder(
      null,
      null,
      true,
      null,
      null,
      null,
      '#E7ECF1',
      SpreadsheetApp.BorderStyle.SOLID,
    );

  const statusCell = sheet.getRange(row, 17);
  if (status === '受付済') {
    statusCell.setBackground('#EAF7EF').setFontColor('#17633B');
  } else {
    statusCell.setBackground('#FDECEC').setFontColor('#B42318');
  }

  const paymentCell = sheet.getRange(row, 18);
  if (paymentStatus === '入金済') {
    paymentCell.setBackground('#EAF7EF').setFontColor('#17633B');
  } else {
    paymentCell.setBackground('#FFF2CC').setFontColor('#8A5300');
  }
  sheet.getRange(row, 19).setBackground('#FFF2CC');
  sheet.hideColumns(20, 2);
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error(`必要なシートが見つかりません: ${name}`);
  return sheet;
}

function normalizeName_(value, strict = true) {
  const normalized = String(value || '').trim().replace(/\s+/g, ' ');
  if (!strict) return normalized;
  if (!normalized) throw new Error('氏名を入力してください。');
  if (normalized.length > 60) throw new Error('氏名は60文字以内で入力してください。');
  return normalized;
}

function normalizePassword_(value, minLength) {
  const password = String(value === null || value === undefined ? '' : value);
  const length = Array.from(password).length;
  if (length < minLength) {
    throw new Error(`編集用パスワードは${minLength}文字以上で入力してください。`);
  }
  if (length > 128) throw new Error('編集用パスワードは128文字以内で入力してください。');
  return password;
}

function createSalt_() {
  return `${Utilities.getUuid()}${Utilities.getUuid()}`;
}

function createTemporaryPassword_() {
  const alphabet = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    `${Utilities.getUuid()}\n${Utilities.getUuid()}\n${new Date().getTime()}`,
    Utilities.Charset.UTF_8,
  );
  const characters = bytes.slice(0, 12).map((byte) => {
    const unsignedByte = (Number(byte) + 256) % 256;
    return alphabet[unsignedByte % alphabet.length];
  });
  return [
    characters.slice(0, 4).join(''),
    characters.slice(4, 8).join(''),
    characters.slice(8, 12).join(''),
  ].join('-');
}

function isSpreadsheetOwner_() {
  const spreadsheet = SpreadsheetApp.getActive();
  const owner = spreadsheet.getOwner();
  const ownerEmail = owner ? String(owner.getEmail() || '').toLowerCase() : '';
  const currentEmail = String(Session.getEffectiveUser().getEmail() || '').toLowerCase();
  return Boolean(ownerEmail && currentEmail && ownerEmail === currentEmail);
}

function hashPassword_(name, password, salt) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    `${salt}\n${name}\n${password}`,
    Utilities.Charset.UTF_8,
  );
  return Utilities.base64Encode(bytes);
}

function verifyRegistrationPassword_(name, password, storedHash, salt) {
  if (!storedHash || !salt) return;
  const actual = hashPassword_(name, password, String(salt));
  if (!safeEqual_(String(storedHash), actual)) {
    throw new Error('氏名またはパスワードが正しくありません。');
  }
}

function safeEqual_(left, right) {
  if (left.length !== right.length) return false;
  let difference = 0;
  for (let index = 0; index < left.length; index += 1) {
    difference |= left.charCodeAt(index) ^ right.charCodeAt(index);
  }
  return difference === 0;
}

function validDateList_(values, validSet, config, overnight) {
  const list = Array.isArray(values) ? values : [];
  const unique = [
    ...new Set(list.map((value) => String(value || '').trim()).filter(Boolean)),
  ];
  unique.forEach((date) => {
    if (!validSet.has(date)) throw new Error(`日程外の日付が含まれています: ${date}`);
    if (overnight && date === formatIso_(config.end, config.timeZone)) {
      throw new Error('最終日は宿泊日に選べません。');
    }
  });
  return unique.sort();
}

function splitDates_(value) {
  if (!value) return [];
  if (value instanceof Date) {
    return [Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd')];
  }
  return String(value)
    .split(',')
    .map((item) => item.trim())
    .filter(Boolean);
}

function addNameToSchedule_(schedule, dates, key, name) {
  dates.forEach((date) => {
    if (schedule[date]) schedule[date][key].push(name);
  });
}

function asDate_(value, label) {
  const date = value instanceof Date ? new Date(value) : new Date(value);
  if (Number.isNaN(date.getTime())) throw new Error(`${label}が正しい日付ではありません。`);
  return stripTime_(date);
}

function parseIso_(iso) {
  const parts = String(iso).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function stripTime_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function addDays_(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function formatIso_(date, timeZone) {
  return Utilities.formatDate(date, timeZone, 'yyyy-MM-dd');
}

function asMoney_(value, label) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) {
    throw new Error(`${label}は0以上の数値にしてください。`);
  }
  return Math.round(number);
}

function nullableMoney_(value, label) {
  if (value === '' || value === null || value === undefined) return null;
  return asMoney_(value, label);
}

function numberOrDefault_(value, fallback) {
  return value === null || value === undefined || value === '' ? fallback : Number(value);
}

function parseDeadline_(value) {
  if (value === '' || value === null || value === undefined) return null;
  let date;
  if (value instanceof Date) {
    date = new Date(value.getTime());
  } else {
    const text = String(value).trim();
    if (!text) return null;
    const parsed = new Date(text.replace(/\//g, '-').replace(' ', 'T'));
    if (isNaN(parsed.getTime())) {
      throw new Error(`変更受付期限の書式が正しくありません: ${text}`);
    }
    date = parsed;
  }
  if (date.getHours() === 0 && date.getMinutes() === 0 && date.getSeconds() === 0) {
    date = new Date(date.getTime() + 86400000 - 1000);
  }
  return date;
}

function formatDeadline_(config) {
  if (!config.editDeadline) return '';
  return Utilities.formatDate(config.editDeadline, config.timeZone, 'yyyy年M月d日 HH:mm');
}

function isEditingClosed_(config) {
  return Boolean(config.editDeadline && new Date().getTime() > config.editDeadline.getTime());
}

function isAdminRequest_(payload, config) {
  if (!config.adminPassword) return false;
  const given = payload && payload.adminPassword ? String(payload.adminPassword) : '';
  if (!given) return false;
  return safeEqual_(given, config.adminPassword);
}

function assertAdmin_(password, config) {
  if (!config.adminPassword) {
    throw new Error('管理者パスワードが未設定です。年度設定シートの ADMIN_PASSWORD に値を入れてください。');
  }
  if (!safeEqual_(String(password || ''), config.adminPassword)) {
    throw new Error('管理者パスワードが正しくありません。');
  }
}

function verifyAdminPassword(password) {
  assertAdmin_(password, getConfig_());
  return true;
}

function recalculateAllCosts() {
  const ui = SpreadsheetApp.getUi();
  let result;
  try {
    result = recalculateAllCosts_(getConfig_());
  } catch (err) {
    ui.alert(
      '再計算できませんでした',
      String(err && err.message ? err.message : err),
      ui.ButtonSet.OK,
    );
    return;
  }
  const lines = [`対象: ${result.checked} 人`, `金額を更新: ${result.updated} 人`];
  if (result.warnings.length) {
    lines.push('');
    lines.push('注意:');
    result.warnings.slice(0, 20).forEach((text) => lines.push(`・${text}`));
    if (result.warnings.length > 20) {
      lines.push(`ほか ${result.warnings.length - 20} 件`);
    }
  }
  ui.alert('金額の再計算が完了しました', lines.join('\n'), ui.ButtonSet.OK);
}

function recalculateAllCosts_(config) {
  const sheet = getSheet_(SHEETS.REGISTRATIONS);
  const lastRow = sheet.getLastRow();
  const warnings = [];
  if (lastRow <= REGISTRATION_HEADER_ROW) return { checked: 0, updated: 0, warnings };

  const values = sheet
    .getRange(
      REGISTRATION_HEADER_ROW + 1,
      1,
      lastRow - REGISTRATION_HEADER_ROW,
      REGISTRATION_HEADERS.length,
    )
    .getValues();

  let checked = 0;
  let updated = 0;

  values.forEach((row, index) => {
    const name = String(row[2] || '').trim();
    if (!name) return;
    checked += 1;

    const payload = {
      stayDates: splitDates_(row[3]),
      breakfast: splitDates_(row[4]),
      dinner: splitDates_(row[5]),
      liftDates: splitDates_(row[6]),
      busOut: row[7] === '利用する',
      busIn: row[8] === '利用する',
      certification: row[9] === '受験する',
    };

    const outside = [
      ...new Set(
        payload.stayDates
          .concat(payload.liftDates)
          .filter((iso) => !config.validDates.has(iso)),
      ),
    ];
    if (outside.length) {
      warnings.push(`${name}: 合宿日程外の日付が残っています（${outside.join(', ')}）`);
    }
    payload.liftDates.forEach((iso) => {
      const plan = config.liftPlan[iso];
      const resort = plan ? plan.resort : '未定';
      if (LIFT_PURCHASABLE_RESORTS.indexOf(resort) < 0) {
        warnings.push(`${name}: ${iso} は「${resort}」のためリフト費0円で計算しました。`);
      }
    });

    const calculation = calculateCost_(payload, config);
    const after = [
      calculation.stay,
      calculation.lift,
      calculation.transport,
      calculation.certification,
      calculation.common,
      calculation.total,
      calculation.status,
    ];
    const before = row.slice(10, 17);
    let unchanged = String(before[6]) === String(calculation.status);
    for (let position = 0; position < 6; position += 1) {
      if (Number(before[position] || 0) !== Number(after[position])) unchanged = false;
    }
    if (unchanged) return;

    const rowNumber = REGISTRATION_HEADER_ROW + 1 + index;
    sheet.getRange(rowNumber, 11, 1, 7).setValues([after]);
    sheet.getRange(rowNumber, 11, 1, 6).setNumberFormat('¥#,##0');
    const statusCell = sheet.getRange(rowNumber, 17);
    if (calculation.status === '受付済') {
      statusCell.setBackground('#EAF7EF').setFontColor('#17633B');
    } else {
      statusCell.setBackground('#FDECEC').setFontColor('#B42318');
    }
    updated += 1;
  });

  if (updated) SpreadsheetApp.flush();
  return { checked, updated, warnings };
}

function autoRecalculateCosts_() {
  const ss = SpreadsheetApp.getActive();
  try {
    const result = recalculateAllCosts_(getConfig_());
    if (result.updated) {
      ss.toast(`${result.updated} 人の金額を新しい設定で更新しました。`, 'SKIHEIL 管理', 6);
    }
  } catch (err) {
    ss.toast(
      `金額を再計算できませんでした: ${err && err.message ? err.message : err}`,
      'SKIHEIL 管理',
      8,
    );
  }
}
