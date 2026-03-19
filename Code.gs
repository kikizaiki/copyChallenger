/**
 * AI-Редактор для Google Sheets (GAS)
 * - Меню: AI-Редактор -> Обработать текущую строку / Обработать выделенный диапазон
 * - Читает контекст из листа "Main" (A1, C, D, E, G) и пишет комментарий в столбец K
 * - Читает API Key/URL/Model и "Золотые правила" из листа "Settings" (A2/B2/C2 и E2:E)
 */

const AI_EDITOR_MENU_ROOT = "AI-Редактор";
const MENU_CURRENT_ROW = "Обработать текущую строку";
const MENU_SELECTED_RANGE = "Обработать выделенный диапазон";

// "Main" columns (1-indexed)
const COL_CHANNEL = 3; // C
const COL_AD_MARK = 4; // D
const COL_POST_DATE = 5; // E
const COL_POST_TEXT = 7; // G
const COL_OUTPUT = 11; // K

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu(AI_EDITOR_MENU_ROOT)
    .addItem(MENU_CURRENT_ROW, "processActiveRow_")
    .addItem(MENU_SELECTED_RANGE, "processSelectedRange_")
    .addToUi();
}

function processActiveRow_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getActiveRange();
  if (!range) {
    ss.toast("Нет активной ячейки. Выделите строку.", "AI-Редактор", 6);
    return;
  }

  const sheet = ss.getSheetByName("Контент");
  if (!sheet) {
    ss.toast('Не найден лист "Контент".', "AI-Редактор", 10);
    return;
  }

  const row = range.getRow();
  processRows_(sheet, [row]);
}

function processSelectedRange_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const range = ss.getActiveRange();
  if (!range) {
    ss.toast("Нет выделения. Выделите диапазон строк.", "AI-Редактор", 6);
    return;
  }

  const sheet = range.getSheet();
  if (sheet.getName() !== "Контент") {
    ss.toast('Выделение должно быть на листе "Контент".', "AI-Редактор", 8);
    return;
  }

  const startRow = range.getRow();
  const numRows = range.getNumRows();
  const rows = [];
  for (let i = 0; i < numRows; i++) rows.push(startRow + i);
  processRows_(sheet, rows);
}

function processRows_(mainSheet, rows) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettingsSheet_(ss);
  if (!settings) return;

  const api = getApiConfig_(settings);
  if (!api) return;

  const goldenRules = getGoldenRules_(settings);
  const movieInfo = (mainSheet.getRange("A1").getDisplayValue() || "").trim();

  const today = new Date();
  const todayAt00 = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx];

    const postText = (mainSheet.getRange(row, COL_POST_TEXT).getDisplayValue() || "").trim();
    if (!postText) continue; // Требование: если текст пустой — пропускать строку

    const channel = (mainSheet.getRange(row, COL_CHANNEL).getDisplayValue() || "").trim();
    const adMarkRaw = (mainSheet.getRange(row, COL_AD_MARK).getDisplayValue() || "").trim();
    const adMark = normalizeRuBoolean_(adMarkRaw);
    const postDateVal = mainSheet.getRange(row, COL_POST_DATE).getValue();
    const postDate = toDateOrNull_(postDateVal);

    const timing = getTimingLabel_(postDate, todayAt00);

    ss.toast(`AI думает... (${idx + 1}/${rows.length})`, "AI-Редактор", 4);

    const prompt = buildPrompt_({
      movieInfo,
      channel,
      adMark,
      adMarkRaw,
      timing,
      goldenRules,
      postText,
      rowId: row,
    });

    const comment = callAiChatCompletions_(api, prompt);
    if (comment) {
      mainSheet.getRange(row, COL_OUTPUT).setValue(comment);
    }
  }
}

function getSettingsSheet_(ss) {
  const settings = ss.getSheetByName("Settings");
  if (!settings) {
    ss.toast('Не найден лист "Settings".', "AI-Редактор", 10);
    return null;
  }
  return settings;
}

function getApiConfig_(settingsSheet) {
  const apiKey = (settingsSheet.getRange("A2").getDisplayValue() || "").trim();
  const apiUrl = (settingsSheet.getRange("B2").getDisplayValue() || "").trim();
  const model = (settingsSheet.getRange("C2").getDisplayValue() || "").trim();

  if (!apiKey || !apiUrl || !model) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Проверьте "Settings": A2 (API Key), B2 (API URL), C2 (Model).',
      "AI-Редактор",
      12
    );
    return null;
  }

  return { apiKey, apiUrl, model };
}

function getGoldenRules_(settingsSheet) {
  const lastRow = Math.max(settingsSheet.getLastRow(), 2);
  if (lastRow < 2) return [];

  const values = settingsSheet.getRange(2, 5, lastRow - 1, 1).getValues(); // column E
  const rules = [];
  for (let i = 0; i < values.length; i++) {
    const v = (values[i][0] || "").toString().trim();
    if (v) rules.push(v);
  }
  return rules;
}

function normalizeRuBoolean_(value) {
  // Допускаем разные форматы в ячейке D:
  // - "есть"/"да"/"true"/"1" => true
  // - "нет"/"нету"/"false"/"0" => false
  // - пусто => false
  const v = (value || "").toString().trim().toLowerCase();
  if (!v) return false;
  if (["есть", "да", "true", "1", "y", "yes", "есть-точ"].includes(v)) return true;
  if (["нет", "нету", "false", "0", "n", "no"].includes(v)) return false;
  // Если пользователь написал что-то отличное — считаем непустое как "есть".
  return true;
}

function toDateOrNull_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  // Иногда встречаются строки даты
  if (typeof value === "string") {
    const s = value.trim();
    if (!s) return null;
    const d = new Date(s);
    if (!isNaN(d.getTime())) return d;
  }
  return null;
}

function getTimingLabel_(postDate, todayAt00) {
  // Требование: если премьера прошла -> "уже в кино", если нет -> "скоро"
  if (!postDate) return "скоро"; // консервативно
  const dAt00 = new Date(postDate.getFullYear(), postDate.getMonth(), postDate.getDate());
  return dAt00.getTime() <= todayAt00.getTime() ? "уже в кино" : "скоро";
}

function buildPrompt_({ movieInfo, channel, adMark, adMarkRaw, timing, goldenRules, postText, rowId }) {
  const channelLc = (channel || "").toLowerCase();
  const isVkLike = /(^|[\s\W])vk($|[\s\W])/.test(channelLc) || /вк/.test(channelLc) || channelLc.includes("vk");

  const goldenRulesBlock = goldenRules.length
    ? goldenRules.map((r) => `- ${r}`).join("\n")
    : "- (список золотых правил пуст)";

  const systemPrompt =
    `Ты опытный SMM-редактор (ЦПШ/GPM), проверяешь работу копирайтера.\n` +
    `Твоя задача: по контексту канала и поста подготовить уникальный редакторский комментарий.\n\n` +
    `Контекст канала: ${channel || "(не указан)"}.\n` +
    (isVkLike
      ? `В канале ВК запрещены прямые обращения «ты/вы». Используй безличные формулировки или обращения к ситуации, но не к читателю напрямую.\n`
      : "") +
    `Рекламная метка: ${adMark ? "есть (усилить CTA)" : "нет (делать нативный сторителлинг)"} (как в ячейке D: "${adMarkRaw}").\n` +
    `Тайминг: если ${timing === "уже в кино" ? "премьера прошла" : "премьера ещё не прошла"}, используйте формулировки соответствующего статуса: "${timing}".\n\n` +
    `Стиль комментариев:\n` +
    `- Комментарий должен быть уникальным по формулировкам для каждой строки.\n` +
    `- Избегай канцеляризмов и «пресс-релизности».\n` +
    `- Используй кавычки «…» для примеров формулировок.\n` +
    `- Вариативность длины: коротко (1 предложение) или средне (2–3) или развернуто (3–4) — выбирай подходящий размер под контекст.\n` +
    `- Соблюдай правила пунктуации и типичные правки к тире/дефисам; акценты делай на актёров/персонажей, если они есть в тексте.\n\n` +
    `Золотые правила клиента:\n${goldenRulesBlock}\n\n` +
    `Вывод: верни только готовый редакторский комментарий одной фразой/абзацем на русском без заголовков, раздели его на список.\n` +
    `Ограничение: не повторяй предыдущие варианты текста; используй уникальные формулировки (ориентир: строка ${rowId}).`;

  const userPrompt =
    `Исходные данные:\n` +
    `1) Инфо о фильме (из A1): ${movieInfo || "(пусто)"}\n` +
    `2) Канал: ${channel || "(пусто)"}\n` +
    `3) Рекламная метка: ${adMark ? "есть" : "нет"}\n` +
    `4) Статус по таймингу: ${timing}\n` +
    `5) Текст поста (из G): ${postText}\n\n` +
    `Сформируй редакторский комментарий: что бы ты поправил/усилил в тексте поста, учитывая метку (CTA или нативность) и статус (уже/скоро).`;

  return {
    model: null, // задается в callAiChatCompletions_
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
  };
}

function callAiChatCompletions_(api, prompt) {
  const body = {
    model: api.model,
    messages: prompt.messages,
    // Небольшая вариативность, чтобы не копировать один и тот же шаблон
    temperature: 0.9,
  };

  const options = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${api.apiKey}`,
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  };

  try {
    const resp = UrlFetchApp.fetch(api.apiUrl, options);
    const code = resp.getResponseCode();
    const text = resp.getContentText();

    if (code < 200 || code >= 300) {
      throw new Error(`AI API HTTP ${code}: ${text}`);
    }

    const json = JSON.parse(text);
    const content =
      json?.choices?.[0]?.message?.content ||
      json?.choices?.[0]?.text ||
      json?.data?.[0]?.content;

    if (!content) return "";
    return content.toString().trim();
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast(`Ошибка AI: ${e.message}`, "AI-Редактор", 20);
    Logger.log(e);
    return "";
  }
}

