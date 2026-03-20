/**
 * AI-Редактор для Google Sheets (GAS)
 * - Меню: AI-Редактор -> Обработать текущую строку / Обработать выделенный диапазон
 * - Читает контекст из листа "Main" (A1, C, D, E, G) и пишет комментарий в столбец K
 * - Читает API Key/URL/Model и "Золотые правила" из листа "Settings" (A2/B2/C2 и E2:E)
 */

const AI_EDITOR_MENU_ROOT = "AI-Редактор";
const MENU_CURRENT_ROW = "Обработать текущую строку";
const MENU_SELECTED_RANGE = "Обработать выделенный диапазон";
const LOG_SHEET_NAME = "Logs";

/**
 * Надстройка «свой промпт»: если true — перед обработкой показывается диалог «Дополнить промпт».
 * Поставьте false при тестах, чтобы сразу вызывать API без окна (дополнительный текст не добавляется).
 */
const ENABLE_EXTRA_PROMPT_DIALOG = false;

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

  const extra = getExtraPromptFromUser_();
  if (extra === null) return;

  const row = range.getRow();
  processRows_(sheet, [row], extra);
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

  const extra = getExtraPromptFromUser_();
  if (extra === null) return;

  processRows_(sheet, rows, extra);
}

/**
 * Учитывает ENABLE_EXTRA_PROMPT_DIALOG: при false возвращает "" без диалога.
 * @returns {string|null} дополнение к промпту, "" если выключено, null если в диалоге нажали «Отмена»
 */
function getExtraPromptFromUser_() {
  if (!ENABLE_EXTRA_PROMPT_DIALOG) {
    return "";
  }
  return promptForExtraPrompt_();
}

/**
 * Всплывающее окно: опционально дополнить промпт своим текстом.
 * @returns {string|null} текст дополнения или null, если пользователь нажал «Отмена»
 */
function promptForExtraPrompt_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    "Дополнить промпт",
    "Можно добавить свои пожелания к генерации комментария (тон, акценты, что важно учесть). Оставьте поле пустым — будет использован только стандартный промпт.",
    ui.ButtonSet.OK_CANCEL
  );
  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  return (response.getResponseText() || "").trim();
}

function processRows_(mainSheet, rows, extraUserPrompt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureLogSheet_(ss); // Лист логов должен существовать всегда, даже если ошибок не было.
  appendLog_("RUN_START", "", `Запуск обработки: ${rows.length} строк`, "");
  const settings = getSettingsSheet_(ss);
  if (!settings) return;

  const api = getApiConfig_(settings);
  if (!api) return;

  const goldenRules = getGoldenRules_(settings);
  const movieInfo = (mainSheet.getRange("A1").getDisplayValue() || "").trim();
  let writtenCount = 0;
  let skippedCount = 0;
  let failedCount = 0;

  const today = new Date();
  const todayAt00 = new Date(today.getFullYear(), today.getMonth(), today.getDate());

  for (let idx = 0; idx < rows.length; idx++) {
    const row = rows[idx];

    const postText = (mainSheet.getRange(row, COL_POST_TEXT).getDisplayValue() || "").trim();
    if (!postText) {
      skippedCount++; // Требование: если текст пустой — пропускать строку
      continue;
    }

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
      extraUserPrompt: extraUserPrompt || "",
    });

    const comment = callAiChatCompletions_(api, prompt, row);
    if (comment) {
      mainSheet.getRange(row, COL_OUTPUT).setValue(comment);
      writtenCount++;
    } else {
      failedCount++;
    }
  }

  const summary =
    `Готово: ${writtenCount}. Пропущено: ${skippedCount}. Ошибок: ${failedCount}.` +
    (failedCount > 0 ? ` Подробности на листе "${LOG_SHEET_NAME}".` : "");
  appendLog_("RUN_SUMMARY", "", summary, `rows=${rows.length}`);
  ss.toast(summary, "AI-Редактор", 10);
}

function getSettingsSheet_(ss) {
  const settings = ss.getSheetByName("Settings");
  if (!settings) {
    appendLog_("CONFIG_ERROR", "", 'Не найден лист "Settings"', "");
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
    appendLog_(
      "CONFIG_ERROR",
      "",
      'Проверьте "Settings": A2 (API Key), B2 (API URL), C2 (Model).',
      `A2=${apiKey ? "OK" : "EMPTY"}, B2=${apiUrl ? "OK" : "EMPTY"}, C2=${model ? "OK" : "EMPTY"}`
    );
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Проверьте "Settings": A2 (API Key), B2 (API URL), C2 (Model).',
      "AI-Редактор",
      12
    );
    return null;
  }

  const normalizedUrl = normalizeChatCompletionsUrl_(apiUrl);
  return { apiKey, apiUrl: normalizedUrl, model };
}

/**
 * OpenAI-совместимые провайдеры (Polza.ai, OpenRouter, ProxyAPI и т.д.) ожидают POST на .../chat/completions.
 * В Settings часто указывают только базу: https://polza.ai/api/v1 — без суффикса.
 */
function normalizeChatCompletionsUrl_(url) {
  var u = (url || "").trim().replace(/\/+$/, "");
  if (!u) return "";
  var lower = u.toLowerCase();
  if (lower.indexOf("/chat/completions") !== -1) return u;
  if (lower.endsWith("/v1")) return u + "/chat/completions";
  if (lower.endsWith("/api")) return u + "/v1/chat/completions";
  return u + "/chat/completions";
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

function buildPrompt_({
  movieInfo,
  channel,
  adMark,
  adMarkRaw,
  timing,
  goldenRules,
  postText,
  rowId,
  extraUserPrompt,
}) {
  const channelLc = (channel || "").toLowerCase();
  const isVkLike = /(^|[\s\W])vk($|[\s\W])/.test(channelLc) || /вк/.test(channelLc) || channelLc.includes("vk");

  const goldenRulesBlock = goldenRules.length
    ? goldenRules.map((r) => `- ${r}`).join("\n")
    : "- (список золотых правил пуст)";

  const systemPrompt =
    `**Роль (System):**\n` +
    `Ты опытный SMM-редактор (ЦПШ/GPM), проверяешь работу копирайтера.\n` +
    `Пишешь как настоящий редактор-клиент: коротко, по делу, профессионально, но с живостью. Без сленга и излишнего официоза.\n\n` +
    `**Задача:**\n` +
    `Подготовь редакторский комментарий для одной строки таблицы — что именно поправить/усилить в тексте поста.\n\n` +
    `**Входные данные (для одной строки):**\n` +
    `1. **Excel-строка** — исходный текст поста + метаданные строки (канал, рекламная метка, тайминг).\n` +
    `2. **Документ с комментариями клиента (Золотые правила)** — набор правил из Settings.\n` +
    `3. **Бриф по фильму** — информация из Main!A1.\n\n` +
    `Золотые правила клиента:\n${goldenRulesBlock}\n\n` +
    `---\n` +
    `### **Что нужно сделать:**\n` +
    `1. Проанализируй исходный пост и подумай, как клиент обычно формулирует правки: тон, структура, уровень эмоций, тип предложений.\n` +
    `2. Учитывай специфику канала (включая запреты ВК на прямые обращения «ты/вы»).\n` +
    `3. Учитывай рекламную метку: если метка **есть** — усили CTA; если **нет** — делай нативный сторителлинг.\n` +
    `4. Учитывай тайминг: «уже в кино» или «скоро» — формулировки должны поддерживать статус.\n` +
    `5. Сформулируй уникальный редакторский комментарий: конкретно что поправить и где усилить.\n\n` +
    `---\n` +
    `### **Требования к комментариям:**\n` +
    `* **Каждый комментарий должен быть уникальным**: не повторяй формулировки и не копируй структуру “слово в слово” от строки к строке.\n` +
    `* **Стиль — “редактор-клиент”.** Коротко, по делу, профессионально, но с живостью.\n` +
    `* **Тон и формат варьируй.** Чередуй длину: короткие (1 предложение), средние (2–3), развёрнутые (3–4) — под плотность исходного текста.\n` +
    `* **Похвала.** Если в тексте есть отличные заходы – надо похвалить копирайтера.\n` +
    `* **Типы комментариев (чередовать):**\n` +
    `  * замечание по тону (*«звучит прессрелизно, нужно живее»*);\n` +
    `  * пример формулировки (*«можно начать с “…”»*);\n` +
    `  * риторический вопрос (*«А если добавить …?»*);\n` +
    `  * структурная подсказка (*«начать с действия, не с описания»*);\n` +
    `  * личное редакторское наблюдение (*«хочется чуть теплее/ближе к герою»*).\n` +
    `* **По смыслу комментарий должен:** усиливать жанровое восприятие; предлагать конкретные правки без лишней “воды”; содержать живые фразы и примеры в кавычках «…».\n` +
    `* **Анти-повторы (важно):** не используй одни и те же обороты «убрать прессрелизность», «добавить эмоцию», «сделать живее»; если используешь эти смыслы — перефразируй и меняй контекст.\n\n` +
    `* **Антишаблоны (важно):** не используй тривиальный CTA в виде фразы «смотрите в кино» и прямых её вариаций. Для CTA подбирай более живые формулировки, которые поддерживают статус «уже в кино»/«скоро», но звучат как редакторская ремарка, а не как рекламный шаблон.\n` +
    (isVkLike
      ? `* **Специфика ВК:** запрещены прямые обращения к читателю «ты/вы». Используй безличные формулировки или обращение к ситуации/ходу текста, но не к аудитории напрямую.\n`
      : "") +
    `---\n` +
    `### **Формат вывода:**\n` +
    `Верни *только готовый редакторский комментарий* на русском.\n` +
    `Никаких заголовков. Один абзац (3–7 предложения).\n\n` +
    `---\n` +
    `### **Технические требования (внутренние для ChatGPT):**\n` +
    `* Используй ротацию не менее 25 шаблонов комментариев разного типа.\n` +
    `* Не допускай повторов ни по смыслу, ни по структуре.\n` +
    `* Следи, чтобы комментарии выглядели естественно (как от разных редакторов).\n` +
    `* Чередуй ритм предложений: короткие, длинные, комбинированные.\n` +
    `* Учитывай данные из брифа: настроение фильма, жанр, эмоциональный ключ.\n` +
    `* При необходимости добавляй 1 небольшой конкретный пример формулировки в кавычках «…».\n\n` +
    `Ограничение: комментарий должен быть уникальным для строки ${rowId}.`;

  const extraBlock =
    extraUserPrompt && extraUserPrompt.length > 0
      ? `\n**Дополнительные пожелания редактора (учти в приоритете, не противоречь золотым правилам):**\n${extraUserPrompt}\n`
      : "";

  const userPrompt =
    `Исходные данные (для строки ${rowId}):\n` +
    `1) Бриф по фильму (Main!A1): ${movieInfo || "(пусто)"}\n` +
    `2) Канал (Main!C): ${channel || "(пусто)"}\n` +
    `3) Рекламная метка (Main!D): ${adMark ? "есть" : "нет"} (как в ячейке D: "${adMarkRaw}")\n` +
    `4) Тайминг (Main!E): ${timing}\n` +
    `5) Текст поста (Main!G): ${postText}\n` +
    extraBlock +
    `\nСформируй редакторский комментарий: что именно поправить/усилить, учитывая метку (CTA или нативность) и тайминг (уже в кино/скоро).`;

  return {
    model: null, // задается в callAiChatCompletions_
    messages: [
      { role: "system", content: systemPrompt },
      { role: "user", content: userPrompt },
    ],
  };
}

function callAiChatCompletions_(api, prompt, row) {
  const bannedNeedle = "смотрите в кино";

  const baseBody = {
    model: api.model,
    // Небольшая вариативность, чтобы не копировать один и тот же шаблон
    temperature: 1,
  };

  const optionsBase = {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: `Bearer ${api.apiKey}`,
      Accept: "application/json",
    },
    muteHttpExceptions: true,
  };

  const antiTemplateSystem =
    `Антишаблон: категорически избегай тривиального CTA «${bannedNeedle}» и любых прямых вариаций. Придумай более живую редакторскую ремарку, которая поддерживает статус «уже в кино»/«скоро», но звучит свежо.`;

  const messages = prompt.messages.concat([{ role: "system", content: antiTemplateSystem }]);

  const body = Object.assign({}, baseBody, { messages });
  const options = Object.assign({}, optionsBase, {
    payload: JSON.stringify(body),
  });

  let lastContent = "";
  try {
    const resp = UrlFetchApp.fetch(api.apiUrl, options);
    const code = resp.getResponseCode();
    const text = resp.getContentText();

    if (code < 200 || code >= 300) {
      throw new Error(`HTTP ${code}: ${text}`);
    }

    const json = JSON.parse(text);
    const content =
      json?.choices?.[0]?.message?.content ||
      json?.choices?.[0]?.text ||
      json?.data?.[0]?.content;

    const usageInfo = extractUsageFromChatJson_(json);
    if (usageInfo) {
      appendLog_(
        "USAGE",
        row,
        (api.model ? "model: " + api.model + " | " : "") + usageInfo.summaryLine,
        usageInfo.detailsJson,
        usageInfo.costRub,
        usageInfo.totalTokens
      );
    }

    if (!content) {
      appendLog_("EMPTY_RESPONSE", row, "Пустой ответ от AI API", text);
      return "";
    }
    lastContent = content.toString().trim();
  } catch (e) {
    const errorMessage = e && e.message ? e.message : String(e);
    appendLog_("AI_ERROR", row, errorMessage, "");
    SpreadsheetApp.getActiveSpreadsheet().toast(`Ошибка AI (строка ${row}): ${errorMessage}`, "AI-Редактор", 20);
    Logger.log(e);
    return "";
  }

  // Если модель всё равно вернула запрещённую фразу — обезвреживаем точную подстроку.
  if (lastContent && lastContent.toLowerCase().includes(bannedNeedle)) {
    return lastContent.replace(/смотрите в кино/gi, "можно будет увидеть в кинотеатрах");
  }
  return lastContent || "";
}

/**
 * Новые события — сверху (строка 2, под заголовком). costRub / totalTokens — для Polza usage (руб., токены).
 */
function appendLog_(type, row, message, details, costRub, totalTokens) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureLogSheet_(ss);
  const cr = costRub === undefined || costRub === null ? "" : costRub;
  const tt = totalTokens === undefined || totalTokens === null ? "" : totalTokens;
  sheet.insertRowBefore(2);

  const rowData = [
    new Date(),
    cellScalar_(type),
    cellScalar_(row),
    cellScalar_(message),
    cellScalar_(details),
    cr,
    tt,
  ];
  if (rowData.length !== 7) {
    throw new Error("appendLog_: ожидается 7 колонок");
  }
  writeLogRowCells_(sheet, 2, rowData);
}

/**
 * Polza.ai и др. OpenAI-совместимые ответы: объект usage с cost_rub / cost и токенами.
 * @see https://polza.ai/docs/osobennosti/usage
 */
function extractUsageFromChatJson_(json) {
  const u = json && json.usage;
  if (!u || typeof u !== "object") return null;
  const hasAny =
    u.prompt_tokens != null ||
    u.completion_tokens != null ||
    u.total_tokens != null ||
    u.cost_rub != null ||
    u.cost != null;
  if (!hasAny) return null;

  const costRaw = u.cost_rub != null ? u.cost_rub : u.cost;
  let costRub = "";
  if (costRaw != null && String(costRaw).trim() !== "") {
    const n = Number(costRaw);
    costRub = isNaN(n) ? "" : n;
  }

  let totalTokens = "";
  if (u.total_tokens != null && String(u.total_tokens).trim() !== "") {
    const t = Number(u.total_tokens);
    totalTokens = isNaN(t) ? "" : t;
  }

  const parts = [];
  if (costRub !== "") parts.push("≈ " + String(costRub) + " ₽");
  if (u.total_tokens != null) parts.push("tokens: " + u.total_tokens);
  if (u.prompt_tokens != null) parts.push("in: " + u.prompt_tokens);
  if (u.completion_tokens != null) parts.push("out: " + u.completion_tokens);

  const summaryLine = parts.length ? parts.join(" · ") : "usage";
  return {
    summaryLine: summaryLine,
    detailsJson: JSON.stringify(u),
    costRub: costRub,
    totalTokens: totalTokens,
  };
}

function ensureLogSheet_(ss) {
  let sheet = ss.getSheetByName(LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET_NAME);
  }
  const fullHeader = [
    "timestamp",
    "type",
    "row",
    "message",
    "details",
    "cost_rub",
    "total_tokens",
  ];
  const header = sheet.getRange(1, 1, 1, 7).getValues()[0];
  const hasHeader = header.slice(0, 5).some((v) => String(v || "").trim() !== "");
  if (!hasHeader) {
    writeLogRowCells_(sheet, 1, fullHeader);
    sheet.setFrozenRows(1);
  } else {
    const needCols = !String(header[5] || "").trim() || !String(header[6] || "").trim();
    if (needCols) {
      const merged = header.slice(0, 7);
      while (merged.length < 7) merged.push("");
      for (let i = 0; i < 7; i++) {
        if (!String(merged[i] || "").trim()) merged[i] = fullHeader[i];
      }
      writeLogRowCells_(sheet, 1, merged);
    }
  }
  return sheet;
}

/** Скаляр в ячейку: массивы/объекты → JSON, чтобы не ломать setValues. */
function cellScalar_(v) {
  if (v == null) return "";
  if (v instanceof Date) return v;
  if (typeof v === "object") {
    try {
      return JSON.stringify(v);
    } catch (e) {
      return String(v);
    }
  }
  return String(v);
}

/** Запись строки лога по ячейкам (устойчиво к объединениям и странным размерам диапазона). */
function writeLogRowCells_(sheet, rowIndex, values) {
  const row = values.map((v) => (v instanceof Date ? v : v === "" || v == null ? "" : v));
  for (let c = 0; c < row.length; c++) {
    sheet.getRange(rowIndex, c + 1).setValue(row[c]);
  }
}

