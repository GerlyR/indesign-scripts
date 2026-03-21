(function(){
  'use strict';
  app.doScript(function() { _main(); }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "\u0422\u0412-\u0433\u0438\u0434");
  function _main() {

  var _f = File(File($.fileName).parent.fsName + "/CommonUtils.jsx");
  if (!_f.exists) _f = File(File($.fileName).parent.fsName + "\\CommonUtils.jsx");
  if (!_f.exists) { alert("CommonUtils.jsx не найден."); return; }
  try { $.evalFile(_f); } catch (e) { alert("Ошибка загрузки CommonUtils.jsx:\n" + (e.message || e)); return; }
  if (!$.global.CommonUtils || typeof $.global.CommonUtils.getActiveDocument !== 'function') { alert("CommonUtils.jsx загружен некорректно."); return; }
  var Utils = $.global.CommonUtils;

  var doc = Utils.getActiveDocument();
  if (!doc) { alert("Откройте документ."); return; }

  if (!app.selection || app.selection.length === 0) {
    alert("Выделите текст или текстовые фреймы.");
    return;
  }

  // --- Стили абзацев ---
  var psPress  = Utils.ensureParaStyle(doc, "Press");
  var psSeraya = Utils.ensureParaStyle(doc, "seraya");
  if (!psPress || !psSeraya) {
    alert("Не удалось создать необходимые стили.");
    return;
  }

  var stories = Utils.getSelectedStories();
  if (!stories.length) {
    alert("Не вижу текст в выделении.");
    return;
  }

  // --- Автодетекция шрифта из первой story ---
  var fontInfo = Utils.detectFontVariants(stories[0]);
  var csBold   = Utils.ensureCharStyleSmart(doc, "tvBold", fontInfo.bold);
  var csItalic = Utils.ensureCharStyleSmart(doc, "tvItalic", fontInfo.italic);

  if (!csBold || !csItalic) {
    alert("Не удалось создать символьные стили.");
    return;
  }

  // --- Загрузка команд-исключений из INI ---
  var teamExceptions = Utils.loadTeamExceptions();

  // --- Вспомогательные функции ---
  var UPPER_CHARS = "АБВГДЕЁЖЗИЙКЛМНОПРСТУФХЦЧШЩЪЫЬЭЮЯABCDEFGHIJKLMNOPQRSTUVWXYZ".toUpperCase();

  // Пропуск точек после инициалов: "Б. Иванов", "A. Griezmann"
  function isInitialDot(t, dotPos) {
    if (!t || dotPos < 1) return false;
    var prev = t.charAt(dotPos - 1);
    var prev2 = (dotPos >= 2) ? t.charAt(dotPos - 2) : "";
    return (UPPER_CHARS.indexOf(prev.toUpperCase()) >= 0)
      && (!prev2 || prev2 === " " || prev2 === "\u00A0" || prev2 === ".");
  }

  // Регулярка для канала + время: «Матч ТВ», 23:00 или «Матч ТВ» 23:00
  var RE_CHANNEL_TIME = /^\u00AB[^\u00BB]+\u00BB[,]?\s*\d{1,2}:\d{2}/;

  // Регулярка для пары команд: «Team» – «Team» или Team — Team
  // Строим динамически с учётом команд-исключений из INI
  var teamAlts = [];
  for (var ti = 0; ti < teamExceptions.length; ti++) {
    var exItem = Utils.trim(teamExceptions[ti]);
    if (exItem) teamAlts.push(Utils.escapeRegex(exItem));
  }
  var teamToken = "\u00AB[^\u00BB]+\u00BB"; // «...»
  if (teamAlts.length > 0) {
    teamToken = "(?:\u00AB[^\u00BB]+\u00BB|" + teamAlts.join("|") + ")";
  }
  // Сокращение города после команды: «Динамо» Мх, «Динамо» М, «Торпедо» Мск
  var cityAbbr = "(?:\\s+[А-ЯЁ][а-яё]*\\.?)?";
  var pairRe = new RegExp("^" + teamToken + cityAbbr + "\\s*[\\-\u2013\u2014\u2012]\\s*" + teamToken + cityAbbr);

  // Определяет, является ли сегмент каналом + временем: «Матч ТВ», 23:00
  function isChannelSegment(trimmed) {
    return RE_CHANNEL_TIME.test(trimmed);
  }

  // Определяет, является ли сегмент парой команд: «Ювентус» – «Сити»
  function isPairSegment(trimmed) {
    return pairRe.test(trimmed);
  }

  // Пропуск URL-сегментов: www, .com, .ru и т.п.
  var RE_URL = /^(?:www\b|https?:|.*\.(?:com|ru|net|org|рф)\b)/i;

  // Inline-дата: "7 декабря, суббота" / "8 декабря, воскресенье"
  var RE_INLINE_DATE = /^\d{1,2}\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s*,\s*(понедельник|вторник|среда|четверг|пятница|суббота|воскресенье)/i;

  // --- Шаг 1: схлопывание двойных переносов ---
  for (var s = 0; s < stories.length; s++) {
    Utils.grepChange(stories[s], "\\r\\r+", {changeTo: "\\r"});
  }

  // --- Шаг 2: размер шрифта 8/8 ---
  for (var s1 = 0; s1 < stories.length; s1++) {
    try {
      var st = stories[s1];
      if (!st || !st.isValid) continue;
      var allChars = st.characters.everyItem();
      if (allChars) {
        allChars.pointSize = 8;
        allChars.leading = 8;
      }
    } catch (e) {}
  }

  // --- Шаг 3: первый абзац → Press ---
  for (var s2 = 0; s2 < stories.length; s2++) {
    try {
      var st = stories[s2];
      if (!st || !st.isValid || !st.paragraphs || st.paragraphs.length === 0) continue;
      var fp = st.paragraphs[0];
      if (fp && fp.isValid) fp.appliedParagraphStyle = psPress;
    } catch (e) {}
  }

  // --- Шаг 4: даты → seraya ---
  var dateFindWhat = "(?m)^\\s*(Понедельник|Вторник|Среда|Четверг|Пятница|Суббота|Воскресенье)\\s*,\\s*\\d{1,2}\\s+" +
    "(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s*\\.?\\s*$";
  var dateParaIds = {};

  for (var s3 = 0; s3 < stories.length; s3++) {
    try {
      var st = stories[s3];
      if (!st || !st.isValid) continue;
      Utils.resetFindGrep();
      app.findGrepPreferences.findWhat = dateFindWhat;
      var hits = st.findGrep();
      if (hits && hits.length > 0) {
        for (var h = 0; h < hits.length; h++) {
          try {
            var hit = hits[h];
            if (!hit || !hit.isValid) continue;
            var pDate = hit.paragraphs[0];
            if (pDate && pDate.isValid) {
              pDate.appliedParagraphStyle = psSeraya;
              dateParaIds[pDate.id] = true;
            }
          } catch (e) {}
        }
      }
      Utils.resetFindGrep();
    } catch (e) {}
  }

  // --- Шаг 5: разметка сегментов (sport→bold, channel→italic, match→skip) ---
  for (var s4 = 0; s4 < stories.length; s4++) {
    try {
      var st = stories[s4];
      if (!st || !st.isValid) continue;
      var paras = st.paragraphs.everyItem().getElements();
      if (!paras || paras.length === 0) continue;

      for (var i = 0; i < paras.length; i++) {
        try {
          var p = paras[i];
          if (!p || !p.isValid) continue;

          // Пропуск абзацев с назначенными стилями
          try {
            if (p.appliedParagraphStyle === psPress || p.appliedParagraphStyle === psSeraya) continue;
          } catch (e) {}

          // Пропуск дат по id
          try { if (dateParaIds[p.id]) continue; } catch (e) {}

          var t = String(p.contents).replace(/\r$/, "");
          if (!t || /^\s*$/.test(t)) continue;

          // Пропуск URL-строк целиком
          if (RE_URL.test(Utils.trim(t))) continue;

          // --- Разбиваем текст по точкам на сегменты ---
          var segments = [];  // {start, end, text}
          var idx = 0;
          while (idx < t.length) {
            var dotPos = t.indexOf(".", idx);
            if (dotPos === -1) {
              // Хвост после последней точки
              var tail = t.substring(idx);
              if (Utils.trim(tail).length > 0) {
                segments.push({start: idx, end: t.length - 1, text: tail});
              }
              break;
            }

            // Пропуск точек после инициалов
            if (isInitialDot(t, dotPos)) {
              idx = dotPos + 1;
              continue;
            }

            // Пропуск многоточия
            if (dotPos + 1 < t.length && t.charAt(dotPos + 1) === ".") {
              idx = dotPos + 1;
              continue;
            }

            var seg = t.substring(idx, dotPos + 1);
            segments.push({start: idx, end: dotPos, text: seg});
            idx = dotPos + 1;
          }

          // --- Классифицируем и стилизуем каждый сегмент ---
          for (var si = 0; si < segments.length; si++) {
            var seg = segments[si];
            var trimmed = Utils.trim(seg.text);
            if (!trimmed) continue;

            // Пара команд → пропуск (без стилизации, обычный текст)
            if (isPairSegment(trimmed)) continue;

            // URL-фрагменты → пропуск
            if (RE_URL.test(trimmed)) continue;

            // Inline-дата → пропуск (обычный текст, не bold)
            if (RE_INLINE_DATE.test(trimmed)) continue;

            // Канал + время → курсив
            if (isChannelSegment(trimmed)) {
              Utils.applyCharStyleToRange(p, seg.start, seg.end, csItalic, fontInfo.italic);
              continue;
            }

            // Иначе → жирный (вид спорта, турнир, описание)
            Utils.applyCharStyleToRange(p, seg.start, seg.end, csBold, fontInfo.bold);
          }
        } catch (e) {}
      }
    } catch (e) {}
  }

  // --- Шаг 6: каналы в скобках → курсив: («Матч ТВ» - 13:50) ---
  for (var s5 = 0; s5 < stories.length; s5++) {
    // Канал в скобках с тире: («Матч Премьер» - 13:50) или («Матч Премьер», 13:50)
    Utils.grepChange(stories[s5], "\\(\u00AB[^\u00BB]+\u00BB\\s*[-\u2013\u2014,]\\s*\\d{1,2}:\\d{2}\\)", {appliedCharacterStyle: csItalic});
    // Канал через запятую без скобок: «Матч ТВ», 23:00
    Utils.grepChange(stories[s5], "\u00AB[^\u00BB]+\u00BB,\\s*\\d{1,2}:\\d{2}", {appliedCharacterStyle: csItalic});
    // Время HH:MM (подстраховка для оставшихся)
    Utils.grepChange(stories[s5], "\\b\\d{1,2}:\\d{2}\\b", {appliedCharacterStyle: csItalic});
  }

  // --- Шаг 7: текстовые замены и очистка ---
  for (var s6 = 0; s6 < stories.length; s6++) {
    try {
      var st = stories[s6];
      if (!st || !st.isValid) continue;

      Utils.grepChange(st, "ТВ-ГИД\\.|ТВ-гид\\.", {changeTo: "ТВ-ГИД"});
      Utils.grepChange(st, "(Понедельник|Вторник|Среда|Четверг|Пятница|Суббота|Воскресенье)\\s*,\\s*(\\d{1,2})\\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s*\\.", {changeTo: "$1, $2 $3"});

      Utils.cleanupBroom(st);
    } catch (e) {}
  }

  Utils.resetFindGrep();
  } // end _main
})();
