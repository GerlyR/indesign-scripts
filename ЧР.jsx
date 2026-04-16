#target "InDesign-6.0"

(function () {
  'use strict';

  var _f = File(File($.fileName).parent.fsName + "/CommonUtils.jsx");
  if (!_f.exists) _f = File(File($.fileName).parent.fsName + "\\CommonUtils.jsx");
  if (!_f.exists) { alert("CommonUtils.jsx не найден."); return; }
  try { $.evalFile(_f); } catch (e) { alert("Ошибка загрузки CommonUtils.jsx:\n" + (e.message || e)); return; }
  if (!$.global.CommonUtils || typeof $.global.CommonUtils.getActiveDocument !== 'function') { alert("CommonUtils.jsx загружен некорректно."); return; }
  var Utils = $.global.CommonUtils;

  var doc = Utils.getActiveDocument();
  if (!doc) {
    alert("Откройте документ перед запуском скрипта.");
    return;
  }

  // --- Регулярные выражения ---
  var RE_RUBRIKA = /^футбол\./i;
  var RE_GROUP = /^\s*Группа\s+[A-ZА-ЯЁ]\s*$/i;
  // Лейблы: Гол/Голы, Предупреждения, Удаления, Судья
  var RE_LABEL = /^(Голы?|Предупреждения|Удаления|Судья)\s*:/i;
  var RE_AFTER = /^После\s+матча$/i;
  // Примечание в скобках: (Окончание. Начало на 1-й стр.)
  var RE_PAREN_NOTE = /^\(\s*[^)]+\)\s*\.?\s*$/;
  // Вопрос журналиста: начинается с тире, заканчивается на ?
  var RE_QUESTION = /^\s*[-\u2013\u2014]\s+.*\?\s*$/;

  var story = Utils.getTargetStory(doc);
  if (!story) {
    alert("Не найдена текстовая история для обработки.");
    return;
  }

  // Сброс глобального состояния от предыдущего запуска
  if ($.global) $.global.lastTeams = null;

  var fontInfo = Utils.detectFontVariants(story);
  var CS_BOLD = Utils.ensureCharStyleSmart(doc, "пж", fontInfo.bold);
  var CS_ITALIC = Utils.ensureCharStyleSmart(doc, "tvitalic", fontInfo.italic);
  var CS_BOLDIT = Utils.ensureCharStyleSmart(doc, "tvbolditalic", fontInfo.boldItalic);

  if (!CS_BOLD) {
    alert("Не удалось создать символьный стиль для жирного текста.");
    return;
  }

  // --- Хелперы ---
  function setBoldRange(range) {
    if (!range || !range.isValid) return;
    if (CS_BOLD) {
      try { range.appliedCharacterStyle = CS_BOLD; return; } catch (e) {}
    }
    if (fontInfo && fontInfo.bold) {
      try { range.fontStyle = fontInfo.bold; return; } catch (e) {}
    }
  }

  function boldLabelBeforeColon(par) {
    if (!par || !par.isValid) return;
    try {
      var t = String(par.contents);
      var i = t.indexOf(":");
      if (i > 0 && i < par.characters.length) {
        var rng = par.characters.itemByRange(0, i);
        if (rng && rng.isValid) {
          setBoldRange(rng);
        }
      }
    } catch (e) {}
  }

  // Целый абзац жирным (для имён тренеров)
  function boldEntirePara(par) {
    if (!par || !par.isValid || par.characters.length === 0) return;
    Utils.applyCharStyleToPara(par, CS_BOLD, fontInfo.bold);
  }

  function applyFootballStyles(story) {
    if (!story || !story.isValid || !story.paragraphs || story.paragraphs.length === 0) {
      return;
    }

    Utils.grepChange(story, "\\r{2,}", {changeTo: "\\r"});
    Utils.grepChange(story, "[ \\t\\x{00A0}\\x{202F}\\x{2009}\\x{200A}]{2,}", {changeTo: " "});

    var psRubrika = Utils.getParaStyle(doc, "Rubrika");
    var psZag = Utils.getParaStyle(doc, "Zagolovok");
    var psPodzag = Utils.getParaStyle(doc, "Podzagolovok");
    var psSeraya = Utils.getParaStyle(doc, "seraya");
    var psPress = Utils.getParaStyle(doc, "Press") || Utils.ensureParaStyle(doc, "Press");

    try {
      var p0 = story.paragraphs[0];
      if (p0 && p0.isValid) {
        var p0Text = Utils.trim(String(p0.contents));
        var hasRubrika = RE_RUBRIKA.test(p0Text);
        if (hasRubrika && psRubrika) {
          try { p0.appliedParagraphStyle = psRubrika; } catch (e) {}
        }

        // --- Заголовок (может быть многострочным — все ЗАГЛАВНЫЕ строки) ---
        var zagIdx = hasRubrika ? 1 : 0;
        var zagEnd = zagIdx; // последний абзац заголовка
        if (story.paragraphs.length > zagIdx && psZag) {
          try {
            var zagPara = story.paragraphs[zagIdx];
            if (zagPara && zagPara.isValid) {
              zagPara.appliedParagraphStyle = psZag;
            }
          } catch (e) {}
          // Проверяем продолжение заголовка: короткие строки, преимущественно ЗАГЛАВНЫЕ
          for (var zi = zagIdx + 1; zi < Math.min(zagIdx + 4, story.paragraphs.length); zi++) {
            try {
              var zt = Utils.trim(Utils.getParaText(story.paragraphs[zi]));
              if (zt.length > 0 && zt.length <= 80 && /^[А-ЯЁA-Z\s«»"„"\'\-\u2013\u2014.,!?0-9():]+$/.test(zt)) {
                story.paragraphs[zi].appliedParagraphStyle = psZag;
                zagEnd = zi;
              } else {
                break;
              }
            } catch (e) { break; }
          }
        }

        // --- Счёт матча: сканируем первые 6 абзацев после заголовка ---
        var exceptions = Utils.loadTeamExceptions();
        var teamTok = (function() {
          var parts = ["\u00AB([^\u00BB]+)\u00BB"];
          if (exceptions && exceptions.length > 0) {
            var alts = [];
            for (var i = 0; i < exceptions.length; i++) {
              var ex = Utils.trim(exceptions[i]);
              if (ex) alts.push(Utils.escapeRegex(ex));
            }
            if (alts.length > 0) {
              parts.push("(" + alts.join("|") + ")");
            }
          }
          return "(?:" + parts.join("|") + ")";
        })();

        var DASH = "[-\u2013\u2014\u2212]";
        var CITY = "(?:\\s*\\([^)]+\\))?";
        var RE_SCORE = new RegExp(
          "^\\s*" + teamTok + CITY + "\\s*" + DASH + "\\s*" + teamTok + CITY + "\\s*" + DASH +
          "\\s*(\\d+):(\\d+)(?:\\s*\\([^)]+\\))?\\s*$",
          "i"
        );

        // --- Составы команд: «Команда» (Город): или ЦСКА: ---
        var teamLineupParts = ["\u00AB[^\u00BB]+\u00BB\\s*(?:\\([^)]+\\)\\s*)?"];
        if (exceptions && exceptions.length > 0) {
          for (var ei = 0; ei < exceptions.length; ei++) {
            var exc = Utils.trim(exceptions[ei]);
            if (exc) teamLineupParts.push(Utils.escapeRegex(exc));
          }
        }
        var RE_TEAM_LINEUP = new RegExp("^\\s*(?:" + teamLineupParts.join("|") + ")\\s*:", "i");

        // Ищем строку счёта сканированием (не по фиксированному индексу!)
        var scoreIdx = -1;
        for (var sc = zagEnd + 1; sc < Math.min(zagEnd + 6, story.paragraphs.length); sc++) {
          try {
            var scTxt = Utils.trim(String(story.paragraphs[sc].contents).replace(/\r$/, ""));
            if (RE_SCORE.test(scTxt)) {
              scoreIdx = sc;
              break;
            }
          } catch (e) {}
        }

        // Применяем bold к строке счёта и запоминаем команды
        if (scoreIdx >= 0) {
          try {
            var pMatch = story.paragraphs[scoreIdx];
            if (pMatch && pMatch.isValid) {
              var matchText = Utils.trim(String(pMatch.contents).replace(/\r$/, ""));
              var m = RE_SCORE.exec(matchText);
              if (m) {
                try {
                  var matchChars = pMatch.characters.itemByRange(0, pMatch.characters.length - 1);
                  if (matchChars && matchChars.isValid) {
                    setBoldRange(matchChars);
                  }
                } catch (e) {}

                if (!$.global) $.global = {};
                $.global.lastTeams = {
                  team1: (m[1] || m[2]) || null,
                  team2: (m[3] || m[4]) || null
                };
              }
            }
          } catch (e) {}
        }

        // --- Подзаголовок: всё между заголовком и счётом ---
        if (scoreIdx > zagEnd + 1 && psPodzag) {
          for (var pi = zagEnd + 1; pi < scoreIdx; pi++) {
            try {
              var cand = story.paragraphs[pi];
              if (cand && cand.isValid) {
                var ct = Utils.trim(String(cand.contents).replace(/\r$/, ""));
                // Подзаголовок: не пустой, не начинается с тире, до 200 символов
                if (ct.length > 0 && !/^\s*[-\u2013\u2014]/.test(ct) && ct.length <= 200) {
                  try { cand.appliedParagraphStyle = psPodzag; } catch (e) {}
                }
              }
            } catch (e) {}
          }
        }

        // --- Detect signature from end (1 or 2 lines) ---
        Utils.trimTailEmptyParas(story);
        var sig = Utils.detectSignature(story.paragraphs);
        var sigStartIdx = sig.sigStartIdx;
        var sigEndIdx = sig.sigEndIdx;

        // Подпись → пж+курсив + правое выравнивание
        if (sigStartIdx >= 0) {
          for (var si = sigStartIdx; si <= sigEndIdx; si++) {
            try {
              var sigPara = story.paragraphs[si];
              if (sigPara && sigPara.isValid) {
                try { sigPara.justification = Justification.RIGHT_ALIGN; } catch (e) {}
                Utils.applyCharStyleToPara(sigPara, CS_BOLDIT, fontInfo.boldItalic);
              }
            } catch (e) {}
          }
        }

        var sigIdx = sigStartIdx;

        // --- Основной цикл по абзацам ---
        var inAfterMatch = false; // находимся ли после "ПОСЛЕ МАТЧА"
        var inQuestion = false;   // продолжение многострочного вопроса

        for (var i2 = 0; i2 < story.paragraphs.length; i2++) {
          if (sigStartIdx >= 0 && i2 >= sigStartIdx && i2 <= sigEndIdx) continue;
          try {
            var par = story.paragraphs[i2];
            if (!par || !par.isValid) continue;

            var txt = Utils.trim(String(par.contents).replace(/\r$/, ""));
            if (!txt) continue;

            // --- "ПОСЛЕ МАТЧА" → Press стиль ---
            if (RE_AFTER.test(Utils.normalizeSpaces(txt))) {
              if (psPress) {
                try { par.appliedParagraphStyle = psPress; } catch (e) {}
              }
              inAfterMatch = true;
              continue;
            }

            // --- Группа X → seraya ---
            if (psSeraya && RE_GROUP.test(txt)) {
              try { par.appliedParagraphStyle = psSeraya; } catch (e) {}
              continue;
            }

            // --- Лейблы: Гол(ы), Предупреждения, Удаления, Судья → bold до двоеточия ---
            if (RE_LABEL.test(txt)) {
              boldLabelBeforeColon(par);
              continue;
            }

            // --- Составы команд: «Команда» (Город): → bold до двоеточия ---
            // Только до секции "ПОСЛЕ МАТЧА", иначе перехватывает «Клуба»: в строке тренера
            if (!inAfterMatch && RE_TEAM_LINEUP.test(txt)) {
              boldLabelBeforeColon(par);
              continue;
            }

            // --- Примечание в скобках: (Окончание. Начало на 1-й стр.) → italic ---
            if (RE_PAREN_NOTE.test(txt)) {
              Utils.applyCharStyleToPara(par, CS_ITALIC, fontInfo.italic);
              continue;
            }

            // --- Секция интервью (после "ПОСЛЕ МАТЧА") ---
            if (inAfterMatch) {
              // Вопрос журналиста: строка с тире, заканчивается на ?
              if (RE_QUESTION.test(txt)) {
                Utils.applyCharStyleToPara(par, CS_BOLDIT, fontInfo.boldItalic);
                inQuestion = true;
                continue;
              }

              // Продолжение вопроса на след. строке (не начинается с тире, заканчивается на ?)
              if (inQuestion && !/^\s*[-\u2013\u2014]/.test(txt) && /\?\s*$/.test(txt)) {
                Utils.applyCharStyleToPara(par, CS_BOLDIT, fontInfo.boldItalic);
                continue;
              }
              inQuestion = false;

              // Имя тренера: "Имя ФАМИЛИЯ, должность «Клуб»:"
              // Паттерн: НЕ начинается с тире, заканчивается на двоеточие
              if (!/^\s*[-\u2013\u2014]/.test(txt) && /:\s*$/.test(txt)) {
                // Содержит запятую (имя, должность) или содержит «» (клуб)
                if (/,/.test(txt) || /\u00AB/.test(txt)) {
                  boldEntirePara(par);
                  continue;
                }
              }
            }

          } catch (e) {}
        }

        // --- Блок статистики 8/8pt ---
        // Ищем от строки счёта до строки с "зрител" включительно
        if (scoreIdx >= 0) {
          var statsStart = scoreIdx;
          var statsEnd = -1;
          for (var st = scoreIdx; st < Math.min(scoreIdx + 20, story.paragraphs.length); st++) {
            try {
              var stTxt = String(story.paragraphs[st].contents);
              if (/зрител/i.test(stTxt)) {
                statsEnd = st;
                break;
              }
            } catch (e) {}
          }
          if (statsEnd >= statsStart) {
            for (var sb = statsStart; sb <= statsEnd; sb++) {
              try {
                var sp = story.paragraphs[sb];
                if (sp && sp.isValid && sp.characters.length > 0) {
                  var srng = sp.characters.itemByRange(0, sp.characters.length - 1);
                  if (srng && srng.isValid) {
                    try { srng.pointSize = 8; } catch (e) {}
                    try { srng.leading = 8; } catch (e) {}
                  }
                }
              } catch (e) {}
            }
          }
        }

        // --- Ремарки эмоций курсивом ---
        Utils.applyEmotionRemarks(story, CS_ITALIC);

        Utils.cleanupBroom(story);
        Utils.applyStandardReplacements(story);
      }
    } catch (e) {
      alert("Ошибка при обработке: " + (e.message || String(e)));
    }
  }

  app.doScript(function() {
    applyFootballStyles(story);
  }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "\u0427\u0420");

  $.global.applyFootballStyles = applyFootballStyles;
})();
