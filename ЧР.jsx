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

  // Алиас для удобства
  function applyCharStyleToPara(para, charStyle, fontStyleName) {
    Utils.applyCharStyleToPara(para, charStyle, fontStyleName);
  }

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

  var RE_RUBRIKA = /^футбол\./i;
  var RE_GROUP = /^\s*Группа\s+[A-ZА-ЯЁ]\s*$/i;
  var RE_LABEL = /^(Голы|Предупреждения|Судья)\s*:/i;
  var RE_AFTER = /^После\s+матча$/i;
  var RE_INTERVIEW_LINE_END = /:\s*$/;
  var RE_BLOCK_8_8 = /Голы:\s*[\s\S]*?\b(?:зрител(?:я|ей)|зрители\.)\b[^\r]*\r/;

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

  if (!CS_BOLD) {
    alert("Не удалось создать символьный стиль для жирного текста.");
    return;
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

        var zagIdx = hasRubrika ? 1 : 0;
        if (story.paragraphs.length > zagIdx && psZag) {
          try {
            var zagPara = story.paragraphs[zagIdx];
            if (zagPara && zagPara.isValid) {
              zagPara.appliedParagraphStyle = psZag;
            }
          } catch (e) {}
        }

        var podIdx = zagIdx + 1;
        var hasPod = false;
        if (story.paragraphs.length > podIdx && psPodzag) {
          try {
            var cand = story.paragraphs[podIdx];
            if (cand && cand.isValid) {
              var t = Utils.trim(String(cand.contents));
              if (t.length > 0 && !/^\s*[-.–«"']/i.test(t) && !/[.?!…]\s*$/.test(t) && t.length <= 140) {
                try {
                  cand.appliedParagraphStyle = psPodzag;
                  hasPod = true;
                } catch (e) {}
              }
            }
          } catch (e) {}
        }

        // --- Detect signature from end (1 or 2 lines) ---
        // Форматы: "Имя ФАМИЛИЯ." или "Имя ФАМИЛИЯ,\nиз Города."
        Utils.trimTailEmptyParas(story);
        var sigStartIdx = -1; // первая строка подписи
        var sigEndIdx = -1;   // последняя строка подписи (может совпадать)
        var pLen = story.paragraphs.length;

        if (pLen >= 2) {
          try {
            var lastTxt = Utils.trim(Utils.getParaText(story.paragraphs[pLen - 1]));
            var prevTxt = Utils.trim(Utils.getParaText(story.paragraphs[pLen - 2]));

            if (Utils.isSignatureCity(lastTxt) && Utils.isSignature(prevTxt)) {
              // Двухстрочная: "Имя ФАМИЛИЯ,\nиз Города."
              sigStartIdx = pLen - 2;
              sigEndIdx = pLen - 1;
            } else if (Utils.isSignature(lastTxt)) {
              // Однострочная: "Имя ФАМИЛИЯ."
              sigStartIdx = pLen - 1;
              sigEndIdx = pLen - 1;
            }
          } catch (e) {}
        } else if (pLen === 1) {
          try {
            var onlyTxt = Utils.trim(Utils.getParaText(story.paragraphs[0]));
            if (Utils.isSignature(onlyTxt)) {
              sigStartIdx = 0;
              sigEndIdx = 0;
            }
          } catch (e) {}
        }

        // Style signature lines: right-align + italic
        if (sigStartIdx >= 0) {
          for (var si = sigStartIdx; si <= sigEndIdx; si++) {
            try {
              var sigPara = story.paragraphs[si];
              if (sigPara && sigPara.isValid) {
                try { sigPara.justification = Justification.RIGHT_ALIGN; } catch (e) {}
                applyCharStyleToPara(sigPara, CS_ITALIC, fontInfo.italic);
              }
            } catch (e) {}
          }
        }

        // sigIdx для пропуска в циклах ниже
        var sigIdx = sigStartIdx;

        var exceptions = Utils.loadTeamExceptions();
        var teamTok = (function() {
          var parts = ["«([^»]+)»"];
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

        var DASH = "[\\u2013\\u2014\\u2212-]";
        var RE_SCORE = new RegExp(
          "^\\s*" + teamTok + "\\s*" + DASH + "\\s*" + teamTok + "\\s*" + DASH +
          "\\s*(\\d+):(\\d+)\\s*\\(\\s*(\\d+):(\\d+)\\s*\\)\\s*$",
          "i"
        );

        var matchLineIdx = hasPod ? (zagIdx + 2) : (zagIdx + 1);
        if (story.paragraphs.length > matchLineIdx) {
          try {
            var pMatch = story.paragraphs[matchLineIdx];
            if (pMatch && pMatch.isValid) {
              var matchText = Utils.trim(String(pMatch.contents));
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

        Utils.resetFindGrep();
        try {
          app.findGrepPreferences.findWhat = RE_BLOCK_8_8.source;
          var hits = story.findGrep();
          if (hits && hits.length > 0) {
            try {
              var hit = hits[0];
              if (hit && hit.isValid) {
                try {
                  hit.pointSize = 8;
                  hit.leading = 8;
                } catch (e) {}
              }
            } catch (e) {}
          }
        } catch (e) {}
        Utils.resetFindGrep();

        if (psSeraya) {
          for (var i1 = 0; i1 < story.paragraphs.length; i1++) {
            try {
              var pG = story.paragraphs[i1];
              if (pG && pG.isValid) {
                var groupText = Utils.trim(String(pG.contents));
                if (RE_GROUP.test(groupText)) {
                  try { pG.appliedParagraphStyle = psSeraya; } catch (e) {}
                }
              }
            } catch (e) {}
          }
        }

        var team1 = ($.global && $.global.lastTeams) ? $.global.lastTeams.team1 : null;
        var team2 = ($.global && $.global.lastTeams) ? $.global.lastTeams.team2 : null;

        for (var i2 = 0; i2 < story.paragraphs.length; i2++) {
          if (sigStartIdx >= 0 && i2 >= sigStartIdx && i2 <= sigEndIdx) continue; // skip signature
          try {
            var par = story.paragraphs[i2];
            if (!par || !par.isValid) continue;

            var txt = Utils.trim(String(par.contents));

            if (RE_LABEL.test(txt)) {
              boldLabelBeforeColon(par);
              continue;
            }

            if (team1 || team2) {
              var names = [];
              if (team1) names.push(team1);
              if (team2) names.push(team2);

              for (var k = 0; k < names.length; k++) {
                var nm = names[k];
                if (!nm) continue;
                var re = new RegExp("^\\s*(?:«\\s*" + Utils.escapeRegex(nm) + "\\s*»|" + Utils.escapeRegex(nm) + ")\\s*:");
                if (re.test(txt)) {
                  boldLabelBeforeColon(par);
                  break;
                }
              }
            }
          } catch (e) {}
        }

        if (psPress) {
          for (var i3 = 0; i3 < story.paragraphs.length; i3++) {
            try {
              var pA = story.paragraphs[i3];
              if (pA && pA.isValid) {
                var afterText = Utils.trim(Utils.normalizeSpaces(String(pA.contents)));
                if (RE_AFTER.test(afterText)) {
                  try { pA.appliedParagraphStyle = psPress; } catch (e) {}
                }
              }
            } catch (e) {}
          }
        }

        for (var i4 = 0; i4 < story.paragraphs.length; i4++) {
          if (sigStartIdx >= 0 && i4 >= sigStartIdx && i4 <= sigEndIdx) continue; // skip signature
          try {
            var pq = story.paragraphs[i4];
            if (!pq || !pq.isValid) continue;

            var line = Utils.normalizeSpaces(String(pq.contents).replace(/\r$/, ""));
            if (/^\s*[-–—]/.test(line) || !RE_INTERVIEW_LINE_END.test(line) || !/,/.test(line)) {
              continue;
            }

            var colon = line.lastIndexOf(":");
            var comma = line.indexOf(",");
            if (comma === -1 || colon === -1 || comma > colon) continue;

            try {
              if (colon < pq.characters.length) {
                var interviewRng = pq.characters.itemByRange(0, colon);
                if (interviewRng && interviewRng.isValid) {
                  setBoldRange(interviewRng);
                }
              }
            } catch (e) {}
          } catch (e) {}
        }
      }

      // --- Ремарки эмоций курсивом: (смеётся), (улыбаясь) и т.п. ---
      Utils.applyEmotionRemarks(story, CS_ITALIC);

      Utils.cleanupBroom(story);
      Utils.applyStandardReplacements(story);
    } catch (e) {
      alert("Ошибка при обработке: " + (e.message || String(e)));
    }
  }

  applyFootballStyles(story);

  if (!$.global) $.global = {};
  $.global.applyFootballStyles = applyFootballStyles;
})();
