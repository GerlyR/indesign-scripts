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

  var PS_PRESS = Utils.ensureParaStyle(doc, "Press");
  var PS_ZAG = Utils.ensureParaStyle(doc, "Zagolovok");
  var PS_POD = Utils.ensureParaStyle(doc, "Podzagolovok");
  if (!PS_PRESS || !PS_ZAG || !PS_POD) {
    alert("Не удалось создать необходимые стили абзацев.");
    return;
  }

  var RE_TWO_WORDS_BEFORE_COLON = /^([\s«"'"‚'']*\S+)\s+(\S+)\s*:(.*)$/;
  var RE_ANY_PAREN_NOTE = /^\(\s*[^)]*\)\s*\.?$/;

  var story = Utils.getTargetStory(doc);
  if (!story) {
    alert("Не найдена текстовая история для обработки.");
    return;
  }

  var fontInfo = Utils.detectFontVariants(story);

  var CS_TVBOLD = Utils.ensureCharStyleSmart(doc, "tvbold", fontInfo.bold);
  var CS_TVITALIC = Utils.ensureCharStyleSmart(doc, "tvitalic", fontInfo.italic);
  var CS_TVBOLDIT = Utils.ensureCharStyleSmart(doc, "tvbolditalic", fontInfo.boldItalic);

  app.doScript(function() { _run(); }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "\u0418\u043D\u0442\u0435\u0440\u0432\u044C\u044E");
  function _run() {
  try {
    Utils.removeWhitespaceParas(story);
    Utils.collapseBreaks(story);

    var paras = story.paragraphs;
    if (!paras || paras.length === 0) {
      alert("Нет абзацев для обработки.");
      return;
    }

    try {
      if (paras[0] && paras[0].isValid) {
        paras[0].appliedParagraphStyle = PS_PRESS;
      }
    } catch (e) {}

    if (paras.length >= 2) {
      try {
        var zag = paras[1];
        if (zag && zag.isValid) {
          zag.appliedParagraphStyle = PS_ZAG;
          var tz = Utils.getParaText(zag);
          var m = tz.match(RE_TWO_WORDS_BEFORE_COLON);

          if (m) {
            var colon = tz.indexOf(":");
            if (colon >= 0 && colon < zag.characters.length) {
              try {
                var rng = zag.characters.itemByRange(0, colon);
                if (rng && rng.isValid && CS_TVITALIC) {
                  rng.appliedCharacterStyle = CS_TVITALIC;
                }
              } catch (e) {}
            }

            var raw = m[1];
            var pure = raw.replace(/^[\s\u00A0]+/, "");
            var a = raw.length - pure.length;
            var b = a + pure.length - 1;

            if (a >= 0 && b < zag.characters.length && b >= a) {
              try {
                var r = zag.characters.itemByRange(a, b);
                if (r && r.isValid) {
                  r.capitalization = Capitalization.NORMAL;
                  var cap = Utils.capitalizeFirst(pure);
                  if (cap !== pure) {
                    r.contents = cap;
                  }
                }
              } catch (e) {}
            }
          }
        }
      } catch (e) {}
    }

    // --- Detect signature and parenthetical note FIRST (from end) ---
    // Поддержка: "Имя ФАМИЛИЯ." или "Имя ФАМИЛИЯ,\nиз Города." + опц. "(примечание)"
    var parenIdx = -1;
    var sigStartIdx = -1;
    var sigEndIdx = -1;

    // Сначала ищем parenthetical note с конца
    for (var j = paras.length - 1; j >= 0; j--) {
      try {
        var para = paras[j];
        if (!para || !para.isValid) continue;
        var s = Utils.trim(Utils.getParaText(para));
        if (!s) continue;
        if (RE_ANY_PAREN_NOTE.test(s)) {
          parenIdx = j;
        }
        break;
      } catch (e) {}
    }

    // Определяем где искать подпись
    var sigSearchEnd = (parenIdx >= 0) ? parenIdx : paras.length;

    // Проверяем двухстрочный формат: "Имя ФАМИЛИЯ,\nиз Города."
    if (sigSearchEnd >= 2) {
      try {
        var cand1Txt = Utils.trim(Utils.getParaText(paras[sigSearchEnd - 1]));
        var cand2Txt = Utils.trim(Utils.getParaText(paras[sigSearchEnd - 2]));
        if (Utils.isSignatureCity(cand1Txt) && Utils.isSignature(cand2Txt)) {
          sigStartIdx = sigSearchEnd - 2;
          sigEndIdx = sigSearchEnd - 1;
        }
      } catch (e) {}
    }

    // Если двухстрочная не найдена, проверяем однострочную
    if (sigStartIdx < 0 && sigSearchEnd >= 1) {
      try {
        var candTxt = Utils.trim(Utils.getParaText(paras[sigSearchEnd - 1]));
        if (Utils.isSignature(candTxt)) {
          sigStartIdx = sigSearchEnd - 1;
          sigEndIdx = sigSearchEnd - 1;
        }
      } catch (e) {}
    }

    // Совместимость: sigIdx = первая строка подписи (для paintEnd)
    var sigIdx = sigStartIdx;

    // --- Scan for subheaders and dash-lines, skipping signature/paren ---
    var dashIdx = [];
    var lastHeaderIdx = 1;

    function isSignatureRange(idx) {
      if (sigStartIdx >= 0 && idx >= sigStartIdx && idx <= sigEndIdx) return true;
      if (idx === parenIdx) return true;
      return false;
    }

    for (var i = 2; i < paras.length; i++) {
      if (isSignatureRange(i)) continue;
      try {
        var p = paras[i];
        if (!p || !p.isValid) continue;

        var t = Utils.getParaText(p);
        if (Utils.beginsWithDash(t)) {
          dashIdx.push(i);
        } else {
          try {
            if (Utils.isSubheader(t)) {
              p.appliedParagraphStyle = PS_POD;
              lastHeaderIdx = i;
            }
          } catch (e) {}
        }
      } catch (e) {}
    }

    // Алиас для удобства
    function applyCharStyleToPara(para, charStyle, fontStyleName) {
      Utils.applyCharStyleToPara(para, charStyle, fontStyleName);
    }

    // --- Bold lead ---
    var bodyIdx = lastHeaderIdx + 1;
    if (bodyIdx < paras.length && bodyIdx !== sigIdx && bodyIdx !== parenIdx) {
      try {
        var bodyPara = paras[bodyIdx];
        if (bodyPara && bodyPara.isValid && !Utils.beginsWithDash(Utils.getParaText(bodyPara))) {
          applyCharStyleToPara(bodyPara, CS_TVBOLD, fontInfo.bold);
        }
      } catch (e) {}
    }

    // --- Auto-detect: interviewer speaks LESS in total → paint interviewer ---
    var paintFirst = true;
    if (dashIdx.length >= 2) {
      var evenLen = 0, oddLen = 0;
      var limEnd = paras.length;
      if (sigIdx >= 0 && sigIdx < limEnd) limEnd = sigIdx;
      if (parenIdx >= 0 && parenIdx < limEnd) limEnd = parenIdx;

      for (var d0 = 0; d0 < dashIdx.length; d0++) {
        var bStart = dashIdx[d0];
        var bEnd = (d0 + 1 < dashIdx.length) ? dashIdx[d0 + 1] : limEnd;
        if (bEnd > limEnd) bEnd = limEnd;
        var bLen = 0;
        for (var r0 = bStart; r0 < bEnd; r0++) {
          try { bLen += Utils.getParaText(paras[r0]).length; } catch (e) {}
        }
        if (d0 % 2 === 0) evenLen += bLen;
        else oddLen += bLen;
      }
      // Короткий набор = журналист → красим его
      paintFirst = (evenLen < oddLen);
    }

    // --- Paint interviewee replies bold+italic ---
    var paintEnd = paras.length;
    if (sigIdx >= 0) paintEnd = sigIdx;
    if (parenIdx >= 0 && parenIdx < paintEnd) paintEnd = parenIdx;

    for (var d = 0; d < dashIdx.length; d++) {
      var paint = paintFirst ? (d % 2 === 0) : (d % 2 === 1);
      if (!paint) continue;

      var blockA = dashIdx[d];
      var blockB = (d + 1 < dashIdx.length) ? dashIdx[d + 1] : paintEnd;
      if (blockB > paintEnd) blockB = paintEnd;

      for (var k = blockA; k < blockB; k++) {
        try {
          applyCharStyleToPara(paras[k], CS_TVBOLDIT, fontInfo.boldItalic);
        } catch (e) {}
      }
    }

    // --- Apply signature styling (1 or 2 lines) ---
    if (sigStartIdx >= 0) {
      for (var si = sigStartIdx; si <= sigEndIdx; si++) {
        try {
          var sigPara = paras[si];
          if (sigPara && sigPara.isValid) {
            try { sigPara.justification = Justification.RIGHT_ALIGN; } catch (e) {}
            applyCharStyleToPara(sigPara, CS_TVBOLDIT, fontInfo.boldItalic);
          }
        } catch (e) {}
      }
    }

    if (parenIdx >= 0 && parenIdx < paras.length) {
      try {
        var parenPara = paras[parenIdx];
        if (parenPara && parenPara.isValid) {
          try { parenPara.justification = Justification.RIGHT_ALIGN; } catch (e) {}
          applyCharStyleToPara(parenPara, CS_TVITALIC, fontInfo.italic);
        }
      } catch (e) {}
    }

    // --- Ремарки эмоций курсивом: (смеётся), (улыбаясь), (усмехается) и т.п. ---
    Utils.applyEmotionRemarks(story, CS_TVITALIC);

    Utils.collapseBreaks(story);
    Utils.resetFindGrep();

    Utils.grepChange(story, "\\?[ \\t\\x{00A0}]*\\.", {changeTo: "?"});
    Utils.grepChange(story, "\\.[ \\t\\x{00A0}]*\\.", {changeTo: "."});

    Utils.cleanupBroom(story);
    Utils.applyStandardReplacements(story);

  } catch (e) {
    alert("Ошибка при обработке: " + (e.message || String(e)));
  }
  } // end _run
})();
