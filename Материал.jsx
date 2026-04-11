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
  var PS_POD = Utils.ensureParaStyle(doc, "Podzagolovok_13 на 80");

  if (!PS_PRESS || !PS_ZAG || !PS_POD) {
    alert("Не удалось создать необходимые стили абзацев.");
    return;
  }

  var story = Utils.getTargetStory(doc);
  if (!story) {
    alert("Не найдена текстовая история для обработки.");
    return;
  }

  var fontInfo = Utils.detectFontVariants(story);

  var CS_BOLD = Utils.ensureCharStyleSmart(doc, "tvbold", fontInfo.bold);
  var CS_ITALIC = Utils.ensureCharStyleSmart(doc, "tvitalic", fontInfo.italic);
  var CS_BOLDIT = Utils.ensureCharStyleSmart(doc, "tvbolditalic", fontInfo.boldItalic);


  app.doScript(function() { _run(); }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B");
  function _run() {
  try {
    Utils.grepChange(story, "\\r\\r+", {changeTo: "\\r"});

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
        if (paras[1] && paras[1].isValid) {
          paras[1].appliedParagraphStyle = PS_ZAG;
        }
      } catch (e) {}
    }

    // --- Detect signature from end FIRST (1 or 2 lines) ---
    Utils.trimTailEmptyParas(story);
    if (!story.paragraphs || story.paragraphs.length === 0) return;

    var sigStartIdx = -1;
    var sigEndIdx = -1;
    var pLen = story.paragraphs.length;

    if (pLen >= 2) {
      try {
        var lastTxt2 = Utils.trim(Utils.getParaText(story.paragraphs[pLen - 1]));
        var prevTxt2 = Utils.trim(Utils.getParaText(story.paragraphs[pLen - 2]));
        if (Utils.isSignatureCity(lastTxt2) && Utils.isSignature(prevTxt2)) {
          sigStartIdx = pLen - 2;
          sigEndIdx = pLen - 1;
        } else if (Utils.isSignature(lastTxt2)) {
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

    var hasSig = (sigStartIdx >= 0);

    // --- Scan for subheaders, skipping signature lines ---
    var firstSub = -1;
    for (var i = 2; i < paras.length; i++) {
      if (hasSig && i >= sigStartIdx && i <= sigEndIdx) continue;
      try {
        var p = paras[i];
        if (!p || !p.isValid) continue;

        var t = Utils.getParaText(p);
        try {
          if (Utils.isSubheader(t)) {
            p.appliedParagraphStyle = PS_POD;
            if (firstSub < 0) firstSub = i;
          }
        } catch (e) {}
      } catch (e) {}
    }

    // Bold lead: itemByRange вместо everyItem
    if (firstSub === 2) {
      if (paras.length >= 4) {
        try {
          Utils.applyCharStyleToPara(paras[3], CS_BOLD, fontInfo.bold);
        } catch (e) {}
      }
    } else {
      if (paras.length >= 3) {
        try {
          Utils.applyCharStyleToPara(paras[2], CS_BOLD, fontInfo.bold);
        } catch (e) {}
      }
    }

    function placeSignature(baseIdx, wordsTwo, bracketText) {
      if (baseIdx < 0 || baseIdx >= story.paragraphs.length) return;

      try {
        var baseP = story.paragraphs[baseIdx];
        if (!baseP || !baseP.isValid) return;
        var baseTxt = Utils.trim(Utils.getParaText(baseP));
        baseP.contents = baseTxt.replace(/[\s\u00A0]+$/, "") + "\r";
      } catch (e) { return; }

      Utils.trimTailEmptyParas(story);
      if (!story.paragraphs || story.paragraphs.length === 0) return;

      var parts = [];
      if (wordsTwo && Utils.trim(wordsTwo).length) {
        parts.push(Utils.trim(wordsTwo));
      }
      if (bracketText && Utils.trim(bracketText).length) {
        parts.push(Utils.trim(bracketText));
      }

      var signLine = Utils.trim(parts.join(" "));
      if (signLine && !/[.!?…]$/.test(signLine)) {
        signLine += ".";
      }

      try {
        var lastIP = story.insertionPoints[story.insertionPoints.length - 1];
        if (lastIP && lastIP.isValid) {
          lastIP.contents = signLine;
        }
      } catch (e) {}

      // Стилизуем все абзацы подписи (может быть 2 при двухстрочной подписи с городом)
      if (story.paragraphs.length > 0) {
        for (var sp = baseIdx + 1; sp < story.paragraphs.length; sp++) {
          try {
            var sign = story.paragraphs[sp];
            if (sign && sign.isValid) {
              sign.justification = Justification.RIGHT_ALIGN;
              Utils.applyCharStyleToPara(sign, CS_BOLDIT, fontInfo.boldItalic);
            }
          } catch (e) {}
        }
      }

      Utils.trimTailEmptyParas(story);
    }

    if (hasSig) {
      // Собираем текст подписи (1 или 2 строки)
      var sigLines = [];
      for (var si = sigStartIdx; si <= sigEndIdx; si++) {
        try {
          sigLines.push(Utils.trim(Utils.getParaText(story.paragraphs[si])));
        } catch (e) {}
      }
      var sigText = sigLines.join(" ");
      var mTwoWords = sigText.match(/^\s*(\S+)\s+(\S+)/);
      var twoWords = mTwoWords ? (mTwoWords[1] + " " + mTwoWords[2]) : sigText.replace(/[.!?…,]+$/, "");
      // Если 2-строчная: добавляем город
      if (sigEndIdx > sigStartIdx) {
        var cityTxt = Utils.trim(sigLines[sigLines.length - 1]);
        if (cityTxt) twoWords = twoWords.replace(/[,]+$/, "") + ",\r" + cityTxt;
      }
      var bracketFromPrev = "";

      // Ищем скобки в абзаце перед подписью
      if (sigStartIdx - 1 >= 0) {
        try {
          var prevP = story.paragraphs[sigStartIdx - 1];
          if (prevP && prevP.isValid) {
            var prevTxt = Utils.getParaText(prevP);
            var exPrev = Utils.extractEndingBrackets(prevTxt);
            if (exPrev.brk) {
              bracketFromPrev = exPrev.brk;
              prevP.contents = Utils.trim(exPrev.base);
            }
          }
        } catch (e) {}
      }

      // Удаляем строки подписи (с конца чтобы индексы не поехали)
      for (var ri = sigEndIdx; ri >= sigStartIdx; ri--) {
        try { story.paragraphs[ri].remove(); } catch (e) {}
      }

      var newBaseIdx = (story.paragraphs.length > 0) ? story.paragraphs.length - 1 : 0;
      placeSignature(newBaseIdx, twoWords, bracketFromPrev);

    } else {
      var lastP2 = story.paragraphs[story.paragraphs.length - 1];
      var lastTxt3 = Utils.trim(Utils.getParaText(lastP2));
      var ex = Utils.extractEndingBrackets(lastTxt3);
      if (ex.brk) {
        try { lastP2.contents = Utils.trim(ex.base); } catch (e) {}
        placeSignature(story.paragraphs.length - 1, "", ex.brk);
      }
    }

    // --- Ремарки эмоций курсивом: (смеётся), (улыбаясь) и т.п. ---
    Utils.applyEmotionRemarks(story, CS_ITALIC);

    Utils.cleanupBroom(story);
    Utils.applyStandardReplacements(story);

  } catch (e) {
    alert("Ошибка при обработке: " + (e.message || String(e)));
  }
  } // end _run
})();
