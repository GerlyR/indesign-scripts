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

  var EXC_LIST = Utils.loadTeamExceptions();
  var EXC_ALT = Utils.buildTeamRegex(EXC_LIST, false);
  var EXC_CAP = Utils.buildTeamRegex(EXC_LIST, true);

  function grepOnTarget(tgt, findWhat, changeProps) {
    if (!tgt || !findWhat) return null;
    return Utils.grepChange(tgt, findWhat, changeProps);
  }
  
  function textFindChangeOnTarget(tgt, findWhat, changeTo) {
    if (!tgt || !findWhat) return null;
    return Utils.textChange(tgt, findWhat, changeTo);
  }

  var GROUP = "Таб_фут";
  var pBody = Utils.getParaStyleFromGroup(doc, "body_text", GROUP);
  var pStand = Utils.getParaStyleFromGroup(doc, "И В П М О", GROUP);
  var cBold = Utils.ensureCharStyle(doc, "bold_gazeta", "Bold");
  
  if (!pBody || !pStand || !cBold) {
    alert("Не найдены необходимые стили. Проверьте наличие стилей: body_text, И В П М О, bold_gazeta");
    return;
  }

  var TEAM =
    "(?:«[^»]+»|" + EXC_ALT + "|[A-Za-zА-Яа-яЁё0-9 .\\-]+)" +
    "(?:\\s*[A-ZА-ЯЁa-zа-яё]{1,2}\\.?){0,2}" +
    "\\s*(?:\\([A-Za-zА-ЯЁа-яё\\-\\s]+\\))?";

  var SCORE = "\\d+\\s*:\\s*\\d+";
  var HEADER = "^\\s*" + TEAM + "\\s*-\\s*" + TEAM + "\\s*-\\s*" + SCORE;
  var HEADER_D = "^\\s*" + TEAM + "\\s*-\\s*" + TEAM + "\\s*-\\s*" + SCORE + "\\.";

  var STAND_ROW = "(?m)^\\s*И\\s*[ \\t]*В\\s*[ \\t]*Н\\s*[ \\t]*П\\s*[ \\t]*М\\s*[ \\t]*О\\s*$";
  var STAND_BLOCK1 = "(?mis)" + STAND_ROW.replace("(?m)", "") + "[\\s\\S]*?(?=^\\s*Бомбардиры\\s*:)";
  var STAND_BLOCK2 = "(?mis)" + STAND_ROW.replace("(?m)", "") + "[\\s\\S]*$";

  function getEightBlockRange(fromTarget) {
    if (!fromTarget) return null;
    
    Utils.resetFindGrep();
    try {
      app.findGrepPreferences.findWhat = ".";
      app.findGrepPreferences.appliedCharacterStyle = cBold;
      var hits = fromTarget.findGrep();
      
      if (!hits || !hits.length) {
        Utils.resetFindGrep();
        return null;
      }

      var firstBoldChar = hits[0];
      if (!firstBoldChar || !firstBoldChar.isValid) {
        Utils.resetFindGrep();
        return null;
      }
      
      var story = firstBoldChar.parentStory;
      if (!story || !story.isValid) {
        Utils.resetFindGrep();
        return null;
      }
      
      var endIP = story.insertionPoints[story.insertionPoints.length - 1];

      Utils.resetFindGrep();
      app.findGrepPreferences.findWhat = STAND_ROW;
      var standHits = fromTarget.findGrep();
      if (standHits && standHits.length > 0) {
        try {
          var standHit = standHits[0];
          if (standHit && standHit.isValid) {
            var standPara = standHit.paragraphs[0];
            if (standPara && standPara.isValid) {
              endIP = standPara.insertionPoints[0];
            }
          }
        } catch (e) {}
      }
      
      Utils.resetFindGrep();
      
      try {
        var startIP = firstBoldChar.insertionPoints[0];
        if (startIP && endIP) {
          return story.texts.itemByRange(startIP, endIP);
        }
      } catch (e) {}
      
      return null;
    } catch (e) {
      Utils.resetFindGrep();
      return null;
    }
  }

  function finalBreaksCleanOnAllEight(fromTarget, cBold, pStand) {
    if (!fromTarget || !cBold || !pStand) return;
    
    Utils.resetFindGrep();
    try {
      app.findGrepPreferences.findWhat = "(~n|\\r)";
      var hits = fromTarget.findGrep();

      if (!hits || hits.length === 0) {
        Utils.resetFindGrep();
        return;
      }

      for (var i = 0; i < hits.length; i++) {
        try {
          var ch = hits[i];
          if (!ch || !ch.isValid) continue;
          
          var charObj = ch.characters[0];
          if (!charObj || !charObj.isValid) continue;
          
          var ps = charObj.pointSize;
          var ld = charObj.leading;
          if (!(ps == 8 && (ld == 8 || ld == 8.0))) continue;

          var story = ch.parentStory;
          if (!story || !story.isValid) continue;
          
          var idx = charObj.index;
          if (idx < 0 || idx >= story.characters.length) continue;
          
          var nextChar = (idx + 1 < story.characters.length) ? story.characters[idx + 1] : null;

          var skip = false;
          try {
            if (nextChar && nextChar.isValid && nextChar.appliedCharacterStyle && 
                nextChar.appliedCharacterStyle.name === cBold.name) {
              skip = true;
            }
          } catch (e) {}

          if (!skip && nextChar && nextChar.isValid) {
            try {
              var nextPara = nextChar.paragraphs[0];
              if (nextPara && nextPara.isValid && nextPara.appliedParagraphStyle && 
                  nextPara.appliedParagraphStyle.name === pStand.name) {
                skip = true;
              }
            } catch (e) {}
          }

          if (!skip) {
            try {
              ch.contents = " ";
            } catch (e) {}
          }
        } catch (e) {}
      }
      
      grepOnTarget(fromTarget, "(?: |~S){2,}", {changeTo: " "});
    } catch (e) {}
    Utils.resetFindGrep();
  }

  var target = null;
  try {
    if (app.selection && app.selection.length > 0 && app.selection[0].hasOwnProperty("texts")) {
      target = app.selection[0];
    } else {
      target = doc;
    }
  } catch (e) {
    target = doc;
  }
  
  if (!target) {
    alert("Не найден целевой объект для обработки.");
    return;
  }

  app.doScript(function() {
    try {
      var targetTexts = target.texts;
      if (!targetTexts || targetTexts.length === 0) return;
      
      var paras = targetTexts[0].paragraphs;
      if (!paras) return;
      
      for (var i = 3; i < paras.length; i++) {
        try {
          var para = paras[i];
          if (para && para.isValid) {
            para.applyParagraphStyle(pBody, false);
          }
        } catch (e) {}
      }

      grepOnTarget(target, "«" + EXC_CAP + "»", { changeTo: "$1" });

      grepOnTarget(target, HEADER, { appliedCharacterStyle: cBold });
      grepOnTarget(target, "(" + HEADER + ")(?!\\.)", { changeTo: "$1." });

      var rng = getEightBlockRange(target);
      if (rng && rng.isValid) {
        try {
          rng.pointSize = 8;
          rng.leading = 8;
        } catch (e) {}
        grepOnTarget(rng, "[ ~S]*~n[ ~S]*", { changeTo: " " });
        textFindChangeOnTarget(rng, "^n", " ");
        grepOnTarget(rng, "(?: |~S){2,}", { changeTo: " " });
      }

      Utils.resetFindGrep();
      app.findGrepPreferences.findWhat = "(" + HEADER_D + "[\\s\\S]*?)(?=\\r\\r)";
      var blocks = target.findGrep();
      
      if (blocks && blocks.length > 0) {
        for (var j = 0; j < blocks.length; j++) {
          try {
            var t = blocks[j];
            if (!t || !t.isValid) continue;
            
            var s = String(t.contents);
            s = s.replace(/\r(?!«|(?:ПСЖ)\s*-\s*|(?:ЦСКА)\s*-\s*)/g, " ");
            s = s.replace(/\n/g, " ");
            s = s.replace(/ {2,}/g, " ").replace(/ \./g, ".").replace(/ ,/g, ",");
            t.contents = s;
          } catch (e) {}
        }
      }
      Utils.resetFindGrep();

      var hit = grepOnTarget(target, STAND_BLOCK1, { appliedParagraphStyle: pStand });
      if (!hit || hit.length === 0) {
        grepOnTarget(target, STAND_BLOCK2, { appliedParagraphStyle: pStand });
      }
      grepOnTarget(target, STAND_ROW, { appliedCharacterStyle: cBold });

      Utils.resetFindGrep();
      app.findGrepPreferences.findWhat = "(?m)^\\s*Бомбардиры\\s*:";
      app.changeGrepPreferences.appliedCharacterStyle = cBold;
      var bombMarks = target.changeGrep();
      
      try {
        if (bombMarks && bombMarks.length > 0) {
          for (var b = 0; b < bombMarks.length; b++) {
            try {
              var mark = bombMarks[b];
              if (!mark || !mark.isValid) continue;
              
              var story2 = mark.parentStory;
              if (!story2 || !story2.isValid) continue;
              
              var ip = mark.insertionPoints[mark.insertionPoints.length - 1];
              if (!ip || !ip.isValid) continue;
              
              while (ip.index < story2.insertionPoints.length - 1) {
                try {
                  var charAtIP = story2.characters[ip.index];
                  if (charAtIP && charAtIP.contents === " ") {
                    ip = story2.insertionPoints[ip.index + 1];
                  } else {
                    break;
                  }
                } catch (e) {
                  break;
                }
              }
              
              var endIP = story2.insertionPoints[story2.insertionPoints.length - 1];
              if (ip && endIP) {
                var rng2 = story2.texts.itemByRange(ip, endIP);
                if (rng2 && rng2.isValid) {
                  rng2.pointSize = 8;
                  rng2.leading = 8;
                }
              }
            } catch (e) {}
          }
        }
      } catch (e) {}
      Utils.resetFindGrep();

      grepOnTarget(target, "  +", { changeTo: " " });
      grepOnTarget(target, "\\s+,", { changeTo: "," });
      grepOnTarget(target, "\\s+\\.", { changeTo: "." });
      textFindChangeOnTarget(target, "\r\r", "\r");

      finalBreaksCleanOnAllEight(target, cBold, pStand);

      grepOnTarget(target, "(?: |~S){2,}", { changeTo: " " });

      Utils.cleanupBroom(target);
      Utils.applyStandardReplacements(target);

      Utils.resetFindText();
      Utils.resetFindGrep();

    } catch (e) {
      alert("Ошибка при обработке: " + (e.message || String(e)));
    }
  }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Футбол: формат со странами + INI исключения");
})();
