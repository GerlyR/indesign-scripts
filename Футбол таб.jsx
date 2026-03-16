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
  
  function grepOn(tgt, find, change, findProps) {
    if (!tgt || !find) return null;
    
    Utils.resetFindGrep();
    try {
      if (findProps) {
        for (var k in findProps) {
          if (findProps.hasOwnProperty(k)) {
            try {
              app.findGrepPreferences[k] = findProps[k];
            } catch (e) {}
          }
        }
      }
      app.findGrepPreferences.findWhat = find;
      if (change) {
        for (var c in change) {
          if (change.hasOwnProperty(c)) {
            try {
              app.changeGrepPreferences[c] = change[c];
            } catch (e) {}
          }
        }
      }
      return tgt.changeGrep();
    } catch (e) {
      return null;
    } finally {
      Utils.resetFindGrep();
    }
  }
  
  function textOn(tgt, find, to) {
    if (!tgt || !find) return;
    Utils.textChange(tgt, find, to);
  }
  
  var GROUP = "Таб_фут";
  var pBody = Utils.getParaStyleFromGroup(doc, "body_text", GROUP);
  var pStand = Utils.getParaStyleFromGroup(doc, "И В П М О", GROUP);
  var pGrey = Utils.getParaStyleFromGroup(doc, "seraya", GROUP);
  var cBold = Utils.ensureCharStyle(doc, "bold_gazeta", "Bold");
  
  if (!pBody || !pStand || !pGrey || !cBold) {
    alert("Не найдены необходимые стили. Проверьте наличие стилей: body_text, И В П М О, seraya, bold_gazeta");
    return;
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

  var P_COUNTRY = "(?mi)^\\s*[A-Za-zА-ЯЁа-яё\\-\\s]+\\.\\s*\\d+\\s*[-–]?\\s*й\\s+тур\\b.*$";
  var P_TEAM = "(?:«[^»]+»|" + EXC_ALT + "|[A-Za-zА-Яа-яЁё0-9 .\\-]+)";
  var P_SCORE = "\\d+\\s*:\\s*\\d+";
  var P_HEADER = "^\\s*" + P_TEAM + "\\s*-\\s*" + P_TEAM + "\\s*-\\s*" + P_SCORE;
  var P_HEADERP = P_HEADER + "\\.";
  var P_STANDROW = "(?m)^\\s*И\\s*[ \\t]*В\\s*[ \\t]*Н\\s*[ \\t]*П\\s*[ \\t]*М\\s*[ \\t]*О\\s*$";
  var P_STANDB1 = "(?mis)" + P_STANDROW.replace("(?m)", "") + "[\\s\\S]*?(?=^\\s*Бомбардиры\\s*:)";
  var P_STANDB2 = "(?mis)" + P_STANDROW.replace("(?m)", "") + "[\\s\\S]*$";

  var RE_NOJOIN = new RegExp("\\r(?!«|(?:" + EXC_ALT + ")\\s*-\\s*)", "g");

  app.doScript(function() {
    try {
      try {
        var targetTexts = target.texts;
        if (targetTexts && targetTexts.length > 0) {
          targetTexts[0].applyParagraphStyle(pBody, false);
        }
      } catch (e) {}
      
      grepOn(target, P_COUNTRY, {appliedParagraphStyle: pGrey});

      grepOn(target, "«" + EXC_CAP + "»", {changeTo: "$1"});

      grepOn(target, P_HEADER, {appliedCharacterStyle: cBold});
      grepOn(target, "(" + P_HEADER + ")(?!\\.)", {changeTo: "$1."});

      var h = grepOn(target, P_STANDB1, {appliedParagraphStyle: pStand});
      if (!h || h.length === 0) {
        grepOn(target, P_STANDB2, {appliedParagraphStyle: pStand});
      }
      grepOn(target, P_STANDROW, {appliedCharacterStyle: cBold});

      Utils.resetFindGrep();
      app.findGrepPreferences.findWhat = "(?m)^\\s*Бомбардиры\\s*:";
      app.changeGrepPreferences.appliedCharacterStyle = cBold;
      var bombMarks = target.changeGrep();

      Utils.resetFindGrep();
      app.findGrepPreferences.findWhat = "(" + P_HEADERP + "[\\s\\S]*?)(?=\\r\\r)";
      var blocks = target.findGrep();
      
      if (blocks && blocks.length > 0) {
        for (var i = 0; i < blocks.length; i++) {
          try {
            var t = blocks[i];
            if (!t || !t.isValid) continue;
            
            var s = String(t.contents);
            s = s.replace(RE_NOJOIN, " ").replace(/\n/g, " ").replace(/ {2,}/g, " ").replace(/ \./g, ".").replace(/ ,/g, ",");
            t.contents = s;
          } catch (e) {}
        }
      }
      Utils.resetFindGrep();

      Utils.resetFindGrep();
      app.findGrepPreferences.findWhat = "(" + P_HEADERP + "[\\s\\S]*?)(?=\\r\\r)";
      var sizeBlocks = target.findGrep();
      
      if (sizeBlocks && sizeBlocks.length > 0) {
        for (var j = 0; j < sizeBlocks.length; j++) {
          try {
            var bl = sizeBlocks[j];
            if (!bl || !bl.isValid) continue;
            
            var txt = String(bl.contents);
            var m = txt.match(/-\s*\d+\s*:\s*\d+\./);
            if (!m) continue;
            
            var endIdx = txt.indexOf(m[0]) + m[0].length;
            if (endIdx < 0 || endIdx >= txt.length) continue;
            
            var story = bl.parentStory;
            if (!story || !story.isValid) continue;
            
            try {
              var startIP = bl.insertionPoints[endIdx];
              var endIP = bl.insertionPoints[bl.insertionPoints.length - 1];
              if (startIP && endIP) {
                var rng = story.texts.itemByRange(startIP, endIP);
                if (rng && rng.isValid) {
                  rng.pointSize = 8;
                  rng.leading = 8;
                }
              }
            } catch (e) {}
          } catch (e) {}
        }
      }
      Utils.resetFindGrep();

      try {
        if (bombMarks && bombMarks.length > 0) {
          for (var b = 0; b < bombMarks.length; b++) {
            try {
              var mark = bombMarks[b];
              if (!mark || !mark.isValid) continue;
              
              var story = mark.parentStory;
              if (!story || !story.isValid) continue;
              
              var ip = mark.insertionPoints[mark.insertionPoints.length - 1];
              if (!ip || !ip.isValid) continue;
              
              while (ip.index < story.insertionPoints.length - 1) {
                try {
                  var charAtIP = story.characters[ip.index];
                  if (charAtIP && charAtIP.contents === " ") {
                    ip = story.insertionPoints[ip.index + 1];
                  } else {
                    break;
                  }
                } catch (e) {
                  break;
                }
              }
              
              var endIP = story.insertionPoints[story.insertionPoints.length - 1];
              if (ip && endIP) {
                var rng2 = story.texts.itemByRange(ip, endIP);
                if (rng2 && rng2.isValid) {
                  rng2.pointSize = 8;
                  rng2.leading = 8;
                }
              }
            } catch (e) {}
          }
        }
      } catch (e) {}

      var fixes = [["  +", " "], ["\\s+,", ","], ["\\s+\\.", "."]];
      for (var f = 0; f < fixes.length; f++) {
        grepOn(target, fixes[f][0], {changeTo: fixes[f][1]});
      }

      Utils.textChange(target, "\r\r", "\r");

      var compressMap = {
        "«Хоффенхайм»": 87, "«Унион Берлин»": 82, "«Райо Вальекано»": 75,
        "«Реал Сосьедад»": 80, "«Кристал Пэлас»": 80, "«Ман. Юнайтед»": 83,
        "«Ноттингем Ф.»": 86, "«Вулверхэмптон»": 77
      };
      
      for (var key in compressMap) {
        if (compressMap.hasOwnProperty(key)) {
          grepOn(target, Utils.escapeRegex(key), {horizontalScale: compressMap[key]}, {appliedParagraphStyle: pStand});
        }
      }

      Utils.cleanupBroom(target);
      Utils.applyStandardReplacements(target);

      Utils.resetFindText();
      Utils.resetFindGrep();

    } catch (e) {
      alert("Ошибка при обработке: " + (e.message || String(e)));
    }
  }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Gazeta (relative INI)");

})();
