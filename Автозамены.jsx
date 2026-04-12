(function () {
  'use strict';

  // --- Config (capture BEFORE $.evalFile which clobbers $.fileName) ---
  var SCRIPT_DIR = File($.fileName).parent.fsName;

  // --- Load CommonUtils ---
  var _f = File(SCRIPT_DIR + "/CommonUtils.jsx");
  if (!_f.exists) _f = File(SCRIPT_DIR + "\\CommonUtils.jsx");
  if (!_f.exists) { alert("CommonUtils.jsx \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D."); return; }
  try { $.evalFile(_f); } catch (e) { alert("\u041E\u0448\u0438\u0431\u043A\u0430 \u0437\u0430\u0433\u0440\u0443\u0437\u043A\u0438 CommonUtils.jsx:\n" + (e.message || e)); return; }
  if (!$.global.CommonUtils || typeof $.global.CommonUtils.getActiveDocument !== 'function') { alert("CommonUtils.jsx \u0437\u0430\u0433\u0440\u0443\u0436\u0435\u043D \u043D\u0435\u043A\u043E\u0440\u0440\u0435\u043A\u0442\u043D\u043E."); return; }
  var Utils = $.global.CommonUtils;
  var TEMP_DIR   = Folder.temp.fsName;
  var RUN_ID     = String((new Date()).getTime()) + "_" + Math.floor(Math.random() * 1000000);

  var PYTHON_EXE = Utils.findPython(SCRIPT_DIR);

  // --- Sync editorial rules from Excel ---
  function syncRulesFromExcel() {
    if (!PYTHON_EXE) return;
    var syncPy = SCRIPT_DIR + "\\sync_editorial_rules.py";
    if (!File(syncPy).exists) return;
    var xlsx = SCRIPT_DIR + "\\editorial_rules.xlsx";
    if (!File(xlsx).exists) return;

    var txtPath = SCRIPT_DIR + "\\editorial_rules.txt";
    var txtFile = File(txtPath);
    var oldMtime = 0;
    if (txtFile.exists) {
      try { oldMtime = txtFile.modified.getTime(); } catch (e) {}
    }

    var batPath = TEMP_DIR + "\\indd_sync_" + RUN_ID + ".bat";
    var bat = File(batPath);
    bat.encoding = "CP1251";
    if (!bat.open("w")) return;
    bat.writeln("@echo off");
    bat.writeln("chcp 65001 >nul");
    bat.writeln('"' + PYTHON_EXE + '" "' + syncPy + '"');
    bat.close();

    var launched = File(batPath).execute();
    if (!launched) {
      try { File(batPath).remove(); } catch (e) {}
      return;
    }

    var waited = 0;
    while (waited < 20) {
      $.sleep(500);
      waited++;
      try {
        txtFile = File(txtPath);
        if (txtFile.exists && txtFile.modified.getTime() > oldMtime) break;
      } catch (e) {}
    }
    try { File(batPath).remove(); } catch (e) {}
  }

  // --- Load config file lines ---
  function loadConfigLines(filename) {
    var f = File(SCRIPT_DIR + "\\" + filename);
    if (!f.exists) return [];
    f.encoding = "UTF-8";
    if (!f.open("r")) return [];
    var lines = [];
    try {
      while (!f.eof) {
        var line = f.readln();
        if (!line) continue;
        line = line.replace(/^\s+|\s+$/g, "");
        if (!line || line.charAt(0) === "#" || line.charAt(0) === ";") continue;
        lines.push(line);
      }
    } finally {
      f.close();
    }
    return lines;
  }

  // --- Load editorial rules ---
  function loadEditorialRules() {
    var lines = loadConfigLines("editorial_rules.txt");
    var rules = [];
    for (var i = 0; i < lines.length; i++) {
      var parts = lines[i].split("\t");
      if (parts.length >= 3 && parts[0].toUpperCase() === "GREP" && parts[1] && parts[2]) {
        rules.push({ type: "grep", find: parts[1], replace: parts[2] });
      } else if (parts[0] && parts[0].toUpperCase() === "GREP") {
        // GREP line with missing replacement — skip, don't fall through to text
        continue;
      } else if (parts.length >= 2 && parts[0] && parts[1]) {
        rules.push({ type: "text", find: parts[0], replace: parts[1] });
      }
    }
    return rules;
  }

  // --- Apply editorial auto-replacements ---
  function applyEditorialRules(story, rules) {
    var count = 0;
    var details = [];
    for (var i = 0; i < rules.length; i++) {
      try {
        var r = rules[i];
        var found = [];
        if (r.type === "grep") {
          Utils.resetFindGrep();
          app.findGrepPreferences.findWhat = r.find;
          try {
            var hits = story.findGrep();
            for (var h = 0; h < hits.length; h++) {
              found.push(String(hits[h].contents));
            }
          } catch (e) {}
          Utils.resetFindGrep();
          var result = Utils.grepChange(story, r.find, { changeTo: r.replace });
        } else {
          Utils.resetFindText();
          app.findTextPreferences.findWhat = r.find;
          try {
            var hits = story.findText();
            for (var h = 0; h < hits.length; h++) {
              found.push(String(hits[h].contents));
            }
          } catch (e) {}
          Utils.resetFindText();
          var result = Utils.textChange(story, r.find, r.replace);
        }
        if (result && result.length) {
          count += result.length;
          for (var d = 0; d < found.length; d++) {
            details.push("  \u00AB" + found[d] + "\u00BB \u2192 \u00AB" + r.replace + "\u00BB");
          }
        }
      } catch (e) {
        details.push("  \u26A0 \u041F\u0440\u0430\u0432\u0438\u043B\u043E #" + (i + 1) + " (" + r.find.substring(0, 30) + "): " + (e.message || e));
      }
    }
    return { count: count, details: details };
  }

  // --- Load obscene patterns ---
  function loadObscenePatterns() {
    var lines = loadConfigLines("obscene_patterns.txt");
    var patterns = [];
    for (var i = 0; i < lines.length; i++) {
      patterns.push(lines[i].toLowerCase());
    }
    return patterns;
  }

  // --- Check obscene line breaks and fix with noBreak ---
  function checkObsceneBreaks(doc, story, patterns) {
    if (!patterns.length) return { count: 0, words: [] };

    var letterRe = /[a-zA-Z\u0410-\u044F\u0451\u0401]/;
    var wordCharRe = /[a-zA-Z\u0410-\u044F\u0451\u0401\u00AD]/;
    var fixRanges = [];

    var containers = story.textContainers;
    for (var ci = 0; ci < containers.length; ci++) {
      var frame = containers[ci];
      if (!frame || !frame.isValid) continue;

      var frameLines = frame.lines;
      for (var li = 0; li < frameLines.length - 1; li++) {
        try {
          var curLine = frameLines[li];
          var nxtLine = frameLines[li + 1];
          if (!curLine.isValid || !nxtLine.isValid) continue;
          if (curLine.characters.length === 0 || nxtLine.characters.length === 0) continue;

          var lastCh = String(curLine.characters[-1].contents);
          var firstCh = String(nxtLine.characters[0].contents);

          if (!lastCh.match(letterRe) || !firstCh.match(letterRe)) continue;

          var lineText = String(curLine.contents).replace(/[\u00AD]/g, "");
          var endFrag = "";
          for (var k = lineText.length - 1; k >= 0; k--) {
            var ch = lineText.charAt(k);
            if (ch.match(/[a-zA-Z\u0410-\u044F\u0451\u0401\-]/)) {
              endFrag = ch + endFrag;
            } else { break; }
          }
          endFrag = endFrag.replace(/[\-]+$/, "");

          var nextText = String(nxtLine.contents).replace(/[\u00AD]/g, "");
          var startFrag = "";
          for (var k2 = 0; k2 < nextText.length; k2++) {
            var ch2 = nextText.charAt(k2);
            if (ch2.match(/[a-zA-Z\u0410-\u044F\u0451\u0401\-]/)) {
              startFrag += ch2;
            } else { break; }
          }
          startFrag = startFrag.replace(/^[\-]+/, "");

          var endLower = endFrag.toLowerCase();
          var startLower = startFrag.toLowerCase();
          var found = false;
          for (var pi = 0; pi < patterns.length; pi++) {
            var pat = patterns[pi];
            if (endLower.length >= pat.length &&
                endLower.lastIndexOf(pat) === endLower.length - pat.length) {
              found = true; break;
            }
            if (startLower.length >= pat.length &&
                startLower.indexOf(pat) === 0) {
              found = true; break;
            }
          }

          if (found) {
            var lc = curLine.characters;
            var wStart = lc.length - 1;
            while (wStart > 0) {
              var c = String(lc[wStart - 1].contents);
              if (c.match(wordCharRe)) { wStart--; } else { break; }
            }

            var nc = nxtLine.characters;
            var wEnd = 0;
            while (wEnd < nc.length - 1) {
              var c2 = String(nc[wEnd + 1].contents);
              if (c2.match(wordCharRe)) { wEnd++; } else { break; }
            }

            var wordText = "";
            for (var wi = wStart; wi < lc.length; wi++) {
              wordText += String(lc[wi].contents);
            }
            wordText = wordText.replace(/[\u00AD\-]+$/, "");
            for (var wi2 = 0; wi2 <= wEnd; wi2++) {
              wordText += String(nc[wi2].contents);
            }
            wordText = wordText.replace(/[\u00AD]/g, "");

            fixRanges.push({ s: lc[wStart], e: nc[wEnd], word: wordText });
          }
        } catch (e) {}
      }
    }

    // Apply in reverse order so reflow doesn't invalidate later refs
    for (var i = fixRanges.length - 1; i >= 0; i--) {
      try {
        var r = fixRanges[i];
        if (r.s.isValid && r.e.isValid) {
          story.characters.itemByRange(r.s, r.e).noBreak = true;
        }
      } catch (e) {}
    }

    var words = [];
    for (var i = 0; i < fixRanges.length; i++) {
      if (fixRanges[i].word) words.push(fixRanges[i].word);
    }
    return { count: fixRanges.length, words: words };
  }

  // ====== MAIN ======
  var doc = Utils.getActiveDocument();
  if (!doc) { alert("\u041E\u0442\u043A\u0440\u043E\u0439\u0442\u0435 \u0434\u043E\u043A\u0443\u043C\u0435\u043D\u0442."); return; }

  var story = Utils.getTargetStory(doc);
  if (!story) { alert("\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u0442\u0435\u043A\u0441\u0442\u043E\u0432\u044B\u0439 \u0444\u0440\u0435\u0439\u043C."); return; }

  syncRulesFromExcel();

  var editRules = loadEditorialRules();
  var obscenePatterns = loadObscenePatterns();

  var editResult = { count: 0, details: [] };
  var obsceneResult = { count: 0, words: [] };

  // Step 1: Auto-replacements (separate undo)
  if (editRules.length) {
    app.doScript(function () {
      editResult = applyEditorialRules(story, editRules);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT,
      "\u0410\u0432\u0442\u043E\u0437\u0430\u043C\u0435\u043D\u044B");
  }

  // Step 2: Obscene break fixes (separate undo)
  if (obscenePatterns.length) {
    app.doScript(function () {
      obsceneResult = checkObsceneBreaks(doc, story, obscenePatterns);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT,
      "\u0421\u043A\u043B\u0435\u0439\u043A\u0430 \u043F\u0435\u0440\u0435\u043D\u043E\u0441\u043E\u0432");
  }

  // --- Store results for batch mode or show report ---
  $.global.__editResult = editResult;
  $.global.__obsceneResult = obsceneResult;

  if ($.global.__BATCH_MODE) return; // wrapper will show combined report

  var editCount = editResult.count;
  var obsceneCount = obsceneResult.count;
  var msg = [];

  if (editCount > 0) {
    msg.push("\u0410\u0432\u0442\u043E\u0437\u0430\u043C\u0435\u043D: " + editCount);
    var maxShow = Math.min(editResult.details.length, 20);
    for (var di = 0; di < maxShow; di++) {
      msg.push(editResult.details[di]);
    }
    if (editResult.details.length > maxShow) {
      msg.push("  ... \u0438 \u0435\u0449\u0451 " + (editResult.details.length - maxShow));
    }
  }

  if (obsceneCount > 0) {
    msg.push("\u0421\u043A\u043B\u0435\u0435\u043D\u043E \u043F\u0435\u0440\u0435\u043D\u043E\u0441\u043E\u0432: " + obsceneCount);
    for (var wi = 0; wi < obsceneResult.words.length; wi++) {
      msg.push("  \u2192 " + obsceneResult.words[wi]);
    }
  }

  if (editCount === 0 && obsceneCount === 0) {
    msg.push("\u0418\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u0439 \u043D\u0435\u0442.");
  }

  msg.push("\nCtrl+Z \u0434\u043B\u044F \u043E\u0442\u043C\u0435\u043D\u044B");
  alert(msg.join("\n"));
})();
