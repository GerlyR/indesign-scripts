(function () {
  'use strict';

  // --- Load CommonUtils ---
  var _f = File(File($.fileName).parent.fsName + "/CommonUtils.jsx");
  if (!_f.exists) _f = File(File($.fileName).parent.fsName + "\\CommonUtils.jsx");
  if (!_f.exists) { alert("CommonUtils.jsx not found."); return; }
  try { $.evalFile(_f); } catch (e) { alert("Error loading CommonUtils.jsx:\n" + (e.message || e)); return; }
  var Utils = $.global.CommonUtils;

  // --- Find Python ---
  function findPython() {
    var sd = File($.fileName).parent.fsName;
    // 1. Config file python_path.txt in script folder
    var cf = File(sd + "\\python_path.txt");
    if (cf.exists) {
      cf.encoding = "UTF-8";
      if (cf.open("r")) {
        var p = cf.read().replace(/[\r\n\s]+/g, "");
        cf.close();
        if (p && File(p).exists) return p;
      }
    }
    // 2. Auto-detect from standard Windows locations
    var vers = ["313", "312", "311", "310", "39", "38"];
    var dirs = [];
    var la = $.getenv("LOCALAPPDATA");
    if (la) for (var i = 0; i < vers.length; i++) dirs.push(la + "\\Programs\\Python\\Python" + vers[i]);
    var pf = $.getenv("ProgramFiles");
    if (pf) for (var i = 0; i < vers.length; i++) dirs.push(pf + "\\Python" + vers[i]);
    for (var i = 0; i < vers.length; i++) dirs.push("C:\\Python" + vers[i]);
    for (var i = 0; i < dirs.length; i++) {
      var f = File(dirs[i] + "\\python.exe");
      if (f.exists) return f.fsName;
    }
    return null;
  }

  // --- Config ---
  var PYTHON_EXE = findPython();
  var SCRIPT_DIR = File($.fileName).parent.fsName;
  var WORKER_PY  = SCRIPT_DIR + "\\spellcheck_worker.py";
  var TEMP_DIR   = Folder.temp.fsName;
  var RUN_ID = String((new Date()).getTime()) + "_" + Math.floor(Math.random() * 1000000);
  var INPUT_FILE  = TEMP_DIR + "\\indd_spell_input_" + RUN_ID + ".json";
  var OUTPUT_FILE = TEMP_DIR + "\\indd_spell_output_" + RUN_ID + ".json";
  var BAT_FILE    = TEMP_DIR + "\\indd_spell_run_" + RUN_ID + ".bat";

  // --- JSON serializer ---
  function toJSON(obj) {
    if (obj === null || obj === undefined) return "null";
    if (typeof obj === "number" || typeof obj === "boolean") return String(obj);
    if (typeof obj === "string") {
      var s = obj.replace(/\\/g, "\\\\").replace(/"/g, '\\"')
                 .replace(/\n/g, "\\n").replace(/\r/g, "\\r")
                 .replace(/\t/g, "\\t").replace(/\x08/g, "\\b")
                 .replace(/\x0C/g, "\\f");
      // Escape remaining control chars as \u00XX
      var out = "";
      for (var ci = 0; ci < s.length; ci++) {
        var cc = s.charCodeAt(ci);
        if (cc < 0x20) {
          var hex = cc.toString(16);
          out += "\\u" + ("0000" + hex).slice(-4);
        } else {
          out += s.charAt(ci);
        }
      }
      return '"' + out + '"';
    }
    if (obj instanceof Array) {
      var a = [];
      for (var i = 0; i < obj.length; i++) a.push(toJSON(obj[i]));
      return "[" + a.join(",") + "]";
    }
    if (typeof obj === "object") {
      var p = [];
      for (var k in obj) {
        if (obj.hasOwnProperty(k)) p.push(toJSON(k) + ":" + toJSON(obj[k]));
      }
      return "{" + p.join(",") + "}";
    }
    return String(obj);
  }

  // --- Create color swatch ---
  function getColor(doc, name, r, g, b) {
    // Check existing
    for (var i = 0; i < doc.colors.length; i++) {
      try {
        if (doc.colors[i].name === name) return doc.colors[i];
      } catch (e) {}
    }
    // Create
    var c = doc.colors.add();
    c.name = name;
    c.model = ColorModel.PROCESS;
    c.space = ColorSpace.RGB;
    c.colorValue = [r, g, b];
    return c;
  }

  // --- Extract paragraphs ---
  function extractParas(story) {
    var out = [];
    for (var i = 0; i < story.paragraphs.length; i++) {
      try {
        var p = story.paragraphs[i];
        if (!p || !p.isValid) continue;
        var t = String(p.contents).replace(/\r$/, "");
        if (t.replace(/[\s\u00A0]+/g, "")) out.push({ idx: i, text: t });
      } catch (e) {}
    }
    return out;
  }

  function cleanupTempFiles() {
    try { File(INPUT_FILE).remove(); } catch (e) {}
    try { File(OUTPUT_FILE).remove(); } catch (e) {}
    try { File(BAT_FILE).remove(); } catch (e) {}
  }

  // --- Call Python ---
  function runPython() {
    var bat = File(BAT_FILE);
    bat.encoding = "CP1251";
    if (!bat.open("w")) { alert("\u041D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u0441\u043E\u0437\u0434\u0430\u0442\u044C bat."); return false; }
    bat.writeln("@echo off");
    bat.writeln("chcp 65001 >nul");
    var cmd = '"' + PYTHON_EXE + '" "' + WORKER_PY + '" "' + INPUT_FILE + '" "' + OUTPUT_FILE + '"';
    bat.writeln(cmd);
    bat.close();

    File(BAT_FILE).execute();

    // Wait for output
    var waited = 0;
    var out = File(OUTPUT_FILE);
    while (!out.exists && waited < 120) {
      $.sleep(1000);
      waited++;
      out = File(OUTPUT_FILE);
    }
    try { File(BAT_FILE).remove(); } catch (e) {}

    if (!out.exists) {
      alert("\u0422\u0430\u0439\u043C\u0430\u0443\u0442 (2 \u043C\u0438\u043D).");
      return false;
    }
    $.sleep(300);
    return true;
  }

  // --- Read result ---
  function readResult() {
    var f = File(OUTPUT_FILE);
    if (!f.exists) return null;
    f.encoding = "UTF-8";
    if (!f.open("r")) return null;
    var s = f.read();
    f.close();
    try {
      if (typeof JSON !== "undefined" && JSON.parse) {
        return JSON.parse(s);
      }
    } catch (e) {}
    try { return eval("(" + s + ")"); } catch (e2) { return null; }
  }

  // --- Annotate inline ---
  function annotate(doc, story, matches) {
    var red   = getColor(doc, "_SC_Red",   220, 40, 40);
    var green = getColor(doc, "_SC_Green", 30, 140, 30);
    var gray  = getColor(doc, "_SC_Gray",  140, 140, 140);

    // Only with replacements, sort reverse
    var items = [];
    for (var i = 0; i < matches.length; i++) {
      var m = matches[i];
      if (m.replacements && m.replacements.length > 0) items.push(m);
    }
    items.sort(function (a, b) {
      return a.para !== b.para ? b.para - a.para : b.offset - a.offset;
    });

    var n = 0;
    for (var i = 0; i < items.length; i++) {
      try {
        var m = items[i];
        var para = story.paragraphs[m.para];
        if (!para || !para.isValid) continue;

        var pLen = para.characters.length;
        // Skip if paragraph ends with just \r (1 char)
        if (pLen <= 1) continue;
        if (m.offset + m.length > pLen) continue;

        // Get error range
        var c0 = para.characters[m.offset];
        var c1 = para.characters[m.offset + m.length - 1];
        if (!c0.isValid || !c1.isValid) continue;

        var rng = para.characters.itemByRange(c0, c1);
        if (!rng.isValid) continue;

        // Color error red + underline
        rng.fillColor = red;
        rng.underline = true;

        // Insert suggestion after error: "→fix"
        var fix = m.replacements[0];
        var tag = "\u2192" + fix;  // →fix

        // Insert at the insertion point after the last error character
        var ip = c1.insertionPoints[-1];
        if (!ip.isValid) continue;
        ip.contents = tag;

        // Now style the inserted characters
        // They start at position (m.offset + m.length) in the paragraph
        var tagStart = m.offset + m.length;
        var tagLen = tag.length;

        // Style arrow "→" gray
        try {
          para.characters[tagStart].fillColor = gray;
        } catch (e) {}

        // Style suggestion green
        if (tagLen > 1) {
          try {
            var s0 = para.characters[tagStart + 1];
            var s1 = para.characters[tagStart + tagLen - 1];
            var sr = para.characters.itemByRange(s0, s1);
            sr.fillColor = green;
          } catch (e) {}
        }

        n++;
      } catch (e) {}
    }
    return n;
  }

  // --- Sync editorial rules from Excel ---
  function syncRulesFromExcel() {
    if (!PYTHON_EXE) return;
    var syncPy = SCRIPT_DIR + "\\sync_editorial_rules.py";
    if (!File(syncPy).exists) return;
    var xlsx = SCRIPT_DIR + "\\editorial_rules.xlsx";
    if (!File(xlsx).exists) return;

    // Remember old mtime of txt to know when sync is done
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

    File(batPath).execute();

    // Wait until txt file is updated or timeout (max 10 sec)
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
    while (!f.eof) {
      var line = f.readln();
      if (!line) continue;
      line = line.replace(/^\s+|\s+$/g, "");
      if (!line || line.charAt(0) === "#" || line.charAt(0) === ";") continue;
      lines.push(line);
    }
    f.close();
    return lines;
  }

  // --- Load editorial rules ---
  // Supports two formats:
  //   find\treplace         — plain text replacement
  //   GREP\tpattern\treplace — GREP (regex) replacement
  function loadEditorialRules() {
    var lines = loadConfigLines("editorial_rules.txt");
    var rules = [];
    for (var i = 0; i < lines.length; i++) {
      var parts = lines[i].split("\t");
      if (parts.length >= 3 && parts[0].toUpperCase() === "GREP" && parts[1] && parts[2]) {
        rules.push({ type: "grep", find: parts[1], replace: parts[2] });
      } else if (parts.length >= 2 && parts[0] && parts[1]) {
        rules.push({ type: "text", find: parts[0], replace: parts[1] });
      }
    }
    return rules;
  }

  // --- Apply editorial auto-replacements ---
  function applyEditorialRules(story, rules) {
    var count = 0;
    for (var i = 0; i < rules.length; i++) {
      try {
        var r = rules[i];
        var result;
        if (r.type === "grep") {
          result = Utils.grepChange(story, r.find, { changeTo: r.replace });
        } else {
          result = Utils.textChange(story, r.find, r.replace);
        }
        if (result && result.length) count += result.length;
      } catch (e) {}
    }
    return count;
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
    if (!patterns.length) return 0;

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

          // Skip if not a word break (letter-to-letter across lines)
          if (!lastCh.match(letterRe) || !firstCh.match(letterRe)) continue;

          // Extract end fragment (text without soft hyphens, for pattern check)
          var lineText = String(curLine.contents).replace(/[\u00AD]/g, "");
          var endFrag = "";
          for (var k = lineText.length - 1; k >= 0; k--) {
            var ch = lineText.charAt(k);
            if (ch.match(/[a-zA-Z\u0410-\u044F\u0451\u0401\-]/)) {
              endFrag = ch + endFrag;
            } else { break; }
          }
          endFrag = endFrag.replace(/[\-]+$/, "");

          // Extract start fragment
          var nextText = String(nxtLine.contents).replace(/[\u00AD]/g, "");
          var startFrag = "";
          for (var k = 0; k < nextText.length; k++) {
            var ch = nextText.charAt(k);
            if (ch.match(/[a-zA-Z\u0410-\u044F\u0451\u0401\-]/)) {
              startFrag += ch;
            } else { break; }
          }
          startFrag = startFrag.replace(/^[\-]+/, "");

          // Check patterns: end of current line ENDS WITH pattern,
          // or start of next line STARTS WITH pattern
          var endLower = endFrag.toLowerCase();
          var startLower = startFrag.toLowerCase();
          var found = false;
          for (var pi = 0; pi < patterns.length; pi++) {
            var pat = patterns[pi];
            // endsWith
            if (endLower.length >= pat.length &&
                endLower.lastIndexOf(pat) === endLower.length - pat.length) {
              found = true; break;
            }
            // startsWith
            if (startLower.length >= pat.length &&
                startLower.indexOf(pat) === 0) {
              found = true; break;
            }
          }

          if (found) {
            // Walk backward through actual characters to find word start
            var lc = curLine.characters;
            var wStart = lc.length - 1;
            while (wStart > 0) {
              var c = String(lc[wStart - 1].contents);
              if (c.match(wordCharRe)) { wStart--; } else { break; }
            }

            // Walk forward through next line characters to find word end
            var nc = nxtLine.characters;
            var wEnd = 0;
            while (wEnd < nc.length - 1) {
              var c = String(nc[wEnd + 1].contents);
              if (c.match(wordCharRe)) { wEnd++; } else { break; }
            }

            fixRanges.push({ s: lc[wStart], e: nc[wEnd] });
          }
        } catch (e) {}
      }
    }

    // Apply noBreak — word moves to next line entirely instead of splitting
    for (var i = 0; i < fixRanges.length; i++) {
      try {
        var r = fixRanges[i];
        if (r.s.isValid && r.e.isValid) {
          story.characters.itemByRange(r.s, r.e).noBreak = true;
        }
      } catch (e) {}
    }

    return fixRanges.length;
  }

  // ====== MAIN ======
  var doc = Utils.getActiveDocument();
  if (!doc) { alert("\u041E\u0442\u043A\u0440\u043E\u0439\u0442\u0435 \u0434\u043E\u043A\u0443\u043C\u0435\u043D\u0442."); return; }

  var story = Utils.getTargetStory(doc);
  if (!story) { alert("\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u0442\u0435\u043A\u0441\u0442\u043E\u0432\u044B\u0439 \u0444\u0440\u0435\u0439\u043C."); return; }

  // Sync xlsx -> txt (if xlsx exists and is newer)
  syncRulesFromExcel();

  // Load configs
  var editRules = loadEditorialRules();
  var obscenePatterns = loadObscenePatterns();

  var editCount = 0;
  var spellCount = 0;
  var totalMatches = 0;
  var totalChecked = 0;
  var obsceneCount = 0;
  var matches = [];

  // Step 1: Editorial auto-replacements (separate undo — these are real corrections)
  if (editRules.length) {
    app.doScript(function () {
      editCount = applyEditorialRules(story, editRules);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT,
      "\u0410\u0432\u0442\u043E\u0437\u0430\u043C\u0435\u043D\u044B");
  }

  // Step 2: Spellcheck via Python (if available)
  if (PYTHON_EXE && File(WORKER_PY).exists) {
    var paras = extractParas(story);
    if (paras.length) {
      cleanupTempFiles();
      var inf = File(INPUT_FILE);
      inf.encoding = "UTF-8";
      if (inf.open("w")) {
        inf.write(toJSON(paras));
        inf.close();

        if (runPython()) {
          var result = readResult();
          if (result && !result.error) {
            matches = result.matches || [];
            totalMatches = matches.length;
            totalChecked = result.totalChecked || 0;
          } else if (result && result.error) {
            alert("\u041E\u0448\u0438\u0431\u043A\u0430 \u043F\u0440\u043E\u0432\u0435\u0440\u043A\u0438:\n" + result.error);
          }
        }
        cleanupTempFiles();
      }
    }
  }

  // Step 3: Annotate spelling + check obscene breaks (single undo for all markers)
  if (matches.length || obscenePatterns.length) {
    app.doScript(function () {
      if (matches.length) {
        spellCount = annotate(doc, story, matches);
      }
      if (obscenePatterns.length) {
        obsceneCount = checkObsceneBreaks(doc, story, obscenePatterns);
      }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT,
      "\u041F\u0440\u043E\u0432\u0435\u0440\u043A\u0430 \u0442\u0435\u043A\u0441\u0442\u0430");
  }

  // Summary
  var msg = [];
  if (editCount > 0) msg.push("\u0410\u0432\u0442\u043E\u0437\u0430\u043C\u0435\u043D: " + editCount);
  if (spellCount > 0 || totalMatches > 0) {
    msg.push("\u041E\u0448\u0438\u0431\u043E\u043A: " + spellCount + " \u0438\u0437 " + totalMatches);
    msg.push("  \u041A\u0440\u0430\u0441\u043D\u044B\u043C = \u043E\u0448\u0438\u0431\u043A\u0430, \u0437\u0435\u043B\u0451\u043D\u044B\u043C = \u0438\u0441\u043F\u0440\u0430\u0432\u043B\u0435\u043D\u0438\u0435");
  }
  if (obsceneCount > 0) {
    msg.push("\u0421\u043A\u043B\u0435\u0435\u043D\u043E \u043F\u0435\u0440\u0435\u043D\u043E\u0441\u043E\u0432: " + obsceneCount + " (\u0441\u043B\u043E\u0432\u0430 \u043F\u0435\u0440\u0435\u043D\u0435\u0441\u0435\u043D\u044B \u0446\u0435\u043B\u0438\u043A\u043E\u043C)");
  }
  if (editCount === 0 && totalMatches === 0 && obsceneCount === 0) {
    msg.push("\u041E\u0448\u0438\u0431\u043E\u043A \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u043E.");
    if (totalChecked > 0) msg.push("\u041F\u0440\u043E\u0432\u0435\u0440\u0435\u043D\u043E: " + totalChecked + " \u0430\u0431\u0437.");
  }
  if (!PYTHON_EXE) {
    msg.push("\n\u26A0 Python \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D \u2014 \u043E\u0440\u0444\u043E\u0433\u0440\u0430\u0444\u0438\u044F \u043D\u0435 \u043F\u0440\u043E\u0432\u0435\u0440\u044F\u043B\u0430\u0441\u044C");
  }
  msg.push("\nCtrl+Z \u0434\u043B\u044F \u043E\u0442\u043C\u0435\u043D\u044B");
  alert(msg.join("\n"));
})();
