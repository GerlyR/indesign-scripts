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

  // --- Config ---
  var PYTHON_EXE = Utils.findPython(SCRIPT_DIR);
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

  // --- Create color swatch (find-or-create) ---
  function getColor(doc, name, r, g, b) {
    var existing = Utils.findColorByName(doc, name);
    if (existing) return existing;
    var c = doc.colors.add();
    c.name = name;
    c.model = ColorModel.PROCESS;
    c.space = ColorSpace.RGB;
    c.colorValue = [r, g, b];
    return c;
  }

  // --- Extract paragraphs ---
  // Use .everyItem().contents to fetch all paragraph texts in ONE DOM round-trip
  // instead of N (story.paragraphs.length) separate accesses.
  function extractParas(story) {
    var out = [];
    var texts;
    try { texts = story.paragraphs.everyItem().contents; } catch (e) { texts = null; }
    if (!texts) return out;
    // everyItem().contents returns a string when there's only one paragraph
    if (typeof texts === "string") texts = [texts];
    for (var i = 0; i < texts.length; i++) {
      var t = String(texts[i]).replace(/\r$/, "");
      if (t.replace(/[\s\u00A0]+/g, "")) out.push({ idx: i, text: t });
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
    if (!bat.open("w")) { if (!$.global.__BATCH_MODE) alert("\u041D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u0441\u043E\u0437\u0434\u0430\u0442\u044C bat."); return false; }
    bat.writeln("@echo off");
    bat.writeln("chcp 65001 >nul");
    var cmd = '"' + PYTHON_EXE + '" "' + WORKER_PY + '" "' + INPUT_FILE + '" "' + OUTPUT_FILE + '"';
    bat.writeln(cmd);
    bat.close();

    var launched = File(BAT_FILE).execute();
    if (!launched) {
      if (!$.global.__BATCH_MODE) alert("\u041D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u0437\u0430\u043F\u0443\u0441\u0442\u0438\u0442\u044C \u043F\u0440\u043E\u0432\u0435\u0440\u043A\u0443.");
      return false;
    }

    // Wait for output with exponential backoff (fast initial checks, then 1s)
    var waitedMs = 0;
    var maxWaitMs = 120000; // 2 min
    var delays = [100, 150, 250, 400, 700, 1000]; // ms, then repeat 1000
    var step = 0;
    var out = File(OUTPUT_FILE);
    while (!out.exists && waitedMs < maxWaitMs) {
      var delay = (step < delays.length) ? delays[step] : 1000;
      $.sleep(delay);
      waitedMs += delay;
      step++;
      out = File(OUTPUT_FILE);
    }
    try { File(BAT_FILE).remove(); } catch (e) {}

    if (!out.exists) {
      if (!$.global.__BATCH_MODE) alert("\u0422\u0430\u0439\u043C\u0430\u0443\u0442 (2 \u043C\u0438\u043D).");
      return false;
    }
    $.sleep(200);
    return true;
  }

  // --- Read result ---
  // Returns { parsed: object|null, rawSnippet: string|null, ioError: string|null }
  // Caller can distinguish missing file, I/O failure, and parse failure.
  function readResult() {
    var f = File(OUTPUT_FILE);
    if (!f.exists) return { parsed: null, rawSnippet: null, ioError: "missing" };
    f.encoding = "UTF-8";
    if (!f.open("r")) return { parsed: null, rawSnippet: null, ioError: "locked" };
    var s = "";
    try { s = f.read(); } catch (e) { s = ""; }
    try { f.close(); } catch (e) {}
    if (!s || !s.replace(/\s+/g, "")) {
      return { parsed: null, rawSnippet: "", ioError: "empty" };
    }
    try {
      if (typeof JSON !== "undefined" && JSON.parse) {
        return { parsed: JSON.parse(s), rawSnippet: null, ioError: null };
      }
    } catch (e) {}
    try {
      return { parsed: eval("(" + s + ")"), rawSnippet: null, ioError: null };
    } catch (e2) {
      return { parsed: null, rawSnippet: s.substring(0, 200), ioError: "parse" };
    }
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

        // Style arrow "→" gray, no underline
        try {
          para.characters[tagStart].fillColor = gray;
          para.characters[tagStart].underline = false;
        } catch (e) {}

        // Style suggestion green, no underline
        if (tagLen > 1) {
          try {
            var s0 = para.characters[tagStart + 1];
            var s1 = para.characters[tagStart + tagLen - 1];
            var sr = para.characters.itemByRange(s0, s1);
            sr.fillColor = green;
            sr.underline = false;
          } catch (e) {}
        }

        n++;
      } catch (e) {}
    }
    return n;
  }


  // ====== MAIN ======
  var doc = Utils.getActiveDocument();
  if (!doc) { $.global.__spellResult = { spellCount: 0, totalMatches: 0, totalChecked: 0, noPython: false, error: null }; if (!$.global.__BATCH_MODE) alert("\u041E\u0442\u043A\u0440\u043E\u0439\u0442\u0435 \u0434\u043E\u043A\u0443\u043C\u0435\u043D\u0442."); return; }

  var story = Utils.getTargetStory(doc);
  if (!story) { $.global.__spellResult = { spellCount: 0, totalMatches: 0, totalChecked: 0, noPython: false, error: null }; if (!$.global.__BATCH_MODE) alert("\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u0442\u0435\u043A\u0441\u0442\u043E\u0432\u044B\u0439 \u0444\u0440\u0435\u0439\u043C."); return; }

  if (!PYTHON_EXE && !$.global.__BATCH_MODE) {
    // Offer to manually specify Python path
    var userPath = prompt("Python \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D \u0430\u0432\u0442\u043E\u043C\u0430\u0442\u0438\u0447\u0435\u0441\u043A\u0438.\n\u0423\u043A\u0430\u0436\u0438\u0442\u0435 \u043F\u0443\u0442\u044C \u043A python.exe\n(\u0438\u043B\u0438 \u043E\u0442\u043C\u0435\u043D\u0438\u0442\u0435 \u0434\u043B\u044F \u043F\u0440\u043E\u043F\u0443\u0441\u043A\u0430 \u043E\u0440\u0444\u043E\u0433\u0440\u0430\u0444\u0438\u0438):", "C:\\Python311\\python.exe");
    if (userPath) {
      userPath = userPath.replace(/[\r\n]/g, "").replace(/^\s+|\s+$/g, "");
      var upFile = File(userPath);
      // Validate: file exists AND basename looks like a python executable
      // This prevents accidental paths to notepad.exe etc. being stored.
      var looksLikePython = /[\\\/]python\d*w?\.exe$/i.test(userPath);
      if (!upFile.exists) {
        alert("\u0424\u0430\u0439\u043B \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D: " + userPath);
      } else if (!looksLikePython) {
        alert("\u042D\u0442\u043E \u043D\u0435 \u043F\u043E\u0445\u043E\u0436\u0435 \u043D\u0430 python.exe: " + userPath + "\n\u041F\u0443\u0442\u044C \u0434\u043E\u043B\u0436\u0435\u043D \u0437\u0430\u043A\u0430\u043D\u0447\u0438\u0432\u0430\u0442\u044C\u0441\u044F \u043D\u0430 python.exe \u0438\u043B\u0438 python3.exe.");
      } else {
        PYTHON_EXE = userPath;
        Utils.savePythonPath(SCRIPT_DIR, userPath);
      }
    }
  }
  if (!PYTHON_EXE) {
    $.global.__spellResult = { spellCount: 0, totalMatches: 0, totalChecked: 0, noPython: true };
    if (!$.global.__BATCH_MODE) alert("\u26A0 Python \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D \u2014 \u043E\u0440\u0444\u043E\u0433\u0440\u0430\u0444\u0438\u044F \u043D\u0435 \u043C\u043E\u0436\u0435\u0442 \u0431\u044B\u0442\u044C \u043F\u0440\u043E\u0432\u0435\u0440\u0435\u043D\u0430.");
    return;
  }

  var spellCount = 0;
  var totalMatches = 0;
  var totalChecked = 0;
  var spellError = null;
  var matches = [];

  // Step 1: Spellcheck via Python
  var paras = extractParas(story);
  if (paras.length) {
    cleanupTempFiles();
    var inf = File(INPUT_FILE);
    inf.encoding = "UTF-8";
    if (inf.open("w")) {
      inf.write(toJSON(paras));
      inf.close();

      if (runPython()) {
        var rr = readResult();
        if (rr.parsed && !rr.parsed.error) {
          matches = rr.parsed.matches || [];
          totalMatches = matches.length;
          totalChecked = rr.parsed.totalChecked || 0;
        } else if (rr.parsed && rr.parsed.error) {
          spellError = rr.parsed.error;
          if (!$.global.__BATCH_MODE) alert("\u041E\u0448\u0438\u0431\u043A\u0430 \u043F\u0440\u043E\u0432\u0435\u0440\u043A\u0438:\n" + rr.parsed.error);
        } else if (rr.ioError === "parse") {
          // Worker output was truncated or corrupted — don't silently report "no errors"
          spellError = "\u041D\u0435\u0432\u0435\u0440\u043D\u044B\u0439 JSON \u043E\u0442 worker: " + (rr.rawSnippet || "").substring(0, 120);
          if (!$.global.__BATCH_MODE) alert(spellError);
        } else if (rr.ioError === "empty" || rr.ioError === "missing") {
          spellError = "\u041F\u0443\u0441\u0442\u043E\u0439 \u043E\u0442\u0432\u0435\u0442 \u043E\u0442 \u043F\u0440\u043E\u0432\u0435\u0440\u043A\u0438";
        }
      }
      cleanupTempFiles();
    }
  }

  // Step 2: Annotate spelling errors
  if (matches.length) {
    app.doScript(function () {
      spellCount = annotate(doc, story, matches);
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT,
      "\u041F\u0440\u043E\u0432\u0435\u0440\u043A\u0430 \u043E\u0440\u0444\u043E\u0433\u0440\u0430\u0444\u0438\u0438");
  }

  // --- Store results for batch mode or show report ---
  $.global.__spellResult = { spellCount: spellCount, totalMatches: totalMatches, totalChecked: totalChecked, noPython: !PYTHON_EXE, error: spellError };

  if ($.global.__BATCH_MODE) return;

  var msg = [];
  if (spellCount > 0 || totalMatches > 0) {
    msg.push("\u041E\u0448\u0438\u0431\u043E\u043A: " + spellCount + " \u0438\u0437 " + totalMatches);
    msg.push("  \u041A\u0440\u0430\u0441\u043D\u044B\u043C = \u043E\u0448\u0438\u0431\u043A\u0430, \u0437\u0435\u043B\u0451\u043D\u044B\u043C = \u0438\u0441\u043F\u0440\u0430\u0432\u043B\u0435\u043D\u0438\u0435");
  } else {
    msg.push("\u041E\u0448\u0438\u0431\u043E\u043A \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u043E.");
    if (totalChecked > 0) msg.push("\u041F\u0440\u043E\u0432\u0435\u0440\u0435\u043D\u043E: " + totalChecked + " \u0430\u0431\u0437.");
  }
  msg.push("\nCtrl+Z \u0434\u043B\u044F \u043E\u0442\u043C\u0435\u043D\u044B");
  alert(msg.join("\n"));
})();
