(function () {
  'use strict';

  // --- Load CommonUtils ---
  var _f = File(File($.fileName).parent.fsName + "/CommonUtils.jsx");
  if (!_f.exists) _f = File(File($.fileName).parent.fsName + "\\CommonUtils.jsx");
  if (!_f.exists) { alert("CommonUtils.jsx not found."); return; }
  try { $.evalFile(_f); } catch (e) { alert("Error loading CommonUtils.jsx:\n" + (e.message || e)); return; }
  var Utils = $.global.CommonUtils;

  // --- Config ---
  var PYTHON_EXE = "C:\\Users\\Administrator\\AppData\\Local\\Programs\\Python\\Python312\\python.exe";
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
      return '"' + obj.replace(/\\/g, "\\\\").replace(/"/g, '\\"')
                      .replace(/\n/g, "\\n").replace(/\r/g, "\\r")
                      .replace(/\t/g, "\\t") + '"';
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
    bat.encoding = "UTF-8";
    if (!bat.open("w")) { alert("\u041D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u0441\u043E\u0437\u0434\u0430\u0442\u044C bat."); return false; }
    bat.writeln("@echo off");
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

  // ====== MAIN ======
  var doc = Utils.getActiveDocument();
  if (!doc) { alert("\u041E\u0442\u043A\u0440\u043E\u0439\u0442\u0435 \u0434\u043E\u043A\u0443\u043C\u0435\u043D\u0442."); return; }

  var story = Utils.getTargetStory(doc);
  if (!story) { alert("\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u0442\u0435\u043A\u0441\u0442\u043E\u0432\u044B\u0439 \u0444\u0440\u0435\u0439\u043C."); return; }

  if (!File(PYTHON_EXE).exists) { alert("Python \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D:\n" + PYTHON_EXE); return; }
  if (!File(WORKER_PY).exists) { alert("Worker \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D:\n" + WORKER_PY); return; }

  // Extract
  var paras = extractParas(story);
  if (!paras.length) { alert("\u041D\u0435\u0442 \u0442\u0435\u043A\u0441\u0442\u0430."); return; }

  // Run
  cleanupTempFiles();
  var inf = File(INPUT_FILE);
  inf.encoding = "UTF-8";
  if (!inf.open("w")) { alert("\u041D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u0437\u0430\u043F\u0438\u0441\u0430\u0442\u044C input."); return; }
  inf.write(toJSON(paras));
  inf.close();

  if (!runPython()) {
    cleanupTempFiles();
    return;
  }

  var result = readResult();
  if (!result) { cleanupTempFiles(); alert("\u041D\u0435\u0442 \u0440\u0435\u0437\u0443\u043B\u044C\u0442\u0430\u0442\u0430."); return; }
  if (result.error) { cleanupTempFiles(); alert("\u041E\u0448\u0438\u0431\u043A\u0430:\n" + result.error); return; }

  var matches = result.matches || [];
  if (!matches.length) {
    alert("\u041E\u0448\u0438\u0431\u043E\u043A \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u043E.\n\u041F\u0440\u043E\u0432\u0435\u0440\u0435\u043D\u043E: " + (result.totalChecked || 0) + " \u0430\u0431\u0437.");
    cleanupTempFiles();
    return;
  }

  // Annotate with single undo
  var count = 0;
  app.doScript(function () {
    count = annotate(doc, story, matches);
  }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT,
    "\u041F\u0440\u043E\u0432\u0435\u0440\u043A\u0430 \u0442\u0435\u043A\u0441\u0442\u0430");

  alert("\u0410\u043D\u043D\u043E\u0442\u0438\u0440\u043E\u0432\u0430\u043D\u043E: " + count + " \u0438\u0437 " + matches.length +
        "\n\n\u041A\u0440\u0430\u0441\u043D\u044B\u043C = \u043E\u0448\u0438\u0431\u043A\u0430, \u0437\u0435\u043B\u0451\u043D\u044B\u043C = \u0438\u0441\u043F\u0440\u0430\u0432\u043B\u0435\u043D\u0438\u0435" +
        "\nCtrl+Z \u0434\u043B\u044F \u043E\u0442\u043C\u0435\u043D\u044B");

  cleanupTempFiles();
})();
