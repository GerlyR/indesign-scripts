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

  // ====== MAIN ======
  var doc = Utils.getActiveDocument();
  if (!doc) { alert("\u041E\u0442\u043A\u0440\u043E\u0439\u0442\u0435 \u0434\u043E\u043A\u0443\u043C\u0435\u043D\u0442."); return; }

  var story = Utils.getTargetStory(doc);
  if (!story) { alert("\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u0442\u0435\u043A\u0441\u0442\u043E\u0432\u044B\u0439 \u0444\u0440\u0435\u0439\u043C."); return; }

  var removed = 0;

  app.doScript(function () {
    // Step 1: Remove inserted suggestion arrows and their text
    // SpellCheck inserts "\u2192suggestion" after each error word
    Utils.resetFindGrep();
    try {
      app.findGrepPreferences.findWhat = "\u2192[^\u2192\\r]+";
      app.changeGrepPreferences.changeTo = "";
      var r1 = story.changeGrep();
      if (r1 && r1.length) removed += r1.length;
    } catch (e) {}
    Utils.resetFindGrep();

    // Step 2: Reset red-colored text (error marks)
    var red = null;
    for (var i = 0; i < doc.colors.length; i++) {
      try {
        if (doc.colors[i].name === "_SC_Red") { red = doc.colors[i]; break; }
      } catch (e) {}
    }
    if (red) {
      Utils.resetFindGrep();
      try {
        app.findGrepPreferences.findWhat = ".+";
        app.findGrepPreferences.fillColor = red;
        var black = doc.swatches.itemByName("Black");
        app.changeGrepPreferences.fillColor = black;
        app.changeGrepPreferences.underline = false;
        app.changeGrepPreferences.changeTo = "$0";
        var r2 = story.changeGrep();
        if (r2 && r2.length) removed += r2.length;
      } catch (e) {}
      Utils.resetFindGrep();
    }

    // Step 3: Reset green-colored text (suggestions that weren't removed)
    var green = null;
    for (var i = 0; i < doc.colors.length; i++) {
      try {
        if (doc.colors[i].name === "_SC_Green") { green = doc.colors[i]; break; }
      } catch (e) {}
    }
    if (green) {
      Utils.resetFindGrep();
      try {
        app.findGrepPreferences.findWhat = ".+";
        app.findGrepPreferences.fillColor = green;
        var black = doc.swatches.itemByName("Black");
        app.changeGrepPreferences.fillColor = black;
        app.changeGrepPreferences.changeTo = "$0";
        story.changeGrep();
      } catch (e) {}
      Utils.resetFindGrep();
    }

    // Step 4: Reset gray-colored text (arrows)
    var gray = null;
    for (var i = 0; i < doc.colors.length; i++) {
      try {
        if (doc.colors[i].name === "_SC_Gray") { gray = doc.colors[i]; break; }
      } catch (e) {}
    }
    if (gray) {
      Utils.resetFindGrep();
      try {
        app.findGrepPreferences.findWhat = ".+";
        app.findGrepPreferences.fillColor = gray;
        var black = doc.swatches.itemByName("Black");
        app.changeGrepPreferences.fillColor = black;
        app.changeGrepPreferences.changeTo = "$0";
        story.changeGrep();
      } catch (e) {}
      Utils.resetFindGrep();
    }
  }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT,
    "\u041E\u0447\u0438\u0441\u0442\u043A\u0430 \u043E\u0440\u0444\u043E\u0433\u0440\u0430\u0444\u0438\u0438");

  if (removed > 0) {
    alert("\u041C\u0430\u0440\u043A\u0435\u0440\u044B \u0443\u0434\u0430\u043B\u0435\u043D\u044B: " + removed + "\n\nCtrl+Z \u0434\u043B\u044F \u043E\u0442\u043C\u0435\u043D\u044B");
  } else {
    alert("\u041C\u0430\u0440\u043A\u0435\u0440\u043E\u0432 \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u043E.");
  }
})();
