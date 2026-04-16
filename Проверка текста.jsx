(function () {
  'use strict';
  var dir = File($.fileName).parent.fsName;

  // Batch mode — sub-scripts store results instead of showing alerts
  $.global.__BATCH_MODE = true;
  $.global.__editResult = null;
  $.global.__obsceneResult = null;
  $.global.__spellResult = null;

  var s1 = File(dir + "\\Автозамены.jsx");
  var s2 = File(dir + "\\SpellCheck.jsx");
  var s1Error = null, s2Error = null;

  // Wrap all sub-scripts in a single undo group so Ctrl+Z reverts everything
  try {
    app.doScript(function () {
      if (s1.exists) {
        try { $.evalFile(s1); } catch (e) { s1Error = e.message || String(e); }
      }
      if (s2.exists) {
        try { $.evalFile(s2); } catch (e) { s2Error = e.message || String(e); }
      }
    }, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT,
      "\u041F\u0440\u043E\u0432\u0435\u0440\u043A\u0430 \u0442\u0435\u043A\u0441\u0442\u0430");
  } catch (eBatch) {
    // doScript itself failed — fall back to direct eval (undo may be partial)
    if (!s1Error && s1.exists) {
      try { $.evalFile(s1); } catch (e) { s1Error = e.message || String(e); }
    }
    if (!s2Error && s2.exists) {
      try { $.evalFile(s2); } catch (e) { s2Error = e.message || String(e); }
    }
  } finally {
    $.global.__BATCH_MODE = false;
    // Clean up global result pointers so a subsequent direct run doesn't see stale data
    // (done after report is built — moved to end of function)
  }

  // --- Combined report ---
  var ed = $.global.__editResult || { count: 0, details: [] };
  var ob = $.global.__obsceneResult || { count: 0, words: [] };
  var sp = $.global.__spellResult || { spellCount: 0, totalMatches: 0, totalChecked: 0, noPython: false };

  var msg = [];

  // Auto-replacements
  if (ed.count > 0) {
    msg.push("\u0410\u0432\u0442\u043E\u0437\u0430\u043C\u0435\u043D: " + ed.count);
    var maxShow = Math.min(ed.details.length, 20);
    for (var di = 0; di < maxShow; di++) {
      msg.push(ed.details[di]);
    }
    if (ed.details.length > maxShow) {
      msg.push("  ... \u0438 \u0435\u0449\u0451 " + (ed.details.length - maxShow));
    }
  }

  // Obscene breaks
  if (ob.count > 0) {
    msg.push("\u0421\u043A\u043B\u0435\u0435\u043D\u043E \u043F\u0435\u0440\u0435\u043D\u043E\u0441\u043E\u0432: " + ob.count);
    for (var wi = 0; wi < ob.words.length; wi++) {
      msg.push("  \u2192 " + ob.words[wi]);
    }
  }
  if (ob.hasOverset) {
    msg.push("\u26A0 \u0422\u0435\u043A\u0441\u0442 \u0432\u044B\u0445\u043E\u0434\u0438\u0442 \u0437\u0430 \u0444\u0440\u0435\u0439\u043C \u2014 \u043F\u0440\u043E\u0432\u0435\u0440\u044C\u0442\u0435 \u043F\u0435\u0440\u0435\u043D\u043E\u0441\u044B \u0432\u0440\u0443\u0447\u043D\u0443\u044E.");
  }

  // Spelling
  if (sp.spellCount > 0 || sp.totalMatches > 0) {
    msg.push("\u041E\u0448\u0438\u0431\u043E\u043A: " + sp.spellCount + " \u0438\u0437 " + sp.totalMatches);
    msg.push("  \u041A\u0440\u0430\u0441\u043D\u044B\u043C = \u043E\u0448\u0438\u0431\u043A\u0430, \u0437\u0435\u043B\u0451\u043D\u044B\u043C = \u0438\u0441\u043F\u0440\u0430\u0432\u043B\u0435\u043D\u0438\u0435");
  }

  if (ed.count === 0 && ob.count === 0 && sp.totalMatches === 0) {
    msg.push("\u0418\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u0439 \u043D\u0435\u0442. \u041E\u0448\u0438\u0431\u043E\u043A \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u043E.");
    if (sp.totalChecked > 0) msg.push("\u041F\u0440\u043E\u0432\u0435\u0440\u0435\u043D\u043E: " + sp.totalChecked + " \u0430\u0431\u0437.");
  }

  if (sp.noPython) {
    msg.push("\n\u26A0 Python \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D \u2014 \u043E\u0440\u0444\u043E\u0433\u0440\u0430\u0444\u0438\u044F \u043D\u0435 \u043F\u0440\u043E\u0432\u0435\u0440\u044F\u043B\u0430\u0441\u044C");
  }
  if (sp.error) {
    msg.push("\n\u26A0 \u041E\u0440\u0444\u043E\u0433\u0440\u0430\u0444\u0438\u044F: " + sp.error);
  }
  if (s1Error) {
    msg.push("\n\u26A0 \u0410\u0432\u0442\u043E\u0437\u0430\u043C\u0435\u043D\u044B: " + s1Error);
  }
  if (s2Error) {
    msg.push("\n\u26A0 SpellCheck: " + s2Error);
  }

  msg.push("\nCtrl+Z \u0434\u043B\u044F \u043E\u0442\u043C\u0435\u043D\u044B");
  alert(msg.join("\n"));
})();
