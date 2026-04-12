(function () {
  'use strict';

  var SCRIPT_DIR = File($.fileName).parent.fsName;

  // --- Prompt for word ---
  var word = prompt("\u0412\u0432\u0435\u0434\u0438\u0442\u0435 \u0441\u043B\u043E\u0432\u043E \u0434\u043B\u044F \u0434\u043E\u0431\u0430\u0432\u043B\u0435\u043D\u0438\u044F \u0432 \u0441\u043B\u043E\u0432\u0430\u0440\u044C:", "");
  if (!word) return;
  word = word.replace(/^\s+|\s+$/g, "");
  if (!word) return;

  var dictPath = SCRIPT_DIR + "\\user_dictionary.txt";
  var f = File(dictPath);
  f.encoding = "UTF-8";

  // Check for duplicates
  var existing = "";
  if (f.exists && f.open("r")) {
    existing = f.read();
    f.close();
  }

  var norm = word.toLowerCase().replace(/\s+/g, " ");
  var lines = existing.split(/\r?\n/);
  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].replace(/^\s+|\s+$/g, "");
    if (line.toLowerCase().replace(/\s+/g, " ") === norm) {
      alert("\u0421\u043B\u043E\u0432\u043E \u00AB" + word + "\u00BB \u0443\u0436\u0435 \u0435\u0441\u0442\u044C \u0432 \u0441\u043B\u043E\u0432\u0430\u0440\u0435.");
      return;
    }
  }

  // Append
  if (!f.open("a")) {
    alert("\u041D\u0435 \u0443\u0434\u0430\u043B\u043E\u0441\u044C \u043E\u0442\u043A\u0440\u044B\u0442\u044C \u0444\u0430\u0439\u043B \u0441\u043B\u043E\u0432\u0430\u0440\u044F.");
    return;
  }
  f.writeln(word);
  f.close();

  alert("\u0414\u043E\u0431\u0430\u0432\u043B\u0435\u043D\u043E: \u00AB" + word + "\u00BB\n\n\u041F\u0440\u0438 \u0441\u043B\u0435\u0434\u0443\u044E\u0449\u0435\u0439 \u043F\u0440\u043E\u0432\u0435\u0440\u043A\u0435 \u044D\u0442\u043E \u0441\u043B\u043E\u0432\u043E \u043D\u0435 \u0431\u0443\u0434\u0435\u0442 \u043E\u0442\u043C\u0435\u0447\u0435\u043D\u043E \u043A\u0430\u043A \u043E\u0448\u0438\u0431\u043A\u0430.");
})();
