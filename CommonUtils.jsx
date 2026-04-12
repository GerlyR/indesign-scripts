/**
 * Общая библиотека утилит для скриптов InDesign
 * Содержит переиспользуемые функции для работы со стилями, текстом и документами
 * 
 * Использование: $.evalFile(File("путь/к/CommonUtils.jsx"));
 */

(function() {
  'use strict';

  if (typeof $.global === 'undefined') $.global = {};

  // Если уже загружен и инициализирован — выходим
  if ($.global.CommonUtils && typeof $.global.CommonUtils.getActiveDocument === 'function') {
    return;
  }

  var Utils = {};
  $.global.CommonUtils = Utils;

  try {
  
  // ============================================================================
  // Работа с документами и историями
  // ============================================================================
  
  /**
   * Получает активный документ с проверкой
   */
  Utils.getActiveDocument = function() {
    if (!app.documents || app.documents.length === 0) {
      return null;
    }
    return app.activeDocument || null;
  };
  
  /**
   * Получает целевую историю из выделения или документа
   */
  Utils.getTargetStory = function(doc) {
    if (!doc) doc = Utils.getActiveDocument();
    if (!doc) return null;
    
    if (app.selection && app.selection.length > 0) {
      try {
        var s0 = app.selection[0];
        if (s0 && s0.hasOwnProperty("texts") && s0.texts && s0.texts.length > 0) {
          return s0.texts[0].parentStory;
        }
        if (s0 && s0.constructor) {
          var cname = s0.constructor.name;
          if (cname === "InsertionPoint" || cname === "TextFrame") {
            return s0.parentStory;
          }
        }
      } catch (e) {}
    }
    
    if (doc.stories && doc.stories.length > 0) {
      try {
        return doc.stories[0];
      } catch (e) {}
    }
    
    return null;
  };
  
  /**
   * Получает все истории из выделения
   */
  Utils.getSelectedStories = function() {
    if (!app.selection || app.selection.length === 0) return [];
    
    var sel = app.selection;
    var out = [];
    var seen = {};
    
    for (var i = 0; i < sel.length; i++) {
      try {
        var it = sel[i];
        var st = null;
        
        if (it && it.parentStory) {
          st = it.parentStory;
        } else if (it && it.constructor && it.constructor.name === "TextFrame" && it.parentStory) {
          st = it.parentStory;
        }
        
        if (st && st.isValid && !seen[st.id]) {
          out.push(st);
          seen[st.id] = 1;
        }
      } catch (e) {}
    }
    
    return out;
  };
  
  // ============================================================================
  // Работа со стилями
  // ============================================================================
  
  /**
   * Получает или создает стиль абзаца
   */
  Utils.ensureParaStyle = function(doc, name) {
    if (!doc || !name) return null;
    try {
      var ps = doc.paragraphStyles.itemByName(name);
      if (ps && ps.isValid) return ps;
    } catch (e) {}
    try {
      return doc.paragraphStyles.add({name: name});
    } catch (e) {
      return null;
    }
  };
  
  /**
   * Получает стиль абзаца (без создания)
   */
  Utils.getParaStyle = function(doc, name) {
    if (!doc || !name) return null;
    try {
      var ps = doc.paragraphStyles.itemByName(name);
      if (ps && ps.isValid) return ps;
    } catch (e) {}
    return null;
  };
  
  /**
   * Получает стиль абзаца из группы
   */
  Utils.getParaStyleFromGroup = function(doc, styleName, groupName) {
    if (!doc || !styleName) return null;
    
    function walk(g) {
      if (!g) return null;
      try {
        for (var i = 0; i < g.paragraphStyles.length; i++) {
          var ps = g.paragraphStyles[i];
          if (ps && ps.name === styleName) return ps;
        }
        for (var j = 0; j < g.paragraphStyleGroups.length; j++) {
          var r = walk(g.paragraphStyleGroups[j]);
          if (r) return r;
        }
      } catch (e) {}
      return null;
    }
    
    var grp = null;
    if (groupName) {
      try {
        for (var i = 0; i < doc.paragraphStyleGroups.length; i++) {
          var g = doc.paragraphStyleGroups[i];
          if (g && g.name === groupName) {
            grp = g;
            break;
          }
        }
      } catch (e) {}
    }
    
    if (grp) {
      try {
        var s = grp.paragraphStyles.itemByName(styleName);
        if (s && s.isValid) return s;
      } catch (e) {}
      var f = walk(grp);
      if (f) return f;
    }
    
    try {
      var s2 = doc.paragraphStyles.itemByName(styleName);
      if (s2 && s2.isValid) return s2;
    } catch (e) {}
    
    for (var i = 0; i < doc.paragraphStyleGroups.length; i++) {
      var h = walk(doc.paragraphStyleGroups[i]);
      if (h) return h;
    }
    
    return null;
  };
  
  /**
   * Получает или создает символьный стиль
   */
  Utils.ensureCharStyle = function(doc, name, fontStyleName) {
    if (!doc || !name) return null;
    try {
      var cs = doc.characterStyles.itemByName(name);
      if (cs && cs.isValid) return cs;
    } catch (e) {}
    try {
      cs = doc.characterStyles.add({name: name});
      if (fontStyleName) {
        try {
          cs.fontStyle = fontStyleName;
        } catch (e) {}
      }
      return cs;
    } catch (e) {
      return null;
    }
  };
  
  /**
   * Получает символьный стиль (без создания)
   */
  Utils.getCharStyle = function(doc, name) {
    if (!doc || !name) return null;
    try {
      var cs = doc.characterStyles.itemByName(name);
      if (cs && cs.isValid) return cs;
    } catch (e) {}
    return null;
  };
  
  /**
   * Применяет жирный стиль к тексту
   */
  Utils.applyBold = function(textObj, charStyle) {
    if (!textObj) return false;
    try { if (textObj.isValid === false) return false; } catch (e) {}
    
    if (charStyle) {
      try {
        textObj.appliedCharacterStyle = charStyle;
        return true;
      } catch (e) {}
    }
    
    var candidates = ["Bold", "SemiBold", "Semibold", "Demi", "Black", "Heavy", 
                      "Жирный", "Полужирный", "Medium", "DemiBold"];
    for (var i = 0; i < candidates.length; i++) {
      try {
        textObj.fontStyle = candidates[i];
        return true;
      } catch (e) {}
    }
    return false;
  };
  
  /**
   * Применяет курсив к тексту
   */
  Utils.applyItalic = function(textObj, charStyle) {
    if (!textObj) return false;
    try { if (textObj.isValid === false) return false; } catch (e) {}
    
    if (charStyle) {
      try {
        textObj.appliedCharacterStyle = charStyle;
        return true;
      } catch (e) {}
    }
    
    var candidates = ["Italic", "Oblique", "Курсив", "Наклонный"];
    for (var i = 0; i < candidates.length; i++) {
      try {
        textObj.fontStyle = candidates[i];
        return true;
      } catch (e) {}
    }
    return false;
  };
  
  /**
   * Применяет жирный курсив к тексту
   */
  Utils.applyBoldItalic = function(textObj, charStyle) {
    if (!textObj) return false;
    try { if (textObj.isValid === false) return false; } catch (e) {}
    
    if (charStyle) {
      try {
        textObj.appliedCharacterStyle = charStyle;
        return true;
      } catch (e) {}
    }
    
    var combos = [
      "Bold Italic", "BoldItalic", "Bold Oblique", "BoldOblique",
      "Black Italic", "BlackItalic", "Heavy Italic", "HeavyItalic",
      "Semibold Italic", "SemiBold Italic", "Demi Italic",
      "Жирный курсив", "Полужирный курсив", "Полужирный наклонный"
    ];
    for (var i = 0; i < combos.length; i++) {
      try {
        textObj.fontStyle = combos[i];
        return true;
      } catch (e) {}
    }
    
    // Фолбэк: fontStyle напрямую (не charStyle, чтобы не перезаписывать)
    var boldCandidates = ["Bold", "SemiBold", "Semibold", "Demi", "Black", "Heavy", "Жирный", "Полужирный"];
    var italicCandidates = ["Italic", "Oblique", "Курсив", "Наклонный"];
    for (var bi = 0; bi < boldCandidates.length; bi++) {
      for (var ii = 0; ii < italicCandidates.length; ii++) {
        try { textObj.fontStyle = boldCandidates[bi] + " " + italicCandidates[ii]; return true; } catch (e) {}
      }
    }
    return false;
  };
  
  // ============================================================================
  // Работа с текстом
  // ============================================================================
  
  /**
   * Обрезка пробелов с обеих сторон
   */
  Utils.trim = function(s) {
    if (s == null) return "";
    s = String(s);
    return s.replace(/^[\s\u00A0]+/, "").replace(/[\s\u00A0]+$/, "");
  };
  
  /**
   * Получает текст абзаца без завершающего \r
   */
  Utils.getParaText = function(p) {
    if (!p || !p.contents) return "";
    return String(p.contents).replace(/\r$/, "");
  };
  
  /**
   * Экранирование для регулярных выражений
   */
  Utils.escapeRegex = function(s) {
    if (s == null) return "";
    return String(s).replace(/([\\\/\.\^\$\|\?\*\+\(\)\[\]\{\}])/g, "\\$1");
  };
  
  /**
   * Замена неразрывных пробелов на обычные
   */
  Utils.normalizeSpaces = function(s) {
    if (s == null) return "";
    return String(s).replace(/[\u00A0\u202F\u2009\u200A]/g, " ");
  };
  
  // ============================================================================
  // Find/Change операции
  // ============================================================================
  
  /**
   * Сброс настроек Find/Change
   */
  Utils.resetFindGrep = function() {
    try {
      app.findGrepPreferences = NothingEnum.nothing;
      app.changeGrepPreferences = NothingEnum.nothing;
    } catch (e) {}
  };
  
  /**
   * Сброс настроек Find/Change для текста
   */
  Utils.resetFindText = function() {
    try {
      app.findTextPreferences = NothingEnum.nothing;
      app.changeTextPreferences = NothingEnum.nothing;
    } catch (e) {}
  };
  
  /**
   * Выполняет Grep замену в истории
   */
  Utils.grepChange = function(story, findWhat, changeProps) {
    if (!story || !story.isValid || !findWhat) return null;
    
    Utils.resetFindGrep();
    try {
      app.findGrepPreferences.findWhat = findWhat;
      if (changeProps) {
        for (var k in changeProps) {
          if (changeProps.hasOwnProperty(k)) {
            try {
              app.changeGrepPreferences[k] = changeProps[k];
            } catch (e) {}
          }
        }
      }
      return story.changeGrep();
    } catch (e) {
      return null;
    } finally {
      Utils.resetFindGrep();
    }
  };
  
  /**
   * Выполняет текстовую замену
   */
  Utils.textChange = function(story, findWhat, changeTo) {
    if (!story || !story.isValid || !findWhat) return null;
    
    Utils.resetFindText();
    try {
      app.findTextPreferences.findWhat = findWhat;
      app.changeTextPreferences.changeTo = changeTo;
      return story.changeText();
    } catch (e) {
      return null;
    } finally {
      Utils.resetFindText();
    }
  };
  
  // ============================================================================
  // Очистка текста
  // ============================================================================
  
  /**
   * Удаляет пустые абзацы в конце истории
   */
  Utils.trimTailEmptyParas = function(story) {
    if (!story || !story.paragraphs) return;
    
    while (story.paragraphs.length > 0) {
      try {
        var p = story.paragraphs[story.paragraphs.length - 1];
        if (!p || !p.isValid) break;
        
        var t = Utils.trim(Utils.getParaText(p));
        if (t.length === 0) {
          try {
            p.remove();
          } catch (e) {
            break;
          }
        } else {
          break;
        }
      } catch (e) {
        break;
      }
    }
  };
  
  /**
   * Удаляет абзацы, содержащие только пробелы
   */
  Utils.removeWhitespaceParas = function(story) {
    if (!story || !story.paragraphs) return;
    try {
      for (var i = story.paragraphs.length - 1; i >= 0; i--) {
        var p = story.paragraphs[i];
        if (p && p.isValid && /^[ \t\u00A0]*$/.test(Utils.getParaText(p))) {
          try {
            p.remove();
          } catch (e) {}
        }
      }
    } catch (e) {}
  };
  

  /**
   * Схлопывает множественные переносы строк
   */
  Utils.collapseBreaks = function(story) {
    if (!story) return;
    var changed = 1;
    var iterations = 0;
    var maxIterations = 10;
    
    while (changed > 0 && iterations < maxIterations) {
      changed = 0;
      iterations++;
      
      Utils.resetFindGrep();
      try {
        app.findGrepPreferences.findWhat = "\\r[ \\x{00A0}\\t]*\\r+";
        app.changeGrepPreferences.changeTo = "\\r";
        var result = story.changeGrep();
        changed += (result && result.length) ? result.length : 0;
      } catch (e) {}
      
      Utils.resetFindGrep();
      try {
        app.findGrepPreferences.findWhat = "\\r\\r+";
        app.changeGrepPreferences.changeTo = "\\r";
        var result2 = story.changeGrep();
        changed += (result2 && result2.length) ? result2.length : 0;
      } catch (e) {}
      
      Utils.removeWhitespaceParas(story);
    }
  };
  
  /**
   * Универсальная функция очистки текста (cleanupBroom)
   * Оптимизированная версия с единым списком правил
   */
  Utils.cleanupBroom = function(story) {
    if (!story || !story.isValid) return;
    
    // Правила очистки: [findWhat, changeTo]
    var rules = [
      ["\\.\\.(?!\\.)", "."],
      ["\\.\\s+\\.(?!\\.)", "."],
      ["\\(\\(([^)]*)\\)\\)", "($1)"],
      ["\\(\\(", "("],
      ["\\)\\)", ")"],
      ["\\[\\[", "["],
      ["\\]\\]", "]"],
      ["\\{\\{", "{"],
      ["\\}\\}", "}"],
      ["\u00AB\u00AB", "\u00AB"],
      ["\u00BB\u00BB", "\u00BB"]
    ];
    
    // Применяем правила
    for (var i = 0; i < rules.length; i++) {
      Utils.resetFindGrep();
      try {
        app.findGrepPreferences.findWhat = rules[i][0];
        app.changeGrepPreferences.changeTo = rules[i][1];
        story.changeGrep();
      } catch (e) {}
    }
    
    Utils.resetFindGrep();
  };
  
  /**
   * Применяет стандартные замены (ТВ-ГИД, даты)
   */
  Utils.applyStandardReplacements = function(story) {
    if (!story || !story.isValid) return;
    
    Utils.resetFindGrep();
    try {
      app.findGrepPreferences.findWhat = "ТВ-ГИД\\.|ТВ-гид\\.";
      app.changeGrepPreferences.changeTo = "ТВ-ГИД";
      story.changeGrep();
    } catch (e) {}
    
    Utils.resetFindGrep();
    try {
      app.findGrepPreferences.findWhat = "(Понедельник|Вторник|Среда|Четверг|Пятница|Суббота|Воскресенье)\\s*,\\s*(\\d{1,2})\\s+(января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\\s*\\.";
      app.changeGrepPreferences.changeTo = "$1, $2 $3";
      story.changeGrep();
    } catch (e) {}
    
    Utils.resetFindGrep();
  };
  
  // ============================================================================
  // Работа с файлами
  // ============================================================================

  /**
   * Ищет установленный Python (3.8–3.13) на Windows.
   * Сначала проверяет python_path.txt в scriptDir, затем стандартные пути.
   * @param {string} scriptDir — fsName папки, содержащей скрипт
   * @returns {string|null} — fsName python.exe или null
   */
  Utils.findPython = function(scriptDir) {
    var cf = File(scriptDir + "\\python_path.txt");
    if (cf.exists) {
      cf.encoding = "UTF-8";
      if (cf.open("r")) {
        var p = cf.read().replace(/[\r\n\s]+/g, "");
        cf.close();
        if (p && File(p).exists) return p;
      }
    }
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
  };

  /**
   * Получает папку Scripts Panel
   */
  Utils.getScriptsPanelFolder = function() {
    var curFile = null;
    try {
      curFile = File(app.activeScript);
    } catch (e) {
      try {
        curFile = File($.fileName);
      } catch (e2) {
        return null;
      }
    }
    
    if (!curFile || !curFile.parent) return null;
    
    var f = curFile.parent;
    while (f && f.exists) {
      try {
        var name = f.displayName || f.name || "";
        if (/Scripts Panel$/i.test(name)) return f;
        f = f.parent;
      } catch (e) {
        break;
      }
    }
    return null;
  };
  
  /**
   * Загружает список команд-исключений из INI файла
   */
  Utils.loadTeamExceptions = function() {
    var list = [];
    try {
      var panel = Utils.getScriptsPanelFolder();
      if (!panel) return list;
      
      var ini = File(panel.fullName + "/Команды-исключения.ini");
      if (ini && ini.exists && ini.open("r")) {
        try {
          while (!ini.eof) {
            var line = Utils.trim(ini.readln());
            if (line && !/^#/.test(line) && !/^;/.test(line)) {
              list.push(line);
            }
          }
        } finally {
          ini.close();
        }
      }
    } catch (e) {}
    return list;
  };
  
  /**
   * Строит регулярное выражение для команд из списка
   */
  Utils.buildTeamRegex = function(teamList, capture) {
    if (!teamList || teamList.length === 0) {
      teamList = ["ПСЖ", "ЦСКА"];
    }
    
    var arr = [];
    for (var i = 0; i < teamList.length; i++) {
      var s = Utils.escapeRegex(teamList[i]);
      if (s) arr.push(s);
    }
    
    if (!arr.length) {
      arr = ["ПСЖ", "ЦСКА"];
    }
    
    return (capture ? "(" : "(?:") + arr.join("|") + ")";
  };
  
  // ============================================================================
  // Вспомогательные функции
  // ============================================================================
  
  /**
   * Извлекает скобки в конце строки
   */
  Utils.extractEndingBrackets = function(s) {
    if (!s) return {base: "", brk: ""};
    s = String(s);
    var m = s.match(/\(([^\(\)]*)\)[\s.!?…]*$/);
    if (!m) return {base: s, brk: ""};
    
    var brk = "(" + m[1] + ")";
    var lastIdx = s.lastIndexOf(brk);
    var base = (lastIdx >= 0) ? s.slice(0, lastIdx) : s;
    return {base: base, brk: brk};
  };
  
  /**
   * Проверяет, является ли текст подзаголовком.
   * Подзаголовок: короткий текст (до maxLen символов), не начинается с тире,
   * не заканчивается завершающей пунктуацией предложения.
   * maxLen по умолчанию 200.
   */
  Utils.isSubheader = function(txt, maxLen) {
    var t = Utils.trim(txt);
    if (!t) return false;
    if (typeof maxLen !== 'number') maxLen = 200;
    if (t.length > maxLen) return false;
    if (/^[\s\u00A0]*[\-\u2013\u2014]/.test(t)) return false;
    if (/^\(.*\)\s*\.?\s*$/.test(t)) return false;
    return !/[.!?\u2026,:;)"\u00BB\u2019\u203A\]]\s*$/.test(t);
  };

  /**
   * Проверяет, начинается ли текст с тире
   */
  Utils.beginsWithDash = function(s) {
    if (!s) return false;
    return /^[\s\u00A0]*[-–—][\s\u00A0]/.test(String(s));
  };

  /**
   * Проверяет, является ли текст подписью автора (Имя Фамилия).
   * Оба слова должны начинаться с заглавной кириллической буквы.
   * Допускается точка/пунктуация в конце.
   */
  Utils.isSignature = function(txt) {
    var t = Utils.trim(txt);
    if (!t) return false;
    t = t.replace(/\s*\([^)]*\)\s*[.!?\u2026]*\s*$/, "");
    t = t.replace(/[.!?\u2026,]+$/, "");
    t = Utils.trim(t);
    if (!t) return false;
    // Два слова: Имя ФАМИЛИЯ, Имя Фамилия, или ИМЯ ФАМИЛИЯ
    return /^[А-ЯЁ][а-яёА-ЯЁ]+\s+[А-ЯЁ][а-яёА-ЯЁ]+$/.test(t);
  };

  /**
   * Проверяет, является ли строка продолжением подписи ("из Москвы.", "из Калининграда.")
   */
  Utils.isSignatureCity = function(txt) {
    var t = Utils.trim(txt);
    if (!t) return false;
    return /^из\s+[А-ЯЁа-яё][а-яё\-]+\s*[.!?\u2026]*\s*$/.test(t);
  };

  /**
   * Делает первую букву заглавной
   */
  Utils.capitalizeFirst = function(word) {
    if (!word) return word;
    return word.charAt(0).toUpperCase() + word.slice(1).toLowerCase();
  };

  // ============================================================================
  // Автодетекция шрифта и умные символьные стили
  // ============================================================================

  /**
   * Определяет варианты шрифта (Bold/Italic/BoldItalic) из story.
   * Сканирует app.fonts, подбирает варианты той же ширины (Condensed/Regular),
   * предпочитает "Bold" над "Semibold", короткие имена.
   * @param {Story} story — текстовая история для анализа
   * @returns {{ bold:string|null, italic:string|null, boldItalic:string|null, family:string|null, baseStyle:string|null }}
   */
  Utils.detectFontVariants = function(story) {
    var result = { bold: null, italic: null, boldItalic: null, family: null, baseStyle: null };
    try {
      var samplePara = null;
      for (var pi = 0; pi < story.paragraphs.length && pi < 5; pi++) {
        try {
          var pp = story.paragraphs[pi];
          if (pp && pp.isValid && pp.characters.length > 2) {
            samplePara = pp;
            if (pi >= 2) break;
          }
        } catch (e) {}
      }
      if (!samplePara) return result;

      var baseFont = samplePara.characters[0].appliedFont;
      result.family = baseFont.fontFamily;
      result.baseStyle = baseFont.fontStyleName || "Regular";

      var baseWidth = "";
      var widthMatch = result.baseStyle.match(/(Condensed|Cond|Narrow|Compressed|Extended|Wide)/i);
      if (widthMatch) baseWidth = widthMatch[1].toLowerCase();

      var candidates = { bold: [], italic: [], boldItalic: [] };
      var allFonts = app.fonts;
      for (var fi = 0; fi < allFonts.length; fi++) {
        try {
          var f = allFonts[fi];
          if (f.fontFamily !== result.family) continue;
          var sn = f.fontStyleName;
          var isBold = /\b(bold|black|heavy|demi|semibold)\b/i.test(sn) || /полужирн|жирн/i.test(sn);
          var isItalic = /\b(italic|oblique)\b/i.test(sn) || /курсив|наклон/i.test(sn);
          if (isBold && isItalic) candidates.boldItalic.push(sn);
          else if (isBold && !isItalic) candidates.bold.push(sn);
          else if (isItalic && !isBold) candidates.italic.push(sn);
        } catch (e) {}
      }

      function scoreName(name) {
        var s = 0;
        var hasWidth = /(condensed|cond|narrow|compressed|extended|wide)/i.test(name);
        var nameWidth = "";
        var nw = name.match(/(condensed|cond|narrow|compressed|extended|wide)/i);
        if (nw) nameWidth = nw[1].toLowerCase();
        if (baseWidth && nameWidth === baseWidth) s += 100;
        else if (!baseWidth && !hasWidth) s += 100;
        else if (!baseWidth && hasWidth) s -= 50;
        else if (baseWidth && !hasWidth) s -= 50;
        if (/^bold\b/i.test(name)) s += 10;
        if (/\bitalic\b/i.test(name)) s += 5;
        s -= name.length;
        return s;
      }

      function pickBest(arr) {
        if (!arr || arr.length === 0) return null;
        var best = arr[0], bestScore = scoreName(arr[0]);
        for (var i = 1; i < arr.length; i++) {
          var sc = scoreName(arr[i]);
          if (sc > bestScore) { best = arr[i]; bestScore = sc; }
        }
        return best;
      }

      result.bold = pickBest(candidates.bold);
      result.italic = pickBest(candidates.italic);
      result.boldItalic = pickBest(candidates.boldItalic);
    } catch (e) {}
    return result;
  };

  /**
   * Создаёт символьный стиль под конкретный fontStyle, не перезаписывая общий стиль по имени.
   * Это защищает уже оформленный текст от побочных изменений при повторном запуске скриптов.
   * @param {Document} doc
   * @param {string} name — базовое имя стиля
   * @param {string|null} detectedFontStyle — fontStyle из detectFontVariants
   * @returns {CharacterStyle|null}
   */
  Utils.ensureCharStyleSmart = function(doc, name, detectedFontStyle) {
    if (!doc || !name) return null;

    function getStyle(styleName) {
      if (!styleName) return null;
      try {
        var cs = doc.characterStyles.itemByName(styleName);
        if (cs && cs.isValid) return cs;
      } catch (e) {}
      return null;
    }

    function makeVariantName(baseName, fontStyle) {
      if (!fontStyle) return baseName;
      var token = String(fontStyle)
        .replace(/[^\w\u0400-\u04FF]+/g, "_")
        .replace(/^_+|_+$/g, "");
      if (!token) return baseName;
      return baseName + "__" + token;
    }

    var baseStyle = getStyle(name);
    if (!detectedFontStyle) {
      if (baseStyle) return baseStyle;
      try {
        return doc.characterStyles.add({name: name});
      } catch (e) {
        return null;
      }
    }

    try {
      if (baseStyle && baseStyle.fontStyle === detectedFontStyle) {
        return baseStyle;
      }
    } catch (e) {}

    var variantName = makeVariantName(name, detectedFontStyle);
    var variantStyle = getStyle(variantName);
    if (variantStyle) {
      try { variantStyle.fontStyle = detectedFontStyle; } catch (e) {}
      return variantStyle;
    }

    try {
      var cs2 = doc.characterStyles.add({name: variantName});
      try { cs2.fontStyle = detectedFontStyle; } catch (e) {}
      return cs2;
    } catch (e) {}

    if (baseStyle) return baseStyle;
    return null;
  };

  /**
   * Применяет символьный стиль ко всему абзацу через itemByRange (не everyItem!).
   * Фолбэк: если стиль не сработал — пробует fontStyle напрямую,
   * затем перебор Bold Italic комбинаций.
   * @param {Paragraph} para
   * @param {CharacterStyle|null} charStyle
   * @param {string|null} fontStyleName — фолбэк fontStyle
   */
  Utils.applyCharStyleToPara = function(para, charStyle, fontStyleName) {
    if (!para || !para.isValid || para.characters.length === 0) return;
    try {
      var rng = para.characters.itemByRange(0, para.characters.length - 1);
      if (!rng || !rng.isValid) return;

      if (charStyle) {
        try { rng.appliedCharacterStyle = charStyle; return; } catch (e) {}
      }
      if (fontStyleName) {
        try { rng.fontStyle = fontStyleName; return; } catch (e) {}
      }
      // Последний фолбэк — перебор комбинаций Bold Italic
      var combos = ["Bold Italic", "BoldItalic", "Bold Oblique", "Полужирный курсив"];
      for (var c = 0; c < combos.length; c++) {
        try { rng.fontStyle = combos[c]; return; } catch (e) {}
      }
    } catch (e) {}
  };

  /**
   * Применяет символьный стиль к диапазону символов абзаца.
   * @param {Paragraph} para
   * @param {number} startIdx — индекс начального символа
   * @param {number} endIdx — индекс конечного символа (включительно)
   * @param {CharacterStyle|null} charStyle
   * @param {string|null} fontStyleName — фолбэк fontStyle
   */
  Utils.applyCharStyleToRange = function(para, startIdx, endIdx, charStyle, fontStyleName) {
    if (!para || !para.isValid || para.characters.length === 0) return;
    if (startIdx < 0) startIdx = 0;
    if (endIdx >= para.characters.length) endIdx = para.characters.length - 1;
    if (startIdx > endIdx) return;
    try {
      var rng = para.characters.itemByRange(startIdx, endIdx);
      if (!rng || !rng.isValid) return;
      if (charStyle) {
        try { rng.appliedCharacterStyle = charStyle; return; } catch (e) {}
      }
      if (fontStyleName) {
        try { rng.fontStyle = fontStyleName; return; } catch (e) {}
      }
    } catch (e) {}
  };

  /**
   * Применяет курсив к эмоциональным ремаркам в скобках: (смеётся), (улыбаясь) и т.п.
   * @param {Story} story
   * @param {CharacterStyle|null} italicStyle
   */
  Utils.applyEmotionRemarks = function(story, italicStyle) {
    if (!story || !story.isValid || !italicStyle) return;
    Utils.resetFindGrep();
    try {
      app.findGrepPreferences.findWhat = "\\([а-яё]+\\)";
      app.changeGrepPreferences.appliedCharacterStyle = italicStyle;
      app.changeGrepPreferences.changeTo = "$0";
      story.changeGrep();
    } catch (e) {}
    Utils.resetFindGrep();
  };

  } catch (e) {
    $.global.CommonUtils = Utils || {};
  }

})();

