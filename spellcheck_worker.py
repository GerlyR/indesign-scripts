# -*- coding: utf-8 -*-
"""
Text proofreading worker for InDesign CS4 ExtendScript.
Uses Yandex Speller (free) for spelling checking of Russian text.

Usage: python spellcheck_worker.py <input.json> <output.json>
"""
import sys
import os
import json
import re

# Force UTF-8 (reconfigure requires Python 3.7+)
try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except AttributeError:
    import io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

try:
    import urllib.request
    import urllib.parse
    import urllib.error
except ImportError:
    print("Python 3 required", file=sys.stderr)
    sys.exit(1)


def check_yandex_speller(text):
    """Check text with Yandex Speller API (free, no key needed)."""
    url = "https://speller.yandex.net/services/spellservice.json/checkText"
    params = urllib.parse.urlencode({
        'text': text,
        'lang': 'ru',
        'options': 14  # IGNORE_URLS(4) | IGNORE_DIGITS(2) | FIND_REPEAT_WORDS(8)
    }).encode('utf-8')

    try:
        req = urllib.request.Request(url, data=params, method='POST')
        req.add_header('Content-Type', 'application/x-www-form-urlencoded')
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            return data  # list of {code, pos, row, col, len, word, s:[suggestions]}
    except Exception as e:
        print(f"Yandex Speller error: {e}", file=sys.stderr)
        return None  # distinguish failure from "no errors"


def normalize_word(word):
    return re.sub(r"\s+", " ", (word or "").strip().lower().replace("ё", "е"))


SURNAME_SUFFIXES = [
    'ов', 'ова', 'ев', 'ева',
    'ин', 'ина', 'ын', 'ына',
    'ович', 'овна', 'евич', 'евна',
    'ский', 'ская', 'цкий', 'цкая',
    'енко', 'чук', 'юк',
    'дзе', 'швили', 'ян', 'янц',
    # 'ец', 'иц' removed — too many false positives (Борец, Кузнец, Певец)
]


def is_likely_name(word):
    """Check if a capitalized word looks like a Russian proper name/surname."""
    if not word or not word[0].isupper():
        return False
    letters = re.sub(r"[^А-Яа-яЁё]", "", word)
    if len(letters) < 3:
        return False
    lower = letters.lower().replace('ё', 'е')
    for suffix in SURNAME_SUFFIXES:
        if lower.endswith(suffix) and len(lower) > len(suffix) + 2:
            return True
    return False


def is_case_only_suggestion(word, suggestions):
    """Skip if all suggestions differ only in capitalization."""
    if not suggestions or not word or not word[0].isupper():
        return False
    wl = word.lower().replace('\u0451', '\u0435')
    return all(s.lower().replace('\u0451', '\u0435') == wl for s in suggestions)


def load_user_dictionary(worker_dir):
    words = set()
    for filename in ["user_dictionary.txt", "known_names.txt"]:
        dict_path = os.path.join(worker_dir, filename)
        try:
            with open(dict_path, "r", encoding="utf-8") as f:
                for raw_line in f:
                    line = raw_line.strip()
                    if not line or line.startswith("#") or line.startswith(";"):
                        continue
                    words.add(normalize_word(line))
        except OSError:
            pass
    return words


def is_likely_sentence_start(text, pos):
    if pos <= 0:
        return True

    i = pos - 1
    while i >= 0 and text[i].isspace():
        i -= 1
    if i < 0:
        return True
    return text[i] in ".!?…:;)\"]»}—–"


def should_skip_uppercase_word(word, global_pos, clean_text, user_words):
    normalized = normalize_word(word)
    if normalized in user_words:
        return True

    letters_only = re.sub(r"[^A-Za-zА-Яа-яЁё]", "", word or "")
    if not letters_only:
        return True

    if letters_only.isupper() and len(letters_only) <= 5:
        return True

    if word and word[0].isupper() and not is_likely_sentence_start(clean_text, global_pos):
        return True

    # Even at sentence start, skip if it looks like a surname
    if word and word[0].isupper() and is_likely_name(word):
        return True

    return False


def is_hyphenation_artifact(word, suggestions):
    if "-" not in word or not suggestions:
        return False
    joined = word.replace("-", "")
    return any(joined.lower() == suggestion.lower() for suggestion in suggestions)


def is_ne_prefix_split(word, suggestions):
    """Skip 'не сознательно'→'несознательно' — Yandex wants to join 'не ' prefix."""
    if not suggestions:
        return False
    wl = word.lower()
    if not wl.startswith("не "):
        return False
    # "не сознательно" → remove space → "несознательно"
    joined = wl[:2] + wl[3:]
    return any(joined == s.lower() for s in suggestions)


def find_paragraph_data(para_data, global_pos):
    for pd in para_data:
        start = pd["clean_offset"]
        end = start + len(pd["clean"])
        if start <= global_pos < end:
            return pd, global_pos - start
    return None, global_pos


def map_clean_span_to_original(paragraph_data, local_clean_offset, length):
    cmap = paragraph_data["map"]
    if local_clean_offset < len(cmap):
        orig_offset = cmap[local_clean_offset]
    else:
        orig_offset = local_clean_offset

    if local_clean_offset + length - 1 < len(cmap):
        orig_end = cmap[local_clean_offset + length - 1]
        orig_length = orig_end - orig_offset + 1
    else:
        orig_length = length

    return paragraph_data["idx"], orig_offset, orig_length


def add_match(all_matches, seen_spans, para_idx, offset, length, word, message, replacements, source):
    filtered_replacements = []
    normalized_word = normalize_word(word)
    for replacement in replacements or []:
        replacement = (replacement or "").strip()
        if not replacement:
            continue
        if normalize_word(replacement) == normalized_word:
            continue
        if replacement not in filtered_replacements:
            filtered_replacements.append(replacement)

    if not filtered_replacements:
        return False

    key = (para_idx, offset, length, tuple(filtered_replacements))
    if key in seen_spans:
        return False

    seen_spans.add(key)
    all_matches.append({
        "para": para_idx,
        "offset": offset,
        "length": length,
        "word": word,
        "message": message,
        "replacements": filtered_replacements,
        "source": source,
    })
    return True


def main():
    if len(sys.argv) < 3:
        print("Usage: python spellcheck_worker.py <input.json> <output.json>",
              file=sys.stderr)
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    try:
        _run(input_path, output_path)
    except Exception as e:
        # Always write output so JSX doesn't hang waiting for the file
        print(f"Fatal error: {e}", file=sys.stderr)
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump({"error": f"Worker crashed: {e}", "matches": [], "totalChecked": 0},
                          f, ensure_ascii=False)
        except Exception as write_err:
            print(f"Cannot write output: {write_err}", file=sys.stderr)
        sys.exit(1)


def _run(input_path, output_path):
    worker_dir = os.path.dirname(os.path.abspath(__file__))
    user_words = load_user_dictionary(worker_dir)

    # Read input
    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            paragraphs = json.load(f)
    except Exception as e:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump({"error": str(e), "matches": [], "totalChecked": 0},
                      f, ensure_ascii=False)
        sys.exit(1)

    # Build per-paragraph data with soft hyphen handling
    SOFT_HYPHEN = '\u00AD'

    para_data = []
    clean_parts = []
    clean_offset = 0

    for para in paragraphs:
        idx = para.get('idx', 0)
        orig = para.get('text', '').replace('\r', '')
        if not orig or not orig.strip():
            continue

        clean_chars = []
        clean_to_orig = []
        for oi, ch in enumerate(orig):
            if ch != SOFT_HYPHEN:
                clean_to_orig.append(oi)
                clean_chars.append(ch)

        clean = ''.join(clean_chars)
        para_data.append({
            'idx': idx,
            'orig': orig,
            'clean': clean,
            'map': clean_to_orig,
            'clean_offset': clean_offset
        })
        clean_parts.append(clean)
        clean_offset += len(clean) + 1

    clean_text = '\n'.join(clean_parts)

    if not clean_text.strip():
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump({"error": None, "matches": [], "totalChecked": 0},
                      f, ensure_ascii=False)
        sys.exit(0)

    all_matches = []
    seen_spans = set()

    # --- Yandex Speller ---
    print("Checking with Yandex Speller...", file=sys.stderr)
    speller_results = check_yandex_speller(clean_text)

    api_error = None
    if speller_results is None:
        api_error = "Yandex Speller API unavailable"
        speller_results = []

    skipped = 0
    for item in speller_results:
        global_pos = item.get('pos', 0)
        word = item.get('word', '')
        length = item.get('len', len(word))
        suggestions = item.get('s', [])

        if not suggestions:
            continue

        if word and word[0].isupper() and is_case_only_suggestion(word, suggestions):
            skipped += 1
            continue

        if should_skip_uppercase_word(word, global_pos, clean_text, user_words):
            skipped += 1
            continue

        if is_hyphenation_artifact(word, suggestions):
            skipped += 1
            continue

        if is_ne_prefix_split(word, suggestions):
            skipped += 1
            continue

        matched_para, local_clean_offset = find_paragraph_data(para_data, global_pos)
        if not matched_para:
            continue

        para_idx, orig_offset, orig_length = map_clean_span_to_original(
            matched_para, local_clean_offset, length
        )
        add_match(
            all_matches,
            seen_spans,
            para_idx,
            orig_offset,
            orig_length,
            word,
            'Орфография',
            suggestions[:3],
            'yandex'
        )

    print(f"  Yandex: {len(speller_results)} found, {skipped} skipped", file=sys.stderr)

    all_matches.sort(key=lambda item: (item['para'], item['offset'], item['length']))

    # Write output
    result = {
        'error': api_error,
        'matches': all_matches,
        'totalChecked': len(para_data)
    }

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Total: {len(all_matches)} issues in {len(para_data)} paragraphs",
          file=sys.stderr)
    sys.exit(0)


if __name__ == '__main__':
    main()
