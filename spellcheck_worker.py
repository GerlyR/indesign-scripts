# -*- coding: utf-8 -*-
"""
Text proofreading worker v4 for InDesign CS4 ExtendScript.
Uses Yandex Speller (free) + Claude API (optional) for comprehensive
spelling, grammar, punctuation and style checking of Russian text.

Usage: python spellcheck_worker.py <input.json> <output.json> [claude_api_key]
"""
import sys
import os
import json
import re

# Force UTF-8
sys.stdout.reconfigure(encoding='utf-8', errors='replace')
sys.stderr.reconfigure(encoding='utf-8', errors='replace')

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
        'options': 518  # IGNORE_URLS | IGNORE_DIGITS | FIND_REPEAT_WORDS | IGNORE_CAPITALIZATION
    }).encode('utf-8')

    try:
        req = urllib.request.Request(url, data=params, method='POST')
        req.add_header('Content-Type', 'application/x-www-form-urlencoded')
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            return data  # list of {code, pos, row, col, len, word, s:[suggestions]}
    except Exception as e:
        print(f"Yandex Speller error: {e}", file=sys.stderr)
        return []


def check_claude_api(text, api_key):
    """Check text with Claude API for grammar/style/punctuation."""
    if not api_key:
        return []

    url = "https://api.anthropic.com/v1/messages"

    prompt = """Ты — профессиональный корректор русского текста для газеты.
Проверь текст и найди ТОЛЬКО реальные ошибки:
- Орфографические ошибки (но НЕ имена собственные, НЕ названия клубов, НЕ фамилии)
- Грамматические ошибки (согласование, падежи, числа)
- Пунктуационные ошибки (пропущенные/лишние запятые, тире, двоеточия)
- Стилистические ошибки (тавтология, канцеляризмы)

НЕ трогай:
- Имена, фамилии, прозвища
- Названия команд, городов, стадионов
- Спортивную терминологию
- Цитаты в кавычках
- Аббревиатуры

Верни JSON-массив найденных ошибок. Каждая ошибка:
{"word": "ошибочное слово/фраза", "suggestion": "исправление", "type": "spelling|grammar|punctuation|style", "context": "...фрагмент текста с ошибкой..."}

Если ошибок нет — верни пустой массив [].
Отвечай ТОЛЬКО JSON-массивом, без пояснений."""

    body = json.dumps({
        "model": "claude-sonnet-4-20250514",
        "max_tokens": 4096,
        "messages": [
            {"role": "user", "content": prompt + "\n\nТекст:\n" + text}
        ]
    })

    try:
        req = urllib.request.Request(url, data=body.encode('utf-8'), method='POST')
        req.add_header('Content-Type', 'application/json')
        req.add_header('x-api-key', api_key)
        req.add_header('anthropic-version', '2023-06-01')

        with urllib.request.urlopen(req, timeout=60) as resp:
            result = json.loads(resp.read().decode('utf-8'))

        # Extract text from response
        content = result.get('content', [])
        if not content:
            return []

        response_text = content[0].get('text', '').strip()

        # Parse JSON from response (handle markdown code blocks)
        if response_text.startswith('```'):
            # Remove markdown code block
            lines = response_text.split('\n')
            response_text = '\n'.join(lines[1:-1])

        return json.loads(response_text)

    except Exception as e:
        print(f"Claude API error: {e}", file=sys.stderr)
        return []


def normalize_word(word):
    return re.sub(r"\s+", " ", (word or "").strip().lower().replace("ё", "е"))


def load_claude_key(worker_dir, cli_key):
    if cli_key:
        return cli_key.strip()

    key_path = os.path.join(worker_dir, "claude_api_key.txt")
    try:
        with open(key_path, "r", encoding="utf-8") as f:
            return re.sub(r"[\r\n\s]+", "", f.read())
    except OSError:
        return ""


def load_user_dictionary(worker_dir):
    words = set()
    dict_path = os.path.join(worker_dir, "user_dictionary.txt")
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
    return text[i] in ".!?…:;)\"]»}"


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

    return False


def is_hyphenation_artifact(word, suggestions):
    if "-" not in word or not suggestions:
        return False
    joined = word.replace("-", "")
    return any(joined.lower() == suggestion.lower() for suggestion in suggestions)


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


def iter_occurrences(text, needle):
    start = 0
    step = max(1, len(needle))
    while True:
        pos = text.find(needle, start)
        if pos < 0:
            return
        yield pos
        start = pos + step


def normalize_context(text):
    if not text:
        return ""
    text = text.lower().replace("ё", "е")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def pick_best_occurrence(clean_text, word, context, para_data, used_locations):
    positions = list(iter_occurrences(clean_text, word))
    if not positions:
        return None

    normalized_context = normalize_context(context)
    context_tokens = set(re.findall(r"\w+", normalized_context))
    best = None
    best_score = None

    for pos in positions:
        paragraph_data, local_clean_offset = find_paragraph_data(para_data, pos)
        if not paragraph_data:
            continue

        para_idx, offset, length = map_clean_span_to_original(paragraph_data, local_clean_offset, len(word))
        location_key = (para_idx, offset, length)

        window = clean_text[max(0, pos - 50): min(len(clean_text), pos + len(word) + 50)]
        normalized_window = normalize_context(window)
        window_tokens = set(re.findall(r"\w+", normalized_window))

        score = 0
        if normalized_context:
            if normalized_context in normalized_window or normalized_window in normalized_context:
                score += 1000
            score += len(context_tokens & window_tokens) * 10
        if location_key in used_locations:
            score -= 500
        score -= pos / 100000.0

        if best is None or score > best_score:
            best = {
                "para_idx": para_idx,
                "offset": offset,
                "length": length,
                "location_key": location_key,
            }
            best_score = score

    return best


def main():
    if len(sys.argv) < 3:
        print("Usage: python spellcheck_worker.py <input.json> <output.json> [api_key]",
              file=sys.stderr)
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]
    worker_dir = os.path.dirname(os.path.abspath(__file__))
    claude_key = load_claude_key(worker_dir, sys.argv[3] if len(sys.argv) > 3 else "")
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
    # Original text (with \u00AD) is used for offset mapping back to InDesign
    # Clean text (without \u00AD) is sent to spellchecker
    SOFT_HYPHEN = '\u00AD'

    para_data = []  # list of {idx, orig_text, clean_text, clean_to_orig_map}
    clean_parts = []
    clean_offset = 0

    for para in paragraphs:
        idx = para.get('idx', 0)
        orig = para.get('text', '').replace('\r', '')
        if not orig or not orig.strip():
            continue

        # Build clean text and offset mapping
        clean_chars = []
        clean_to_orig = []  # clean_pos → orig_pos
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
            'clean_offset': clean_offset  # global offset in joined clean text
        })
        clean_parts.append(clean)
        clean_offset += len(clean) + 1  # +1 for \n

    clean_text = '\n'.join(clean_parts)

    if not clean_text.strip():
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump({"error": None, "matches": [], "totalChecked": 0},
                      f, ensure_ascii=False)
        sys.exit(0)

    all_matches = []
    seen_spans = set()

    # --- 1. Yandex Speller ---
    print("Checking with Yandex Speller...", file=sys.stderr)
    speller_results = check_yandex_speller(clean_text)

    skipped = 0
    for item in speller_results:
        global_pos = item.get('pos', 0)
        word = item.get('word', '')
        length = item.get('len', len(word))
        suggestions = item.get('s', [])

        if not suggestions:
            continue

        if should_skip_uppercase_word(word, global_pos, clean_text, user_words):
            skipped += 1
            continue

        if is_hyphenation_artifact(word, suggestions):
            skipped += 1
            continue

        if normalize_word(word) in user_words:
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

    # --- 2. Claude API (if key provided) ---
    if claude_key:
        print("Checking with Claude API...", file=sys.stderr)
        claude_results = check_claude_api(clean_text, claude_key)

        for item in claude_results:
            if not isinstance(item, dict):
                continue
            word = item.get('word', '')
            suggestion = item.get('suggestion', '')
            err_type = item.get('type', 'grammar')
            context = item.get('context', '')

            if not word or not suggestion:
                continue

            if normalize_word(word) in user_words:
                continue

            best = pick_best_occurrence(clean_text, word, context, para_data, {
                (match['para'], match['offset'], match['length']) for match in all_matches
            })
            if not best:
                continue

            type_labels = {
                'spelling': 'Орфография',
                'grammar': 'Грамматика',
                'punctuation': 'Пунктуация',
                'style': 'Стилистика'
            }

            add_match(
                all_matches,
                seen_spans,
                best['para_idx'],
                best['offset'],
                best['length'],
                word,
                type_labels.get(err_type, err_type),
                [suggestion],
                'claude'
            )

        print(f"  Claude: {len(claude_results)} issues", file=sys.stderr)
    else:
        print("  Claude API: skipped (no key)", file=sys.stderr)

    all_matches.sort(key=lambda item: (item['para'], item['offset'], item['length']))

    # Write output
    result = {
        'error': None,
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
