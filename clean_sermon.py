import email
from email import policy
from html.parser import HTMLParser
import re
import sys
from pathlib import Path

try:
    import docx
except ImportError:
    docx = None

try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None

MAX_LINE_CHARS = 105
DEFAULT_SPEAKER = "Pastor Ernest Fink"
OUTPUT_DIR = Path("~/Desktop").expanduser()

DANGLING_END_WORDS = {
    "a", "an", "the", "of", "to", "in", "on", "at", "by", "for", "with", "from",
    "into", "unto", "upon", "about", "as", "is", "are", "was", "were", "be",
    "has", "have", "had", "will", "shall", "may", "can", "who", "whom", "whose",
    "which", "that", "this", "these", "those", "his", "her", "their", "our", "your",
    "my", "its", "so", "very", "too", "not", "no", "every", "all", "both", "each"
}

CLAUSE_STARTERS = {
    "and", "but", "or", "nor", "for", "yet", "so", "because", "since", "although",
    "though", "while", "where", "when", "whenever", "wherever", "whether", "if",
    "unless", "until", "before", "after", "that", "which", "who", "whom", "whose",
    "what", "whatever", "how", "whereby", "therefore", "thus", "hence", "let",
    "looking", "despising", "ignoring", "bringing", "giving", "removing", "exchanging",
    "trading", "soaring", "running", "walking", "symbolizing", "allowing"
}

BOOK_MAP = {
    "genesis": "Genesis", "gen": "Genesis",
    "exodus": "Exodus", "exod": "Exodus", "exo": "Exodus",
    "leviticus": "Leviticus", "lev": "Leviticus",
    "numbers": "Numbers", "num": "Numbers",
    "deuteronomy": "Deuteronomy", "deut": "Deuteronomy", "deu": "Deuteronomy",
    "joshua": "Joshua", "josh": "Joshua", "jos": "Joshua",
    "judges": "Judges", "judg": "Judges", "jdg": "Judges",
    "ruth": "Ruth", "rut": "Ruth",
    "1 samuel": "1 Samuel", "1 sam": "1 Samuel", "1sam": "1 Samuel", "i samuel": "1 Samuel", "i sam": "1 Samuel",
    "2 samuel": "2 Samuel", "2 sam": "2 Samuel", "2sam": "2 Samuel", "ii samuel": "2 Samuel", "ii sam": "2 Samuel",
    "1 kings": "1 Kings", "1 kgs": "1 Kings", "1kgs": "1 Kings", "i kings": "1 Kings", "i kgs": "1 Kings",
    "2 kings": "2 Kings", "2 kgs": "2 Kings", "2kgs": "2 Kings", "ii kings": "2 Kings", "ii kgs": "2 Kings",
    "1 chronicles": "1 Chronicles", "1 chr": "1 Chronicles", "1 chron": "1 Chronicles", "1chr": "1 Chronicles", "i chronicles": "1 Chronicles", "i chr": "1 Chronicles",
    "2 chronicles": "2 Chronicles", "2 chr": "2 Chronicles", "2 chron": "2 Chronicles", "2chr": "2 Chronicles", "ii chronicles": "2 Chronicles", "ii chr": "2 Chronicles",
    "ezra": "Ezra", "ezr": "Ezra",
    "nehemiah": "Nehemiah", "neh": "Nehemiah",
    "esther": "Esther", "esth": "Esther", "est": "Esther",
    "job": "Job",
    "psalms": "Psalm", "psalm": "Psalm", "ps": "Psalm", "psa": "Psalm",
    "proverbs": "Proverbs", "prov": "Proverbs", "pro": "Proverbs",
    "ecclesiastes": "Ecclesiastes", "eccl": "Ecclesiastes", "ecc": "Ecclesiastes",
    "song of solomon": "Song of Solomon", "song of songs": "Song of Solomon", "song": "Song of Solomon", "sos": "Song of Solomon",
    "isaiah": "Isaiah", "isa": "Isaiah",
    "jeremiah": "Jeremiah", "jer": "Jeremiah",
    "lamentations": "Lamentations", "lam": "Lamentations",
    "ezekiel": "Ezekiel", "ezek": "Ezekiel", "eze": "Ezekiel",
    "daniel": "Daniel", "dan": "Daniel",
    "hosea": "Hosea", "hos": "Hosea",
    "joel": "Joel", "joe": "Joel",
    "amos": "Amos", "amo": "Amos",
    "obadiah": "Obadiah", "obad": "Obadiah", "oba": "Obadiah",
    "jonah": "Jonah", "jon": "Jonah",
    "micah": "Micah", "mic": "Micah",
    "nahum": "Nahum", "nah": "Nahum",
    "habakkuk": "Habakkuk", "hab": "Habakkuk",
    "zephaniah": "Zephaniah", "zeph": "Zephaniah", "zep": "Zephaniah",
    "haggai": "Haggai", "hag": "Haggai",
    "zechariah": "Zechariah", "zech": "Zechariah", "zec": "Zechariah",
    "malachi": "Malachi", "mal": "Malachi",
    "matthew": "Matthew", "matt": "Matthew", "mat": "Matthew",
    "mark": "Mark", "mrk": "Mark", "mk": "Mark",
    "luke": "Luke", "luk": "Luke", "lk": "Luke",
    "john": "John", "joh": "John", "jn": "John",
    "acts": "Acts", "act": "Acts",
    "romans": "Romans", "rom": "Romans",
    "1 corinthians": "1 Corinthians", "1 cor": "1 Corinthians", "1cor": "1 Corinthians", "i corinthians": "1 Corinthians", "i cor": "1 Corinthians",
    "2 corinthians": "2 Corinthians", "2 cor": "2 Corinthians", "2cor": "2 Corinthians", "ii corinthians": "2 Corinthians", "ii cor": "2 Corinthians",
    "galatians": "Galatians", "gal": "Galatians",
    "ephesians": "Ephesians", "eph": "Ephesians",
    "philippians": "Philippians", "philip": "Philippians", "phil": "Philippians", "php": "Philippians",
    "colossians": "Colossians", "col": "Colossians",
    "1 thessalonians": "1 Thessalonians", "1 thess": "1 Thessalonians", "1 the": "1 Thessalonians", "1thess": "1 Thessalonians", "i thessalonians": "1 Thessalonians", "i thess": "1 Thessalonians",
    "2 thessalonians": "2 Thessalonians", "2 thess": "2 Thessalonians", "2 the": "2 Thessalonians", "2thess": "2 Thessalonians", "ii thessalonians": "2 Thessalonians", "ii thess": "2 Thessalonians",
    "1 timothy": "1 Timothy", "1 tim": "1 Timothy", "1tim": "1 Timothy", "i timothy": "1 Timothy", "i tim": "1 Timothy",
    "2 timothy": "2 Timothy", "2 tim": "2 Timothy", "2tim": "2 Timothy", "ii timothy": "2 Timothy", "ii tim": "2 Timothy",
    "titus": "Titus", "tit": "Titus",
    "philemon": "Philemon", "phlm": "Philemon", "phm": "Philemon",
    "hebrews": "Hebrews", "heb": "Hebrews",
    "james": "James", "jas": "James", "jam": "James",
    "1 peter": "1 Peter", "1 pet": "1 Peter", "1pet": "1 Peter", "i peter": "1 Peter", "i pet": "1 Peter",
    "2 peter": "2 Peter", "2 pet": "2 Peter", "2pet": "2 Peter", "ii peter": "2 Peter", "ii pet": "2 Peter",
    "1 john": "1 John", "1 joh": "1 John", "1 jn": "1 John", "1john": "1 John", "i john": "1 John", "i jn": "1 John",
    "2 john": "2 John", "2 joh": "2 John", "2 jn": "2 John", "2john": "2 John", "ii john": "2 John", "ii jn": "2 John",
    "3 john": "3 John", "3 joh": "3 John", "3 jn": "3 John", "3john": "3 John", "iii john": "3 John", "iii jn": "3 John",
    "jude": "Jude", "jud": "Jude",
    "revelation": "Revelation", "revelations": "Revelation", "rev": "Revelation",
}

TRANSLATION_MAP = {
    "amplified bible, classic edition": "AMPC",
    "amplified bible classic edition": "AMPC",
    "amplified bible": "AMP",
    "amplified": "AMP",
    "ampc": "AMPC",
    "amp": "AMP",
    "new living translation": "NLT",
    "nlt2": "NLT",
    "nlt": "NLT",
    "english standard version": "ESV",
    "esv": "ESV",
    "new king james version": "NKJV",
    "nkjv": "NKJV",
    "king james version": "KJV",
    "kjv": "KJV",
    "new international version": "NIV",
    "niv": "NIV",
    "new american standard bible": "NASB",
    "nasb": "NASB",
    "christian standard bible": "CSB",
    "csb": "CSB",
    "god's word translation": "GW",
    "gods word translation": "GW",
    "gw": "GW",
    "the message": "MSG",
    "msg": "MSG",
    "new revised standard version": "NRSV",
    "nrsv": "NRSV",
}

SORTED_BOOK_KEYS = sorted(BOOK_MAP.keys(), key=len, reverse=True)
BOOKS_REGEX = r"(?:" + "|".join(re.escape(k).replace(r"\ ", r"[\s\-]*") for k in SORTED_BOOK_KEYS) + r")"

SORTED_TRANS_KEYS = sorted(TRANSLATION_MAP.keys(), key=len, reverse=True)
TRANSLATIONS = r"(?:" + "|".join(re.escape(k) for k in SORTED_TRANS_KEYS) + r")"

REF_CORE_PATTERN = rf"({BOOKS_REGEX})\.?\s*(\d+(?:\s*:\s*\d+)?(?:\s*-\s*\d+)?)(?:\s*\(?({TRANSLATIONS})\)?)?"


class _HTMLToLines(HTMLParser):
    def __init__(self):
        super().__init__()
        self.chunks = []

    def handle_starttag(self, tag, attrs):
        if tag in ("p", "div", "br", "li", "tr", "h1", "h2", "h3", "h4"):
            self.chunks.append("\n")

    def handle_data(self, data):
        self.chunks.append(data)

    def get_lines(self) -> list[str]:
        text = "".join(self.chunks)
        return [line.strip() for line in text.splitlines() if line.strip()]


def get_canonical_book(book_raw: str) -> str:
    clean_key = re.sub(r"[\s\-]+", " ", book_raw.lower().strip(". "))
    clean_key = re.sub(r"^([1-3])([a-z])", r"\1 \2", clean_key)
    return BOOK_MAP.get(clean_key, book_raw.title())


def get_canonical_trans(trans_raw: str | None) -> str | None:
    if not trans_raw:
        return None
    cleaned = trans_raw.strip("() .").replace("’", "'").lower()
    cleaned = re.sub(r"\s+", " ", cleaned)
    return TRANSLATION_MAP.get(cleaned, trans_raw.upper().strip("() "))


def normalize_book_and_numbers(book_raw: str, numbers_raw: str) -> str:
    canonical_book = get_canonical_book(book_raw)
    clean_nums = re.sub(r"\s*:\s*", ":", numbers_raw.strip())
    clean_nums = re.sub(r"\s*-\s*", "-", clean_nums)
    return f"{canonical_book} {clean_nums}"


def format_ref_prefix(book_raw: str, numbers_raw: str, trans_raw: str | None = None) -> str:
    base_ref = normalize_book_and_numbers(book_raw, numbers_raw)
    norm_trans = get_canonical_trans(trans_raw)
    if norm_trans:
        return f"{base_ref} ({norm_trans})"
    return base_ref


def standardize_refs_in_text(text: str) -> str:
    def _replace(match):
        book, nums, trans = match.group(1), match.group(2), match.group(3)
        return format_ref_prefix(book, nums, trans)

    return re.sub(REF_CORE_PATTERN, _replace, text, flags=re.IGNORECASE)


def add_scripture_ref(scripture_list: list[str], ref_str: str) -> None:
    ref_str = ref_str.strip()
    if not ref_str:
        return
    if scripture_list:
        last = scripture_list[-1]
        if last == ref_str:
            return
        if ref_str.startswith(last + " ("):
            scripture_list[-1] = ref_str
            return
    scripture_list.append(ref_str)


def is_eml_file(filepath: Path) -> bool:
    name_lower = filepath.name.lower()
    if name_lower.endswith(".eml") or name_lower.endswith(".eml.txt"):
        return True
    if filepath.suffix.lower() == ".txt":
        try:
            head = filepath.read_text(encoding="utf-8", errors="ignore")[:1024]
            if re.search(r"^(?:Delivered-To|Return-Path|Received|MIME-Version|Content-Type):", head, flags=re.MULTILINE):
                return True
        except Exception:
            pass
    return False


def unwrap_email_plain_paragraphs(body_text: str) -> list[str]:
    body_text = body_text.replace("\r\n", "\n").replace("\r", "\n")
    paragraphs = re.split(r"\n\s*\n", body_text)
    logical_lines = []

    for para in paragraphs:
        raw_sublines = [s.strip() for s in para.split("\n") if s.strip()]
        if not raw_sublines:
            continue

        current = raw_sublines[0]
        for nxt in raw_sublines[1:]:
            curr_clean = current.replace("*", "").strip()
            nxt_clean = nxt.replace("*", "").strip()

            curr_is_ref = bool(re.match(rf"^\(?{REF_CORE_PATTERN}\)?\Z", curr_clean, flags=re.IGNORECASE))
            nxt_is_ref = bool(re.match(rf"^\(?{REF_CORE_PATTERN}\)?\Z", nxt_clean, flags=re.IGNORECASE))

            if curr_is_ref or nxt_is_ref or re.match(r"^\d{1,3}\s+[A-Z\"“]", nxt_clean):
                logical_lines.append(current)
                current = nxt
            else:
                current = f"{current} {nxt}"

        if current:
            logical_lines.append(current)

    return logical_lines


def clean_email_subject(subject_raw: str) -> str:
    subj = re.sub(r"\s+", " ", subject_raw).strip()
    subj = re.sub(r"^(?:\[[^\]]+\]\s*|(?:\b(?:Re|Fwd|Fw)\s*:\s*))+", "", subj, flags=re.IGNORECASE).strip()

    generic_subjects = {
        "over head outline", "overhead outline", "sermon outline", "sermon notes",
        "outline", "notes", "sunday sermon", "sunday outline", "message outline"
    }
    if subj.lower().strip(" .:-_") in generic_subjects:
        return ""
    return subj


def extract_lines_from_eml(filepath: Path) -> list[str]:
    raw_bytes = filepath.read_bytes()
    msg = email.message_from_bytes(raw_bytes, policy=policy.default)

    extracted = []
    if msg.get("From"):
        extracted.append(f"From: {msg['From']}")
    if msg.get("Subject"):
        cleaned_subj = clean_email_subject(str(msg["Subject"]))
        if cleaned_subj:
            extracted.append(f"Subject: {cleaned_subj}")

    body_part = msg.get_body(preferencelist=("plain", "html"))
    if body_part:
        content = body_part.get_content()
        if body_part.get_content_type() == "text/html":
            parser = _HTMLToLines()
            parser.feed(content)
            extracted.extend(parser.get_lines())
        else:
            extracted.extend(unwrap_email_plain_paragraphs(content))

    return extracted


def merge_orphan_verse_number_lines(lines: list[str]) -> list[str]:
    merged = []
    i = 0
    while i < len(lines):
        curr = lines[i].strip()
        if re.match(r"^\d{1,3}\Z", curr) and (i + 1) < len(lines):
            merged.append(f"{curr} {lines[i + 1].strip()}")
            i += 2
        else:
            merged.append(curr)
            i += 1
    return merged


def extract_lines_from_file(filepath: Path) -> list[str]:
    if is_eml_file(filepath):
        return merge_orphan_verse_number_lines(extract_lines_from_eml(filepath))

    ext = filepath.suffix.lower()
    raw_lines = []
    if ext == ".docx":
        if not docx:
            raise ImportError("Missing dependency. Run: pip install -r requirements.txt")
        doc = docx.Document(filepath)
        for p in doc.paragraphs:
            for sub in p.text.splitlines():
                if sub.strip():
                    raw_lines.append(sub)
    elif ext == ".pdf":
        if not PdfReader:
            raise ImportError("Missing dependency. Run: pip install -r requirements.txt")
        reader = PdfReader(filepath)
        for page in reader.pages:
            for sub in page.extract_text().splitlines():
                if sub.strip():
                    raw_lines.append(sub)
    else:
        for sub in filepath.read_text(encoding="utf-8").splitlines():
            if sub.strip():
                raw_lines.append(sub)

    return merge_orphan_verse_number_lines(raw_lines)


def clean_basic_typography(text: str) -> str:
    text = text.replace("\xa0", " ").replace("\t", " ")
    text = text.replace("“", '"').replace("”", '"').replace("‘", "'").replace("’", "'")
    text = text.replace("*", "")
    text = re.sub(r"(?<!\w)\[[a-zA-Z]\]\s*", "", text)
    text = re.sub(r"(?<=\d)\s*[—–]\s*(?=\d)", "-", text)
    text = re.sub(r"\s*[—–]\s*", " - ", text)
    text = re.sub(r"(?<!\.)([.?!;])([A-Z\[])", r"\1 \2", text)
    text = re.sub(r"(?<!\.)([.?!;])(\d+\s+[A-Z\[\"“])", r"\1 \2", text)
    text = re.sub(r"([a-z\]])([A-Z])", r"\1 \2", text)
    text = re.sub(r"\s{2,}", " ", text)
    return text.strip()


def clean_verse_body(body: str) -> str:
    body = re.sub(r"(?<=[.?!;:,\"\]])\s+(\d{1,3})\s+(?=[A-Z\[\"“])", r" (\1) ", body)
    body = re.sub(r"^\d{1,3}\s+(?=[A-Z\[\"“])", "", body)
    body = body.strip()

    if body.count('"') == 1:
        if body.endswith('"') and not body.startswith('"'):
            body = body[:-1].strip()
        elif not body.endswith('"'):
            body += '"'
    elif body.count('"') % 2 == 1 and not body.endswith('"'):
        body += '"'

    return body


def format_topic_line(text: str) -> str:
    if text.endswith(":"):
        text = text[:-1].strip()

    if text.isupper() and len(text) > 3:
        text = text.title()
    elif ":" in text:
        prefix, rest = text.split(":", 1)
        if prefix.isupper() and len(prefix) > 3:
            text = f"{prefix.title()}:{rest}"
    return text


def tokenize_preserving_brackets(text: str, max_len: int) -> list[str]:
    tokens = []
    i = 0
    n = len(text)
    max_atomic_len = min(45, max(28, int(max_len * 0.52)))

    while i < n:
        while i < n and text[i].isspace():
            i += 1
        if i >= n:
            break

        if text[i] in "([":
            close_char = ")" if text[i] == "(" else "]"
            close_idx = text.find(close_char, i)
            if close_idx != -1:
                end_idx = close_idx + 1
                while end_idx < n and not text[end_idx].isspace():
                    end_idx += 1
                bracket_token = text[i:end_idx]
                if len(bracket_token) <= max_atomic_len:
                    tokens.append(bracket_token)
                    i = end_idx
                    continue

        end_idx = i
        while end_idx < n and not text[end_idx].isspace():
            end_idx += 1
        tokens.append(text[i:end_idx])
        i = end_idx

    return tokens


def score_split_point(left: str, right: str, max_len: int) -> float:
    score = 0.0
    left_len = len(left)
    right_len = len(right)

    fill_ratio = left_len / max_len
    if fill_ratio >= 0.55:
        score += fill_ratio * 25.0
    else:
        score -= (0.55 - fill_ratio) * 60.0

    if right_len <= max_len and right_len < max_len * 0.30:
        score -= 45.0
    elif max_len < right_len < max_len * 1.25:
        score -= 20.0

    if left.endswith((".", "!", "?", '."', '!"', '?"')):
        score += 55.0
    elif left.endswith((";", ":")):
        score += 48.0
    elif left.endswith((",", ',"', "],", "),")):
        score += 38.0
    elif left.endswith(("]", ")")):
        score += 32.0

    last_word_match = re.search(r"([A-Za-z']+)[^A-Za-z']*$", left)
    last_word = last_word_match.group(1).lower() if last_word_match else ""

    first_word_match = re.match(r"^[^A-Za-z']*([A-Za-z']+)", right)
    first_word = first_word_match.group(1).lower() if first_word_match else ""

    if last_word in DANGLING_END_WORDS and not left.endswith((",", ";", ".", "]", ")")):
        score -= 45.0

    if right.startswith("["):
        score += 22.0
    elif right.startswith("(") and fill_ratio >= 0.72:
        score += 12.0
    elif right.startswith("- "):
        score += 35.0
    elif first_word in CLAUSE_STARTERS:
        if first_word in ("and", "or"):
            next_words = right.split()[:3]
            if len(next_words) >= 2 and next_words[1].lower() in (
                "let", "that", "is", "was", "shall", "will", "he", "we", "they", "do", "to", "not", "nothing", "over", "refreshment"
            ):
                score += 28.0
            else:
                score += 10.0
        else:
            score += 24.0

    return score


def smart_wrap_text(text: str, max_len: int) -> list[str]:
    text = text.strip()
    if not text:
        return []
    if len(text) <= max_len:
        return [text]

    tokens = tokenize_preserving_brackets(text, max_len)
    lines = []

    while tokens:
        remaining = " ".join(tokens)
        if len(remaining) <= max_len:
            lines.append(remaining)
            break

        best_idx = 1
        best_score = -9999.0

        for i in range(1, len(tokens)):
            left = " ".join(tokens[:i])
            if len(left) > max_len:
                break
            right = " ".join(tokens[i:])
            sc = score_split_point(left, right, max_len)
            if sc >= best_score:
                best_score = sc
                best_idx = i

        first_candidate = " ".join(tokens[:best_idx])
        if len(first_candidate) > max_len:
            raw_words = remaining.split()
            cur = []
            while raw_words and len(" ".join(cur + [raw_words[0]])) <= max_len:
                cur.append(raw_words.pop(0))
            if not cur and raw_words:
                cur.append(raw_words.pop(0))
            lines.append(" ".join(cur))
            tokens = tokenize_preserving_brackets(" ".join(raw_words), max_len)
            continue

        lines.append(first_candidate.rstrip(" -"))
        tokens = tokens[best_idx:]
        if tokens and tokens[0] == "-":
            tokens.pop(0)

    return lines


def split_plain_line(text: str, max_len: int = MAX_LINE_CHARS) -> list[str]:
    text = text.strip()
    if not text:
        return []

    hebrew_translit_match = re.match(
        r"^([^\x00-\x7F].*?)\.\.\.([a-zA-Z'].*?)\.\.\.(\".*)\Z", text
    )
    if hebrew_translit_match:
        heb_part = hebrew_translit_match.group(1).strip()
        translit_part = f"...{hebrew_translit_match.group(2).strip()}..."
        quote_part = hebrew_translit_match.group(3).strip()
        res = []
        res.extend(smart_wrap_text(heb_part, max_len))
        res.extend(smart_wrap_text(translit_part, max_len))
        res.extend(smart_wrap_text(quote_part, max_len))
        return res

    return smart_wrap_text(text, max_len)


def split_long_verse(ref_prefix: str, body: str, max_total_len: int = MAX_LINE_CHARS) -> list[str]:
    body = clean_verse_body(body)
    full_prefix = f"{ref_prefix} - "
    avail_len = max(45, max_total_len - len(full_prefix))

    verse_num_splits = re.split(r"\s*(?=\(\d{1,3}\)\s)", body)
    final_lines = []

    for segment in verse_num_splits:
        segment = segment.strip().rstrip(";")
        segment = re.sub(r"^\(([1-9]|1\d)\)\s+", "", segment)
        if not segment:
            continue

        wrapped_chunks = smart_wrap_text(segment, avail_len)
        for chunk in wrapped_chunks:
            final_lines.append(f"{full_prefix}{chunk}")

    return final_lines


def merge_with_previous_standalone_ref(
    output_lines: list[str],
    scripture_list: list[str],
    book_raw: str,
    nums_raw: str,
    lead_trans: str | None,
) -> tuple[str, str | None]:
    if not output_lines:
        return nums_raw, lead_trans

    prev_line = output_lines[-1]
    prev_match = re.match(rf"^{REF_CORE_PATTERN}\Z", prev_line, flags=re.IGNORECASE)
    if not prev_match:
        return nums_raw, lead_trans

    prev_book, prev_nums, prev_trans = (
        prev_match.group(1),
        prev_match.group(2),
        prev_match.group(3),
    )

    curr_canonical = get_canonical_book(book_raw)
    prev_canonical = get_canonical_book(prev_book)

    curr_chap = nums_raw.split(":")[0].strip()
    prev_chap = prev_nums.split(":")[0].strip()

    if curr_canonical == prev_canonical and curr_chap == prev_chap and ":" in prev_nums:
        popped_line = output_lines.pop()
        if scripture_list and scripture_list[-1] == popped_line:
            scripture_list.pop()

        clean_prev_nums = re.sub(r"\s*", "", prev_nums)
        clean_curr_nums = re.sub(r"\s*", "", nums_raw)
        best_nums = (
            clean_prev_nums
            if len(clean_prev_nums) >= len(clean_curr_nums)
            else clean_curr_nums
        )
        best_trans = lead_trans or prev_trans
        return best_nums, best_trans

    return nums_raw, lead_trans


def is_verse_continuation(line: str, has_seen_verse: bool = False) -> bool:
    if re.match(r"^\d{1,3}\s+[A-Z\[\"“]", line):
        return True
    if re.match(r"^(?:The\s+Hebrew|The\s+Greek|In\s+[A-Z]|Key\s+|This\s+implies|The\s+scripture|The\s+verse|Know\b)\b", line):
        return False
    if not has_seen_verse and not line.endswith(":"):
        return True
    if len(line) < 95 and not re.match(r"^[A-Za-z\s\"']+:", line):
        return True
    if re.search(r"\b\d{1,2}\s+[A-Z\[]", line):
        return True
    return False


def extract_title_from_filename(filepath: Path) -> str:
    stem = filepath.name
    for ext in (".eml.txt", ".eml", ".docx", ".pdf", ".txt"):
        if stem.lower().endswith(ext):
            stem = stem[: -len(ext)]
            break
    stem = re.sub(r"\s*[-_]\s*OUTLINES?\b.*$", "", stem, flags=re.IGNORECASE)
    stem = re.split(r"\.\d{1,2}\.\d{1,2}\.\d{2,4}\b", stem)[0]
    cleaned = stem.strip(" ._")
    if cleaned.lower() in ("over head outline", "overhead outline", "sermon outline", "outline"):
        return ""
    return cleaned


def clean_email_sender(line: str) -> str | None:
    match = re.match(r"^From\s*:\s*(.+)\Z", line, flags=re.IGNORECASE)
    if not match:
        return None
    val = match.group(1).strip()
    val = re.sub(r"\s*<[^>]+>", "", val)
    val = re.sub(r"\s*\[mailto:[^\]]+\]", "", val)
    val = re.sub(r"\s*\b[\w.\-+]+@[\w.\-]+\.\w+\b", "", val)
    return val.strip(' "\'')


def process_sermon_lines(
    raw_lines: list[str], input_path: Path | None = None
) -> tuple[str, list[str], list[str], str | None]:
    sermon_title = ""
    email_subject_fallback = ""
    extracted_speaker: str | None = None
    output_lines = []
    scripture_list = []

    active_header_ref: str | None = None
    active_header_has_verse = False
    seen_first_content_line = False

    idx = 0
    while idx < len(raw_lines):
        line = clean_basic_typography(raw_lines[idx])
        idx += 1

        if not line:
            continue

        sender_name = clean_email_sender(line)
        if sender_name is not None:
            if sender_name:
                extracted_speaker = sender_name
            continue

        subject_match = re.match(r"^Subject\s*:\s*(.+)\Z", line, flags=re.IGNORECASE)
        if subject_match:
            cleaned_subj = clean_email_subject(subject_match.group(1))
            if cleaned_subj:
                email_subject_fallback = cleaned_subj.title()
            continue

        if re.match(r"^(?:Date|Sent|To|Cc)\s*:\s*", line, flags=re.IGNORECASE):
            continue

        if line.upper() in ("SERMON:", "TITLE:") and idx < len(raw_lines):
            sermon_title = clean_basic_typography(raw_lines[idx]).title()
            seen_first_content_line = True
            idx += 1
            continue
        elif re.match(r"^(?:SERMON|TITLE)\s*:", line, flags=re.IGNORECASE):
            sermon_title = line.split(":", 1)[1].strip().title()
            seen_first_content_line = True
            continue

        if not seen_first_content_line:
            seen_first_content_line = True
            if not sermon_title and len(line) < 65:
                if not re.search(REF_CORE_PATTERN, line, flags=re.IGNORECASE) and not line.endswith(":"):
                    sermon_title = line.title()
                    continue

        if line.upper() in ("TOPIC:", "TOPIC"):
            continue

        standalone_ref_match = re.match(
            rf"^\(?({BOOKS_REGEX})\.?\s*(\d+\s*:\s*\d+(?:\s*-\s*\d+)?)\s*\(?({TRANSLATIONS})?\)?\)?\Z",
            line,
            flags=re.IGNORECASE,
        )
        if standalone_ref_match:
            book_raw = standalone_ref_match.group(1)
            nums_raw = standalone_ref_match.group(2)
            trans_raw = standalone_ref_match.group(3)

            if not trans_raw and idx < len(raw_lines):
                next_clean = clean_basic_typography(raw_lines[idx])
                if re.match(rf"^\(?({TRANSLATIONS})\)?\Z", next_clean, flags=re.IGNORECASE):
                    trans_raw = next_clean
                    idx += 1

            active_header_ref = format_ref_prefix(book_raw, nums_raw, trans_raw)
            active_header_has_verse = False
            add_scripture_ref(scripture_list, active_header_ref)
            continue

        if active_header_ref:
            if is_verse_continuation(line, active_header_has_verse):
                output_lines.extend(split_long_verse(active_header_ref, line))
                active_header_has_verse = True
                continue
            else:
                active_header_ref = None
                active_header_has_verse = False

        inline_paren_refs = re.split(
            rf"(?=\({BOOKS_REGEX}\.?\s*\d+:\s*\d+(?:-\d+)?\s+{TRANSLATIONS}\))",
            line,
            flags=re.IGNORECASE,
        )
        if len(inline_paren_refs) > 1 and any(inline_paren_refs):
            sub_lines = [s.strip() for s in inline_paren_refs if s.strip()]
        else:
            sub_lines = [line]

        for sub in sub_lines:
            topic_colon_ref = re.match(
                rf"^([A-Za-z\s]+?)\s*:\s*{REF_CORE_PATTERN}\s*\Z",
                sub,
                flags=re.IGNORECASE,
            )
            if topic_colon_ref:
                topic_name = format_topic_line(topic_colon_ref.group(1).strip())
                ref_line = format_ref_prefix(
                    topic_colon_ref.group(2),
                    topic_colon_ref.group(3),
                    topic_colon_ref.group(4),
                )
                output_lines.extend(split_plain_line(topic_name))
                output_lines.append(ref_line)
                add_scripture_ref(scripture_list, ref_line)
                continue

            lead_verse_match = re.match(
                rf"^\(?({BOOKS_REGEX})\.?\s*(\d+\s*:\s*\d+(?:\s*-\s*\d+)?)\s*\(?({TRANSLATIONS})?\)?\)?\s*(?:-\s*)?(.*)\Z",
                sub,
                flags=re.IGNORECASE,
            )

            if lead_verse_match and lead_verse_match.group(4):
                book_raw = lead_verse_match.group(1)
                nums_raw = lead_verse_match.group(2)
                lead_trans = lead_verse_match.group(3)
                body = lead_verse_match.group(4).strip()

                if not body.startswith(";"):
                    trail_trans_match = re.search(
                        rf"\s*\(?({TRANSLATIONS})\)?\.?\s*\Z",
                        body,
                        flags=re.IGNORECASE,
                    )
                    if trail_trans_match:
                        if not lead_trans:
                            lead_trans = trail_trans_match.group(1)
                        body = body[: trail_trans_match.start()].strip()

                    nums_raw, lead_trans = merge_with_previous_standalone_ref(
                        output_lines, scripture_list, book_raw, nums_raw, lead_trans
                    )

                    ref_prefix = format_ref_prefix(book_raw, nums_raw, lead_trans)
                    add_scripture_ref(scripture_list, ref_prefix)
                    output_lines.extend(split_long_verse(ref_prefix, body))
                    continue

            formatted_topic = format_topic_line(sub)
            formatted_topic = standardize_refs_in_text(formatted_topic)

            multi_ref_match = re.match(
                rf"^({BOOKS_REGEX}\s+\d+:\d+(?:-\d+)?(?:\s*;\s*{BOOKS_REGEX}\s+\d+:\d+(?:-\d+)?)+)",
                formatted_topic,
                flags=re.IGNORECASE,
            )
            inline_cite = re.search(
                rf"\(({BOOKS_REGEX}\s+\d+:\d+(?:-\d+)?)\)\s+[A-Z\"]",
                formatted_topic,
                flags=re.IGNORECASE,
            )

            output_lines.extend(split_plain_line(formatted_topic))

            if multi_ref_match:
                refs = [r.strip() for r in multi_ref_match.group(1).split(";")]
                output_lines.extend(refs)
                for r in refs:
                    add_scripture_ref(scripture_list, r)
                continue

            if inline_cite:
                output_lines.append(inline_cite.group(1))
                add_scripture_ref(scripture_list, inline_cite.group(1))
                continue

            for m in re.finditer(REF_CORE_PATTERN, formatted_topic, flags=re.IGNORECASE):
                extracted_ref = format_ref_prefix(m.group(1), m.group(2), m.group(3))
                add_scripture_ref(scripture_list, extracted_ref)

    if not sermon_title:
        sermon_title = email_subject_fallback or (extract_title_from_filename(input_path) if input_path else "")

    return sermon_title, output_lines, scripture_list, extracted_speaker


def get_speaker_name(extracted_speaker: str | None) -> str:
    fallback = extracted_speaker or DEFAULT_SPEAKER

    if len(sys.argv) > 2 and sys.argv[2].strip():
        arg_val = sys.argv[2].strip()
        if extracted_speaker and arg_val == DEFAULT_SPEAKER and extracted_speaker.lower() not in DEFAULT_SPEAKER.lower():
            return extracted_speaker
        cleaned_from = clean_email_sender(arg_val)
        return cleaned_from if cleaned_from else arg_val

    if sys.stdin.isatty():
        try:
            entered = input(f"Enter Speaker Name [{fallback}]: ").strip()
            if entered:
                cleaned_from = clean_email_sender(entered)
                return cleaned_from if cleaned_from else entered
        except EOFError:
            pass

    return fallback


def main():
    if len(sys.argv) < 2:
        print("Usage: python clean_sermon.py <input.docx|input.pdf|input.eml|input.txt> [Speaker Name]")
        sys.exit(1)

    input_path = Path(sys.argv[1]).expanduser()

    out_dir = OUTPUT_DIR.expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)

    raw_lines = extract_lines_from_file(input_path)
    sermon_title, cleaned_lines, scripture_list, extracted_speaker = process_sermon_lines(
        raw_lines, input_path
    )

    speaker_name = get_speaker_name(extracted_speaker)

    (out_dir / "MainSpeaker.txt").write_text(speaker_name.strip(), encoding="utf-8")
    (out_dir / "SermonTitle.txt").write_text(sermon_title.strip(), encoding="utf-8")
    (out_dir / "SermonText.txt").write_text("\n".join(cleaned_lines) + "\n", encoding="utf-8")
    (out_dir / "SermonScriptures.txt").write_text("\n".join(scripture_list) + "\n", encoding="utf-8")

    print(
        f"Success! Wrote files to {out_dir}:\n"
        f"  - MainSpeaker.txt ('{speaker_name.strip()}')\n"
        f"  - SermonTitle.txt ('{sermon_title.strip()}')\n"
        f"  - SermonText.txt ({len(cleaned_lines)} lines)\n"
        f"  - SermonScriptures.txt ({len(scripture_list)} scriptures)"
    )


if __name__ == "__main__":
    main()