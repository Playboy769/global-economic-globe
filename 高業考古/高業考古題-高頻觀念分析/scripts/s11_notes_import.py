"""Parse personal study notes (Obsidian 高業備考) into a question bank JSON.

    python s11_notes_import.py

Input : ../筆記練習/{投資學,財務分析(會計),證券交易法規與實務}/*.md
Output: out/notes_corpus.json   (successfully parsed questions)
        out/notes_import_report.txt  (per-file counts + failed-block listing)

Each note file is a numbered list of past-exam-style questions the user wrote
up themselves (question + 4 options + worked explanation + stated answer).
Format is NOT perfectly consistent across files (three different phrasings for
the "answer" line have been observed), so this parser tries several patterns
per block and reports anything it can't confidently extract instead of
guessing -- a wrong answer key silently corrupts the practice bank.
"""
import json, os, re, sys

HERE = os.path.dirname(os.path.abspath(__file__))
SRC_ROOT = os.path.abspath(os.path.join(HERE, '..', '..', '筆記練習'))
OUT = os.path.join(HERE, 'out')

# folder name -> internal subject key used by s10_practice.py / index.html
FOLDER_SUBJECT = {
    '投資學': '投資學',
    '財務分析(會計)': '財務分析',
    '證券交易法規與實務': '法規',
}

FNAME_RE = re.compile(r'^(\d+)\((\d[\d,]*)\)\s*(.+?)(?:\s+(Hard|Medium))?$')
# Most files number questions "1. "; some instead use a flat "- " bullet per
# question (no numbering at all). Both appear in this note set -- picked per
# file by whichever pattern actually finds blocks (see parse_file()).
QBLOCK_NUM_RE = re.compile(r'(?m)^(\d+)\.\s*')
QBLOCK_BULLET_RE = re.compile(r'(?m)^-\s*')
OPT_RE = re.compile(r'\(A\)(.*?)\(B\)(.*?)\(C\)(.*?)\(D\)(.*)$', re.S)

ORDINAL_TO_LETTER = {'一': 'A', '二': 'B', '三': 'C', '四': 'D',
                      '1': 'A', '2': 'B', '3': 'C', '4': 'D'}


def _ordinal_answer(m):
    return ORDINAL_TO_LETTER.get(m.group(1))


ANSWER_PATTERNS = [
    re.compile(r'故正確答案為\s*[：:]?\s*\(([A-D])\)'),
    re.compile(r'答案\s*[：:為]\s*\(([A-D])\)'),
    re.compile(r'正確答案[是為]\s*\(([A-D])\)'),
    re.compile(r'\(([A-D])\)[^\n]{0,4}(?:為|是)?正確答案'),  # reversed: "(B)為正確答案"
]
# explanations that never state "正確答案為(X)" at all, instead giving an
# ordinal ("正確答案為第三個選項") or an item-by-item verdict walk
# ("(C) 五年：...，正確。") -- both lower-confidence than an explicit letter,
# so only consulted if the patterns above found nothing.
FALLBACK_ANSWER_PATTERNS = [
    (re.compile(r'第([一二三四1234])個?選項'), _ordinal_answer),
    (re.compile(r'選項([一二三四1234])'), _ordinal_answer),
    (re.compile(r'\(([A-D])\)[^\n(]{0,60}，?\s*正確[。.)]'), lambda m: m.group(1)),
]


def strip_tabs(text):
    # Obsidian indents continuation lines under a list item with tabs; that
    # indentation carries no meaning for parsing and only gets in the way of
    # the option/answer regexes below.
    return '\n'.join(line.lstrip('\t') for line in text.split('\n'))


def strip_markdown_emphasis(text):
    # "**bold**" and "==highlight==" markers land mid-pattern often enough
    # (e.g. "正確答案為 **(C)**") to break the answer/option regexes below;
    # they carry no meaning for extraction so just drop the marker pairs.
    return text.replace('**', '').replace('==', '')


def extract_answer(block_after_stem):
    seen = set()
    for pat in ANSWER_PATTERNS:
        for m in pat.finditer(block_after_stem):
            seen.add(m.group(1))
    if len(seen) == 1:
        return next(iter(seen))
    if len(seen) > 1:
        return None  # explicit patterns conflict -- don't fall back, just fail
    for pat, get_letter in FALLBACK_ANSWER_PATTERNS:
        for m in pat.finditer(block_after_stem):
            letter = get_letter(m)
            if letter:
                seen.add(letter)
    return next(iter(seen)) if len(seen) == 1 else None


def parse_block(num, raw, subject, fname, period_label):
    text = strip_markdown_emphasis(strip_tabs(raw)).strip()
    # 解析 marks the boundary between "question + options" and "worked
    # explanation" -- explanations for 財務分析/法規 often re-quote
    # "(A) ... (B) ... (C) ... (D) ..." while walking through why each wrong
    # option is wrong, so the real option list must come from BEFORE this
    # marker or that per-option commentary gets mistaken for it.
    split = re.split(r'解析[：:]', text, maxsplit=1)
    head_part = split[0]
    tail_part = split[1] if len(split) > 1 else ''

    head_part = re.sub(r'^題目[：:]\s*', '', head_part.strip())
    opt_match = OPT_RE.search(head_part)
    if not opt_match:
        return None, 'no (A)(B)(C)(D) option run found before 解析'

    stem = head_part[:opt_match.start()].strip()
    stem = re.sub(r'\s+$', '', stem)
    if not stem:
        return None, 'empty stem'

    opts = []
    for letter, seg in zip('ABCD', opt_match.groups()):
        seg = seg.strip()
        # options are separated by a run of spaces (half- or full-width)
        # before the next letter marker was already consumed by OPT_RE, so
        # each seg here is just that option's text with trailing separator
        # whitespace to trim.
        seg = re.sub(r'[　\s]+$', '', seg)
        if not seg:
            return None, f'option ({letter}) is empty'
        opts.append([letter, seg])

    ans = extract_answer(text)  # search whole block incl. explanation
    if not ans:
        return None, 'no unambiguous answer letter found'

    explain = tail_part.strip()
    # trim boilerplate "故正確答案為(X) ..."/"答案為(X)。" sentences that
    # just restate the letter -- the useful reasoning is everything before it.
    explain = re.sub(r'(故)?正確答案[是為].{0,40}$', '', explain).strip()
    explain = re.sub(r'\*\*答案[：:][^\n]*\*\*\s*', '', explain).strip()

    q = {
        'subject': subject,
        'head': stem,
        'opts': opts,
        'ans': ans,
        'explain': explain,
        'srcTxt': f'{period_label} {subject}筆記　第{num}題',
        'srcFile': fname,
    }
    return q, None


def parse_file(path, subject, period_label):
    raw = open(path, encoding='utf-8').read()
    # markdown horizontal rules ("---") also start with "-" and would
    # otherwise be mistaken for an (empty) bulleted question below.
    raw = re.sub(r'(?m)^-{3,}\s*$', '', raw)

    parts = QBLOCK_NUM_RE.split(raw)
    blocks = [(parts[i], parts[i + 1]) for i in range(1, len(parts), 2)]
    if not blocks:
        # fall back to flat "- " bullet questions (no numbering in source)
        pieces = QBLOCK_BULLET_RE.split(raw)[1:]
        blocks = [(str(i + 1), body) for i, body in enumerate(pieces)]

    ok, failed = [], []
    for num, body in blocks:
        q, err = parse_block(num, body, subject, os.path.basename(path), period_label)
        if q:
            ok.append(q)
        else:
            snippet = re.sub(r'\s+', ' ', body.strip())[:80]
            failed.append((num, err, snippet))
    return ok, failed


def period_label_from_filename(fname):
    stem = os.path.splitext(fname)[0]
    m = FNAME_RE.match(stem)
    if not m:
        return stem
    year, sess, _subj, diff = m.groups()
    label = f'{year}年第{sess}次'
    if diff:
        label += f'（{diff}）'
    return label


def main():
    os.makedirs(OUT, exist_ok=True)
    all_ok = []
    report_lines = []
    total_ok = total_fail = 0

    for folder, subject in FOLDER_SUBJECT.items():
        d = os.path.join(SRC_ROOT, folder)
        if not os.path.isdir(d):
            report_lines.append(f'!! missing folder: {d}')
            continue
        report_lines.append(f'== {folder} ({subject}) ==')
        for fname in sorted(os.listdir(d)):
            if not fname.endswith('.md'):
                continue
            path = os.path.join(d, fname)
            period_label = period_label_from_filename(fname)
            ok, failed = parse_file(path, subject, period_label)
            all_ok.extend(ok)
            total_ok += len(ok)
            total_fail += len(failed)
            report_lines.append(f'  {fname}: {len(ok)} ok, {len(failed)} failed')
            for num, err, snippet in failed:
                report_lines.append(f'    [Q{num}] {err} :: {snippet}')

    report_lines.append('')
    report_lines.append(f'TOTAL: {total_ok} ok, {total_fail} failed '
                         f'({total_ok}/{total_ok + total_fail} parsed'
                         + (f', {round(100*total_ok/(total_ok+total_fail))}%)'
                            if (total_ok + total_fail) else ')'))

    by_subject = {}
    for q in all_ok:
        by_subject.setdefault(q['subject'], []).append(q)
    for s, qs in by_subject.items():
        report_lines.append(f'{s}: {len(qs)} 題')

    json.dump(all_ok, open(os.path.join(OUT, 'notes_corpus.json'), 'w', encoding='utf-8'),
               ensure_ascii=False, indent=1)
    report_txt = '\n'.join(report_lines)
    open(os.path.join(OUT, 'notes_import_report.txt'), 'w', encoding='utf-8').write(report_txt)
    print(report_txt)
    print('->', os.path.join(OUT, 'notes_corpus.json'))


if __name__ == '__main__':
    main()
