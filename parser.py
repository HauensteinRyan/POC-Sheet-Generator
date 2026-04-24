"""
Parses a UFC POC Word doc into a list of row dicts ready for the xlsx writer.

Each row dict:
  {
    "number": str,   # e.g. "1", "1-1", "10", "10-1"
    "name":   str,   # original capitalisation from doc
    "cue":    str,   # uppercased; phonetic appended at end
  }
"""

import re
from docx import Document

_HEADER_RE = re.compile(
    r"^\s*#?\s*(\d+)\s*[–\-]\s*(.+)$"
)

_VARIANT_RE = re.compile(
    r"^\s*(ALT\s+READ|Prelim\s+read|Main\s+Card\s+read)\s*$",
    re.IGNORECASE,
)

_PHONETIC_RE = re.compile(r"^\s*PHONETIC\s*[-–]", re.IGNORECASE)

# Lines that are clearly stage-direction notes and belong appended to the name,
# not the cue (short, all-caps, appear immediately after a header with no blank).
_NOTE_RE = re.compile(r"^\s*SPOT\s*\(NOT\s+VIZ", re.IGNORECASE)


def _is_header(text: str):
    m = _HEADER_RE.match(text)
    return (m.group(1).strip(), m.group(2).strip()) if m else None


def parse_doc(path: str) -> list[dict]:
    doc = Document(path)
    rows: list[dict] = []

    current_number: str | None = None
    current_name: str | None = None
    cue_lines: list[str] = []
    variant_count: int = 0
    base_slot_taken: bool = False  # True when a variant already occupies the base number
    last_base_number: int = 0      # tracks the most recent integer section number

    blank_run: int = 0             # consecutive blank paragraphs since last non-blank
    just_set_header: bool = False  # True immediately after parsing a header line

    def flush(number, name, lines):
        # Normalize non-breaking spaces; strip leading ": " artifact from lines
        body = []
        phonetic = []
        for l in lines:
            l = l.replace("\xa0", " ")
            l = re.sub(r"^:\s+", "", l)
            if not l:
                continue
            if _PHONETIC_RE.match(l):
                phonetic.append(l)
            else:
                body.append(l)
        cue = " ".join(body).upper()
        if phonetic:
            cue = cue + "\n\n" + " ".join(phonetic).upper()
        rows.append({"number": number, "name": name, "cue": cue})

    def start_section(number_str: str, name_str: str):
        nonlocal current_number, current_name, cue_lines, variant_count
        nonlocal last_base_number, just_set_header, base_slot_taken
        current_number = number_str
        current_name = name_str.strip()
        cue_lines = []
        variant_count = 0
        base_slot_taken = False
        just_set_header = True
        try:
            last_base_number = int(number_str.split("-")[0])
        except ValueError:
            pass

    for para in doc.paragraphs:
        text = para.text.strip()

        if not text:
            blank_run += 1
            just_set_header = False
            continue

        prev_blank_run = blank_run
        blank_run = 0

        # --- Numbered section header ---
        parsed = _is_header(text)
        if parsed:
            if current_number is not None:
                flush(current_number, current_name, cue_lines)
            start_section(parsed[0], parsed[1])
            continue

        # Skip if we haven't seen any header yet
        if current_number is None:
            continue

        # --- Unlabeled header detection ---
        # 5+ consecutive blank lines before a short, period-free line
        # that appears after we've already collected body text for the current section.
        if (
            prev_blank_run >= 5
            and len(cue_lines) > 0
            and len(text) < 80
            and "." not in text
        ):
            flush(current_number, current_name, cue_lines)
            auto_num = str(last_base_number + 1)
            start_section(auto_num, text)
            continue

        # --- Title-extension note (e.g. "SPOT (NOT VIZ!!!)" on the line right after the header) ---
        # Recognized when: no blank since header, no cue yet, matches note pattern.
        if just_set_header and prev_blank_run == 0 and not cue_lines and _NOTE_RE.match(text):
            current_name = current_name + " " + text
            continue

        # --- Variant label ---
        vm = _VARIANT_RE.match(text)
        if vm:
            label = vm.group(1).strip()
            has_base_cue = bool(cue_lines)
            base = current_number.split("-")[0]

            if has_base_cue:
                # Emit the current row, then suffix the next variant.
                # If base_slot_taken, the "1" slot is already used, so next suffix starts at 1.
                flush(current_number, current_name, cue_lines)
                cue_lines = []
                variant_count += 1
                suffix = variant_count - 1 if base_slot_taken else variant_count
                new_num = f"{base}-{suffix}"
            else:
                # No base-section body — first variant takes the bare base number.
                variant_count += 1
                if variant_count == 1:
                    new_num = base
                    base_slot_taken = True
                else:
                    new_num = f"{base}-{variant_count - 1}"

            base_title = re.split(r" - (ALT READ|Prelim|Main)$", current_name)[0]
            if re.search(r"prelim", label, re.IGNORECASE):
                new_name = f"{base_title} - Prelim"
            elif re.search(r"main", label, re.IGNORECASE):
                new_name = f"{base_title} - Main"
            else:
                new_name = f"{base_title} - ALT READ"

            current_number = new_num
            current_name = new_name
            cue_lines = []
            just_set_header = False
            continue

        # --- Phonetic line — append to cue, not its own row ---
        if _PHONETIC_RE.match(text):
            cue_lines.append(text)
            just_set_header = False
            continue

        # --- Regular body text ---
        cue_lines.append(text)
        just_set_header = False

    # Flush final section
    if current_number is not None:
        flush(current_number, current_name, cue_lines)

    return rows


if __name__ == "__main__":
    import sys
    path = sys.argv[1] if len(sys.argv) > 1 else "/Users/ryanh/Downloads/UFC 327 POC for Scripts - V2.docx"
    rows = parse_doc(path)
    for r in rows:
        print(f"{r['number']:>6}  |  {r['name'][:50]:<50}  |  {r['cue'][:80]}")
    print(f"\nTotal rows: {len(rows)}")
