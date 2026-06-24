"""
check_choices.py — Show choice codes for a question (ready-to-paste for datatable.json).

Usage:
  python check_choices.py <path/to/metadata.json> --question Q10
  python check_choices.py <path/to/metadata.json> --all --type SA

Examples:
  python check_choices.py output/VN8932/data/metadata.json --question Q2
  python check_choices.py output/VN8932/data/metadata.json --all --type SA
"""

import json
import sys
import argparse
from pathlib import Path


def _label_text(v):
    if isinstance(v, dict):
        return v.get("vi") or v.get("en") or next(iter(v.values()), "")
    return str(v) if v else ""


def get_choices(q: dict) -> list[dict]:
    """Extract list of {code, label} from a question's choices_i18n."""
    choices_i18n = q.get("choices_i18n") or {}

    # Matrix format: {"rows": {...}, "columns": {...}}
    # Flat format: {"1": {"vi": ..., "en": ...}, "2": {...}, ...}
    if "rows" in choices_i18n or "columns" in choices_i18n:
        # Use columns for matrix (the actual answer choices)
        cols = choices_i18n.get("columns") or {}
        if isinstance(cols, dict):
            return [
                {"code": int(k) if k.isdigit() else k, "label": _label_text(v)}
                for k, v in sorted(cols.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 0)
            ]
        return []

    # Flat: {"1": {"vi": ..., "en": ...}}
    if all(isinstance(v, dict) and ("vi" in v or "en" in v) for v in choices_i18n.values()):
        return [
            {"code": int(k) if k.isdigit() else k, "label": _label_text(v)}
            for k, v in sorted(choices_i18n.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 0)
        ]

    return []


def print_question_choices(q: dict, label: str):
    answer_type = q.get("answer_type") or q.get("type") or ""
    title_raw = q.get("title") or q.get("question_i18n") or {}
    title = _label_text(title_raw)

    print(f"\n{label} ({answer_type}): {title}")
    choices = get_choices(q)
    if not choices:
        print("  (no choices found)")
        return

    print("  Choices:")
    for c in choices:
        print(f"    {c['code']:>4}  {c['label']}")

    print("\n  Ready-to-paste 'groups' for datatable.json:")
    lines = ['  "groups": [']
    for c in choices:
        lines.append(f'    {{ "label": "{c["label"]}", "value": {c["code"]} }},')
    if lines[-1].endswith(","):
        lines[-1] = lines[-1][:-1]  # remove trailing comma on last item
    lines.append("  ]")
    print("\n".join(lines))


def main():
    parser = argparse.ArgumentParser(description="Show choice codes for survey questions")
    parser.add_argument("file", help="Path to metadata.json")
    parser.add_argument("--question", help="Question label (e.g. Q10)")
    parser.add_argument("--all", action="store_true", help="Show all questions")
    parser.add_argument("--type", help="Filter by answer_type when using --all (e.g. SA, MA)")
    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists():
        print(f"ERROR: File not found: {path}", file=sys.stderr)
        sys.exit(1)

    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    questions_raw = data.get("questions", {})
    if isinstance(questions_raw, dict):
        questions = questions_raw  # {label: {...}}
    else:
        questions = {q.get("label", str(i)): q for i, q in enumerate(questions_raw)}

    if args.question:
        q = questions.get(args.question)
        if not q:
            # Try case-insensitive
            q = next((v for k, v in questions.items() if k.lower() == args.question.lower()), None)
        if not q:
            print(f"ERROR: Question '{args.question}' not found in metadata.", file=sys.stderr)
            print(f"Available labels: {', '.join(sorted(questions.keys()))}", file=sys.stderr)
            sys.exit(1)
        print_question_choices(q, args.question)

    elif args.all:
        type_filter = args.type
        for label, q in sorted(questions.items(), key=lambda x: x[1].get("position", 0)):
            answer_type = q.get("answer_type") or q.get("type") or ""
            if type_filter and answer_type.upper() != type_filter.upper():
                continue
            choices = get_choices(q)
            if not choices:
                continue
            print_question_choices(q, label)

    else:
        parser.print_help()


if __name__ == "__main__":
    main()
