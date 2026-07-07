"""
list_questions.py — List survey questions by type.

Usage:
  python list_questions.py <path/to/definition.json or metadata.json> [--type stub|banner]

Examples:
  python list_questions.py output/VN8932/mcp/definition.json --type stub
  python list_questions.py output/VN8932/data/metadata.json  --type banner
"""

import json
import sys
import argparse
from pathlib import Path

STUB_TYPES = {
    "singlechoice", "multiplechoice",
    "matrix_singlechoice", "matrix_multiplechoice", "matrix_numeric",
    "singlenumber", "multiplenumber",
    "SA", "MA", "Matrix_SA", "Matrix_MA", "Matrix_NUM",
    "NUM", "multiplenumber",
}
BANNER_TYPES = {
    "singlechoice", "multiplechoice",
    "SA", "MA",
}


def _label_text(v):
    """Extract display text from i18n dict or plain string."""
    if isinstance(v, dict):
        return v.get("vi") or v.get("en") or next(iter(v.values()), "")
    return str(v) if v else ""


def list_from_definition(data: dict, mode: str):
    """Parse QMe definition.json format."""
    questions = data.get("questions", [])
    if isinstance(questions, dict):
        questions = list(questions.values())

    results = []
    for q in questions:
        type_name = (q.get("type_name") or "").lower()
        label = q.get("label") or q.get("id") or ""
        title_raw = q.get("question_i18n") or q.get("title") or {}
        title = _label_text(title_raw)
        position = q.get("position", 0)

        # Check filter
        allowed = STUB_TYPES if mode == "stub" else BANNER_TYPES
        # Normalize type for matching
        type_norm = type_name.replace("_", "").replace(" ", "")
        allowed_norm = {t.lower().replace("_", "") for t in allowed}
        if type_norm not in allowed_norm:
            continue

        # Detect matrix rows
        choices_i18n = q.get("choices_i18n") or {}
        rows = choices_i18n.get("rows") or {}
        row_count = len(rows) if isinstance(rows, dict) else 0

        results.append({
            "position": position,
            "label": label,
            "type": type_name,
            "title": title,
            "row_count": row_count,
        })

    results.sort(key=lambda x: x["position"])
    return results


def list_from_metadata(data: dict, mode: str):
    """Parse surveyflow metadata.json format."""
    questions_raw = data.get("questions", {})
    if isinstance(questions_raw, dict):
        questions = list(questions_raw.values())
    else:
        questions = questions_raw

    results = []
    for q in questions:
        answer_type = q.get("answer_type") or q.get("type") or ""
        label = q.get("label") or ""
        title_raw = q.get("title") or q.get("question_i18n") or {}
        title = _label_text(title_raw)
        position = q.get("position", 0)

        # Check filter
        allowed = STUB_TYPES if mode == "stub" else BANNER_TYPES
        if answer_type not in allowed:
            continue

        # Detect matrix rows
        choices_i18n = q.get("choices_i18n") or {}
        rows = choices_i18n.get("rows") or {}
        row_count = len(rows) if isinstance(rows, dict) else 0

        results.append({
            "position": position,
            "label": label,
            "type": answer_type,
            "title": title,
            "row_count": row_count,
        })

    results.sort(key=lambda x: x["position"])
    return results


def main():
    parser = argparse.ArgumentParser(description="List survey questions by type")
    parser.add_argument("file", help="Path to definition.json or metadata.json")
    parser.add_argument("--type", choices=["stub", "banner"], default="stub",
                        help="Filter: stub (SA/MA/Matrix) or banner (SA/MA only)")
    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists():
        print(f"ERROR: File not found: {path}", file=sys.stderr)
        sys.exit(1)

    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    # Detect format
    if "questions" in data and isinstance(data.get("questions"), (list, dict)):
        # Try to detect by presence of type_name (definition) vs answer_type (metadata)
        q_sample = None
        qs = data["questions"]
        if isinstance(qs, dict):
            q_sample = next(iter(qs.values()), None)
        elif isinstance(qs, list) and qs:
            q_sample = qs[0]

        if q_sample and "type_name" in q_sample:
            results = list_from_definition(data, args.type)
        else:
            results = list_from_metadata(data, args.type)
    else:
        print("ERROR: Unrecognized file format (expected definition.json or metadata.json)", file=sys.stderr)
        sys.exit(1)

    if not results:
        print(f"No {args.type} questions found.")
        return

    print(f"\n{'Pos':>4}  {'Label':<15} {'Type':<20} {'Title'}")
    print("-" * 80)
    for q in results:
        row_info = f"  [{q['row_count']} rows]" if q["row_count"] else ""
        print(f"{q['position']:>4}  {q['label']:<15} {q['type']:<20} {q['title']}{row_info}")
    print(f"\nTotal: {len(results)} questions")


if __name__ == "__main__":
    main()
