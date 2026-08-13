"""
package_skill.py — Package a skill directory into a .skill zip file.

Usage:
  python tools/package_skill.py fieldcheck-dp
  python tools/package_skill.py fieldcheck-dp --output dist/fieldcheck-dp.skill

A .skill file is a plain ZIP holding the whole skill folder: SKILL.md at the
archive root, plus scripts/ / references/ / assets/ preserved as-is.

ALWAYS pack with this script rather than zipping SKILL.md by hand. Hand-packing
is how the resource bundle silently went missing for 24 consecutive builds
(2026-07-02 .. 2026-08-11): every archive in that stretch held nothing but a
flat SKILL.md, so every `scripts/...` command the skill documented pointed at a
file that was not shipped.
"""

import argparse
import sys
import zipfile
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent
SKILLS_DIR = REPO_ROOT / ".skill_work"

EXCLUDE_DIRS = {"__pycache__", ".git", "node_modules"}
EXCLUDE_SUFFIXES = {".pyc", ".pyo"}


def _is_excluded(path: Path, skill_dir: Path) -> bool:
    rel = path.relative_to(skill_dir)
    if any(part in EXCLUDE_DIRS for part in rel.parts):
        return True
    return path.suffix in EXCLUDE_SUFFIXES


def package_skill(skill_name: str, output_path: Path | None = None) -> Path:
    skill_dir = SKILLS_DIR / skill_name
    if not skill_dir.is_dir():
        raise FileNotFoundError(f"Skill directory not found: {skill_dir}")

    skill_md = skill_dir / "SKILL.md"
    if not skill_md.is_file():
        raise FileNotFoundError(f"SKILL.md not found: {skill_md}")

    if output_path is None:
        output_path = SKILLS_DIR / f"{skill_name}.skill"
    output_path.parent.mkdir(parents=True, exist_ok=True)

    files = sorted(
        f for f in skill_dir.rglob("*")
        if f.is_file() and not _is_excluded(f, skill_dir)
    )

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in files:
            arcname = file.relative_to(skill_dir).as_posix()
            zf.write(file, arcname)
            print(f"  + {arcname:<50} {file.stat().st_size:>7} B")

    size_kb = output_path.stat().st_size / 1024
    print(f"\nPackaged: {output_path}  ({len(files)} entries, {size_kb:.1f} KB)")
    return output_path


def main() -> None:
    parser = argparse.ArgumentParser(description="Package a skill folder into a .skill file")
    parser.add_argument("skill", help="Skill name (subfolder of .skill_work/)")
    parser.add_argument("--output", "-o", help="Output path (default: .skill_work/<name>.skill)")
    args = parser.parse_args()

    try:
        package_skill(args.skill, Path(args.output) if args.output else None)
    except FileNotFoundError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
