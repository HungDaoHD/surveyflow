"""
package_skill.py — Package a skill directory into a .skill zip file.

Usage:
  python skills/package_skill.py surveyflow
  python skills/package_skill.py surveyflow --output dist/surveyflow.skill

The .skill file is a ZIP archive containing SKILL.md + scripts/ + references/ + assets/
"""

import zipfile
import shutil
import argparse
from pathlib import Path


def package_skill(skill_name: str, output_path: Path | None = None) -> Path:
    skills_dir = Path(__file__).parent
    skill_dir = skills_dir / skill_name

    if not skill_dir.exists():
        raise FileNotFoundError(f"Skill directory not found: {skill_dir}")

    skill_md = skill_dir / "SKILL.md"
    if not skill_md.exists():
        raise FileNotFoundError(f"SKILL.md not found: {skill_md}")

    if output_path is None:
        output_path = skills_dir / f"{skill_name}.skill"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for file in skill_dir.rglob("*"):
            if file.is_file() and "__pycache__" not in str(file):
                arcname = file.relative_to(skill_dir)
                zf.write(file, arcname)
                print(f"  + {arcname}")

    print(f"\nPackaged: {output_path}  ({output_path.stat().st_size / 1024:.1f} KB)")
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Package skill into .skill file")
    parser.add_argument("skill", help="Skill name (subfolder of skills/)")
    parser.add_argument("--output", "-o", help="Output path (default: skills/<name>.skill)")
    args = parser.parse_args()

    output = Path(args.output) if args.output else None
    package_skill(args.skill, output)


if __name__ == "__main__":
    main()
