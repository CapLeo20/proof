#!/usr/bin/env python3
"""Create and list tasks under tasks/.

Usage:
    python3 scripts/tasks.py new <slug> [title]   create tasks/<YYYY-MM-DD>-<slug>/ from tasks/_template
    python3 scripts/tasks.py list                print all tasks as a Markdown table
"""
import datetime
import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
TASKS = ROOT / "tasks"
TEMPLATE = TASKS / "_template"
SUBDIRS = ("input", "output")


def front_matter(readme):
    """Return the `key: value` pairs between the leading --- lines of a README."""
    lines = readme.read_text(encoding="utf-8").splitlines()
    if not lines or lines[0].strip() != "---":
        return {}
    meta = {}
    for line in lines[1:]:
        if line.strip() == "---":
            break
        key, sep, value = line.partition(":")
        if sep:
            meta[key.strip()] = value.strip()
    return meta


def new(slug, title=None):
    if not re.fullmatch(r"[a-z0-9]+(-[a-z0-9]+)*", slug):
        sys.exit(f"slug must be lowercase letters, digits and hyphens: {slug!r}")
    today = datetime.date.today().isoformat()
    task = TASKS / f"{today}-{slug}"
    if task.exists():
        sys.exit(f"already exists: {task.relative_to(ROOT)}")
    for sub in SUBDIRS:
        (task / sub).mkdir(parents=True)
        (task / sub / ".gitkeep").touch()  # keep empty folders in git so files can be uploaded into them
    text = (TEMPLATE / "README.md").read_text(encoding="utf-8")
    text = text.replace("{{title}}", title or slug).replace("{{date}}", today)
    (task / "README.md").write_text(text, encoding="utf-8")
    print(task.relative_to(ROOT))


def list_tasks():
    print("| 任务 | 标题 | 状态 | 创建 |")
    print("|---|---|---|---|")
    for task in sorted(p for p in TASKS.iterdir() if p.is_dir() and not p.name.startswith("_")):
        readme = task / "README.md"
        meta = front_matter(readme) if readme.exists() else {}
        print(f"| {task.name} | {meta.get('title', '')} | {meta.get('status', '?')} | {meta.get('created', '')} |")


if __name__ == "__main__":
    args = sys.argv[1:]
    if args[:1] == ["new"] and len(args) in (2, 3):
        new(*args[1:])
    elif args == ["list"]:
        list_tasks()
    else:
        sys.exit(__doc__)
