#!/usr/bin/env python3


from __future__ import annotations
import argparse
import fnmatch
import os
from pathlib import Path
from typing import Iterable, List, Set, Tuple

# ----------------------------
# Defaults (sane and safe)
# ----------------------------

DEFAULT_IGNORE_DIRS: Set[str] = {
    ".git", ".hg", ".svn",
    ".venv", "venv", "env",
    "build", "dist", "__pycache__",
    ".pytest_cache", ".mypy_cache", ".idea", ".vscode",
}

DEFAULT_IGNORE_NAMES: Set[str] = {
    "token.json",  # OAuth tokens
}

DEFAULT_IGNORE_EXTS: Set[str] = {
    ".pyc", ".pyo", ".pyd",
    ".so", ".dll", ".dylib",
    ".zip", ".tar", ".gz", ".7z",
    ".exe", ".bin",
}

# Basic language hints for code fences
LANG_BY_EXT = {
    ".py": "python",
    ".md": "markdown",
    ".txt": "text",
    ".bat": "bat",
    ".cmd": "bat",
    ".sh": "bash",
    ".ps1": "powershell",
    ".json": "json",
    ".ini": "ini",
    ".env": "ini",
    ".yml": "yaml",
    ".yaml": "yaml",
    ".csv": "csv",
    ".ts": "ts",
    ".tsx": "tsx",
    ".js": "javascript",
    ".jsx": "jsx",
    ".html": "html",
    ".css": "css",
    ".toml": "toml",
}


def match_any(path: Path, patterns: Iterable[str]) -> bool:
    """Return True if path matches ANY of the glob patterns (POSIX-style)."""
    s = path.as_posix()
    for pat in patterns:
        if fnmatch.fnmatch(s, pat):
            return True
    return False


def build_tree(root: Path,
               ignore_dirs: Set[str],
               ignore_exts: Set[str],
               ignore_names: Set[str],
               include_globs: List[str] | None,
               exclude_globs: List[str] | None,
               max_bytes: int) -> Tuple[str, List[Path]]:
    """
    Return a (tree_text, files_list) tuple.
    files_list contains all files to embed in the bundle.
    """
    lines_tree: List[str] = []
    files_out: List[Path] = []

    for dirpath, dirnames, filenames in os.walk(root):
        # prune ignored directories
        dirnames[:] = [
            d for d in dirnames
            if d not in ignore_dirs and not d.startswith(".")
        ]
        cur = Path(dirpath)
        rel_dir = cur.relative_to(root)
        depth = 0 if rel_dir == Path(".") else len(rel_dir.parts)
        indent = "  " * depth
        if rel_dir != Path("."):
            lines_tree.append(f"{indent}{rel_dir.as_posix()}/")

        for fname in sorted(filenames):
            p = cur / fname

            # Ignore by name/extension
            if fname in ignore_names:
                continue
            if p.suffix.lower() in ignore_exts:
                continue

            # Exclude hidden top-level noise by pattern
            if exclude_globs and match_any(p.relative_to(root), exclude_globs):
                continue
            # If include globs were given, only take matches
            if include_globs and not match_any(p.relative_to(root), include_globs):
                continue

            # Size check
            try:
                if p.stat().st_size > max_bytes:
                    # Do not list it in files_out, but show in tree (optional)
                    entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
                    lines_tree.append(entry + "  [skipped: too large]")
                    continue
            except Exception:
                # If can't stat, skip silently
                continue

            # Accept file
            entry = fname if rel_dir == Path(".") else f"{indent}  {fname}"
            lines_tree.append(entry)
            files_out.append(p)

    return "\n".join(lines_tree), files_out


def read_text_safe(p: Path) -> str | None:
    try:
        return p.read_text(encoding="utf-8")
    except UnicodeDecodeError:
        # Binary or non-UTF8; skip
        return None
    except Exception:
        return None


def make_bundle_markdown(root: Path,
                         out: Path,
                         include_globs: List[str] | None,
                         exclude_globs: List[str] | None,
                         max_bytes: int,
                         ignore_dirs: Set[str],
                         ignore_exts: Set[str],
                         ignore_names: Set[str]) -> Path:
    tree_text, files_to_embed = build_tree(
        root, ignore_dirs, ignore_exts, ignore_names,
        include_globs, exclude_globs, max_bytes
    )

    header = f"""# Project Snapshot (CodeBundle)

                This single Markdown file contains a **self-contained snapshot** of your project so another AI/engineer can review or modify it without needing the original folder.

                **How to use this file with an AI**
                1. Upload or paste this file as a single attachment.
                2. Ask for changes; the AI can reference specific `file:` sections below.
                3. Copy updated blocks back into the corresponding files in your project.

                > Notes: secrets like `token.json` are intentionally excluded. Virtual envs and build artifacts are omitted to keep this readable.

                ## Directory tree (filtered)
                {tree_text}"""

    sections: List[str] = [header]

    for p in sorted(files_to_embed, key=lambda x: x.relative_to(root).as_posix()):
        rel = p.relative_to(root)
        lang = LANG_BY_EXT.get(p.suffix.lower(), "")
        content = read_text_safe(p)
        if content is None:
            # Skip non-text files silently
            continue
        fence = "```"
        sections.append(
            f"\n---\n### file: {rel.as_posix()}\n\n{fence}{lang}\n{content}\n{fence}\n"
        )

    out.write_text("".join(sections), encoding="utf-8")
    return out


def parse_args() -> argparse.Namespace:
    ap = argparse.ArgumentParser(
        description="Generate a single-file Markdown 'CodeBundle' of a project."
    )
    # Defaults are computed in main()
    ap.add_argument(
        "--root",
        default=None,
        help="Project root folder (default: parent of this script)",
    )
    ap.add_argument(
        "--out",
        default=None,
        help="Output Markdown file path (default: PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory)",
    )
    ap.add_argument("--include", nargs="*", default=None,
                    help='Optional list of glob patterns to include (e.g. "**/*.py" "**/*.md").')
    ap.add_argument("--exclude", nargs="*", default=None,
                    help='Optional list of glob patterns to exclude (e.g. ".venv/**" "dist/**").')
    ap.add_argument("--max-bytes", type=int, default=1_000_000,
                    help="Max file size to embed (bytes). Oversized files are listed but skipped. Default: 1,000,000.")
    ap.add_argument("--no-default-ignores", action="store_true",
                    help="Disable default ignore sets for dirs/exts/names.")
    return ap.parse_args()


def main():
    print("CodeBundle writing starting...")
    args = parse_args()

    # Locate the script and its parent (project root)
    script_dir = Path(__file__).resolve().parent          # .../favtrip_reporting/documentation
    project_root = script_dir.parent                      # .../favtrip_reporting

    # ROOT: default to the parent of this script unless overridden
    if args.root is None:
        root = project_root
    else:
        root_arg = Path(args.root)
        root = root_arg.resolve() if root_arg.is_absolute() else (Path.cwd() / root_arg).resolve()

    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Root does not exist or is not a directory: {root}")

    # OUT: default to PROJECT_SNAPSHOT_CODEBUNDLE.md in the script's directory (the child)
    if args.out is None:
        out = (script_dir / "PROJECT_SNAPSHOT_CODEBUNDLE.md").resolve()
    else:
        out_arg = Path(args.out)
        # If user gave a relative path, resolve it under the chosen root; absolute paths are respected
        out = out_arg.resolve() if out_arg.is_absolute() else (root / out_arg).resolve()

    # ... keep the rest of your function unchanged ...
    if args.no_default_ignores:
        ignore_dirs = set()
        ignore_exts = set()
        ignore_names = set()
    else:
        ignore_dirs = set(DEFAULT_IGNORE_DIRS)
        ignore_exts = set(DEFAULT_IGNORE_EXTS)
        ignore_names = set(DEFAULT_IGNORE_NAMES)

    out.parent.mkdir(parents=True, exist_ok=True)
    result_path = make_bundle_markdown(
        root=root,
        out=out,
        include_globs=args.include,
        exclude_globs=args.exclude,
        max_bytes=args.max_bytes,
        ignore_dirs=ignore_dirs,
        ignore_exts=ignore_exts,
        ignore_names=ignore_names,
    )
    print(f"CodeBundle written to: {result_path}")


if __name__ == "__main__":
    main()

#cd C:\Users\rjrul\Downloads\favtrip_pipeline\documentation\
#python generate_code_bundle.py