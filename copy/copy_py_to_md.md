# copy_py_to_md.py

```python
from pathlib import Path
import shutil


PROJECT_ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = PROJECT_ROOT / "copy"
SKIP_DIRS = {
    ".git",
    ".idea",
    ".mypy_cache",
    ".pytest_cache",
    ".ruff_cache",
    ".tox",
    ".venv",
    "__pycache__",
    "copy",
    "venv",
}


def should_skip(path: Path) -> bool:
    return any(part in SKIP_DIRS for part in path.relative_to(PROJECT_ROOT).parts)


def iter_python_files() -> list[Path]:
    return sorted(
        path
        for path in PROJECT_ROOT.rglob("*.py")
        if path.is_file() and not should_skip(path)
    )


def markdown_fence(code: str) -> str:
    longest_backticks = 0
    current = 0
    for char in code:
        if char == "`":
            current += 1
            longest_backticks = max(longest_backticks, current)
        else:
            current = 0
    return "`" * max(3, longest_backticks + 1)


def write_markdown(source_path: Path, md_path: Path, code: str) -> None:
    relative_source = source_path.relative_to(PROJECT_ROOT).as_posix()
    fence = markdown_fence(code)
    md_path.write_text(
        f"# {relative_source}\n\n{fence}python\n{code.rstrip()}\n{fence}\n",
        encoding="utf-8",
    )


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)

    copied_count = 0
    for source_path in iter_python_files():
        relative_path = source_path.relative_to(PROJECT_ROOT)
        copied_path = OUTPUT_DIR / relative_path
        md_path = copied_path.with_suffix(".md")

        copied_path.parent.mkdir(parents=True, exist_ok=True)
        code = source_path.read_text(encoding="utf-8")
        shutil.copy2(source_path, copied_path)
        write_markdown(source_path, md_path, code)
        copied_count += 1

    print(f"Copied {copied_count} Python files and generated {copied_count} Markdown files in {OUTPUT_DIR}")


if __name__ == "__main__":
    main()
```
