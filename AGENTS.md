# AGENTS.md

## Commit Workflow

- Before creating any `git commit`, run the Python-to-Markdown export script from the repository root.
- In this Codex desktop workspace, use:
  `C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\copy_py_to_md.py`
- If that bundled Python path is unavailable in another environment, use any working Python interpreter to run:
  `python .\copy_py_to_md.py`
- After the script runs, include all generated or updated `copy/*.md` files in the commit.
- The `copy/` directory is for Markdown archives. Do not add `copy/*.py` files to git.
- If any `copy/*.py` files are created or left over, delete them after generating the Markdown archives.
- If the export script fails, stop and report the failure instead of committing stale Markdown files.
