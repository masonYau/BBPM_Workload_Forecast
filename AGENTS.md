# AGENTS.md

## Commit Workflow

- Before creating any `git commit`, run the Python-to-Markdown export script from the repository root.
- In this Codex desktop workspace, use:
  `C:\Users\HP\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe .\copy_py_to_md.py`
- If that bundled Python path is unavailable in another environment, use any working Python interpreter to run:
  `python .\copy_py_to_md.py`
- After the script runs, include all generated or updated `copy/*.md` files in the commit.
- Do not add the copied `copy/*.py` files to git unless the user explicitly asks for them.
- If the export script fails, stop and report the failure instead of committing stale Markdown files.
