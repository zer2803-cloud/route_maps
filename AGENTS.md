# AGENTS.md

## Cursor Cloud specific instructions

### What this repo contains
- The `main` branch holds only route-map image assets (`overview.png`, `day1.png`–`day4.png`, `第1天.png`–`第4天.png`, `appendix.png`, `附加(可选).png`). There is nothing to build or run on `main` alone.
- The actual application lives on branch `cursor/pab-plan-xlsx-6614`: a Python tool that generates a formatted Excel workbook (`PAB产品计划表.xlsx`) from hardcoded PAB product-schedule data using `openpyxl`. Key files (on that branch): `scripts/build_pab_plan.py`, `scripts/__init__.py`, `tests/test_build_pab_plan.py`.

### Runtime / tooling
- Python 3.12 (system `python3`) with `openpyxl` and `pytest`. There is no dependency manifest or lockfile in the repo, so the environment update script installs these two packages directly.
- No services, databases, servers, or network access are required. The app is a one-shot offline script that writes a local `.xlsx` file.

### Running the app and tests (from the branch that has the code)
- Run everything from the repository root so that `scripts/` is importable as a package (the tests do `from scripts.build_pab_plan import ...`).
- Generate the workbook: `python3 -m scripts.build_pab_plan` (writes `PAB产品计划表.xlsx` to the repo root, prints `已生成：<path>`).
- Run tests: `python3 -m pytest`.

### Non-obvious gotchas
- `pip install` places the `pytest` console script in `~/.local/bin`, which is not on `PATH`. Invoke it as `python3 -m pytest` instead of `pytest`.
- The code is on the `cursor/pab-plan-xlsx-6614` branch, not `main`. To run/test it while `main` is checked out, use a worktree, e.g. `git worktree add /tmp/pab-feature origin/cursor/pab-plan-xlsx-6614`, or check out that branch.
