# Project Structure Review – Excel Explorer

_Date generated: 2025-06-14_

## 1. Purpose
Review the current root-level folder structure against the authoritative `project-documentation.md` to ensure it contains **only** the elements required to reach MVP quickly. Anything not defined in the documentation (and therefore not essential for MVP) is flagged for **deletion**.

## 2. Documentation-Defined Folders / Files (KEEP)
The following items appear in the project documentation and are present in the repository. They are **aligned** and should be retained:

| Path | Reason |
|------|--------|
| `src/` (including `core/`, `modules/`, `utils/`) | Main source code as described in docs |
| `config/` | YAML configuration files referenced in docs |
| `output/` (`reports/`, `structured/`, `cache/`) | Destination for generated documentation |
| `.gitignore` | Standard VCS hygiene |
| `README.md` | High-level overview |
| `requirements.txt` | Python dependencies |
| `project-documentation.md` | Source of truth for project goals |

## 3. Items **Not** in Documentation (SUGGEST DELETE)
The following artefacts are **not** mentioned in `project-documentation.md` and are **non-essential** for MVP. Delete or move them out of the repository until post-MVP phase.

| Path | Rationale |
|------|-----------|
| `docs/` (entire folder) | Additional documentation – postpone until after MVP |
| `scripts/` | Helper/self-test scripts not covered in docs; can be re-added later if needed |
| `schema/` | JSON schema validation utilities – not required for initial MVP logic |
| `create_initial_setup.ps1` | One-off setup script; keep externally or document in README if still needed |
| `mvp_*/*.md`, `phase0_*/*.md` | Planning/status markdown files – move to project management tool or delete |
| `venv/` (Python virtual environment) | Should be ignored via `.gitignore`, not committed |
| `_deprecated_github/` & any `workflows/` dirs | CI/CD workflows explicitly out-of-scope now |

## 4. Summary Recommendations
1. **Delete** all items listed in Section 3 to simplify the codebase.
2. Ensure `.gitignore` excludes virtual-environment directories (e.g. `venv/`, `.env/`).
3. If any removed component proves necessary, add it **and** document it in `project-documentation.md` before re-introduction.

Removing the non-essential folders/files will leave a minimal, clean structure focused solely on building the MVP.
