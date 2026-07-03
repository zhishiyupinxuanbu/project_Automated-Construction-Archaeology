# Changelog

## v1.0.1 - 2026-07-03

- Added Chinese UI metadata in `agents/openai.yaml`.
- Strengthened Chinese trigger phrases in `SKILL.md` description.
- Kept the stable English skill `name` and folder slug for compatibility.

## v1.0 - 2026-06-29

- Formalized `smart-prospecting-report` as the current local delivery version.
- Kept local traceability files: `VERSION`, `MANIFEST.json`, `CHANGELOG.md`, and `DEVELOPMENT_LOG.md`.
- Kept `scripts/verify_manifest.py` for migration integrity checks using file sizes and SHA-256 hashes.
- Kept `scripts/self_test_skill.py` as the package structure and safety self-test.
- Default final report outputs remain on the user's Desktop.

## v0.9 - 2026-06-29

- Added `references/test-plan.md` for install, structure, doctor, smoke, and trigger tests.
- Added `references/safety-checklist.md` for package, runtime, API, and content safety.
- Added `scripts/self_test_skill.py` for offline package validation.
- Split the original smart prospecting workflow into an independent `smart-prospecting-report` skill.
- Bundled only the assets needed for 报告 generation.
- Patched portable script paths so the skill does not depend on the original project workspace.
