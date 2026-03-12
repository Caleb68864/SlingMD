---
type: phase-spec-index
master_spec: "docs/specs/2026-03-12-slingmd-reliability-hardening.md"
date: 2026-03-12
sub_specs: 3
---

# SlingMD Reliability Hardening -- Phase Specs

Refined from [2026-03-12-slingmd-reliability-hardening.md](../2026-03-12-slingmd-reliability-hardening.md).

| Sub-Spec | Title | Dependencies | Phase Spec |
|----------|-------|--------------|------------|
| 1 | Startup and Export Flow Safeguards | none | [sub-spec-1-startup-and-export-flow-safeguards.md](sub-spec-1-startup-and-export-flow-safeguards.md) |
| 2 | File, Metadata, Attachment, and UI Hardening | 1 | [sub-spec-2-file-metadata-attachment-and-ui-hardening.md](sub-spec-2-file-metadata-attachment-and-ui-hardening.md) |
| 3 | Coverage, Regression Tests, and Verification Guidance | 2 | [sub-spec-3-coverage-regression-tests-and-verification-guidance.md](sub-spec-3-coverage-regression-tests-and-verification-guidance.md) |

## Execution

Run `/forge-run docs/specs/slingmd-reliability-hardening/` to execute all phase specs.
Run `/forge-run docs/specs/slingmd-reliability-hardening/ --sub N` to execute a single sub-spec.
