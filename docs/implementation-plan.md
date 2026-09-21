# DOCX Killer implementation plan

Goal: deliver a discoverable Windows DOCX skill backed up to the user's existing GitHub repository.

Approved design: combine python-docx structural editing, docxtpl template rendering, Pandoc conversion, and Mammoth semantic extraction. Use surgical OOXML edits when preservation matters. State limitations rather than promise arbitrary lossless editing.

Architecture: `skills/docx-killer/SKILL.md` routes to focused references and executable helpers. Tests build synthetic documents; no personal documents or local installation paths are committed. A separate local copy is installed into the runtime skill directory after validation.

## Progress

- [x] Confirm Windows environment and authorized GitHub destination.
- [x] Locate python-docx 1.2.0, docxtpl 0.20.2, Pandoc 3.11, and npm Mammoth 1.12.3.
- [x] Run independent no-skill reference scenarios. Baseline already chose the correct tools; its sample replacement wrote output before all checks finished, and allowed replacements across bookmark/hyperlink boundaries. Test those concrete gaps; do not claim a general reasoning failure.
- [x] Read authoritative documentation and record source links and version boundaries.
- [x] Add behavior tests for exact replacements, protected structures, unchanged package parts, strict templates, and failure outputs; observe missing-helper failures before implementation.
- [x] Implement package inspection/comparison, conservative text replacement, strict template rendering, Windows discovery, and npm Mammoth conversion.
- [x] Write compact skill routing and focused references for the four tools, OOXML, and visual verification.
- [x] Exercise all four installed tools against generated examples. An independent forward-test started but its isolated command environment failed before producing a result; automated integration tests and manual WPS rendering provide the recorded evidence instead.
- [x] Validate skill metadata and install/verify the local skill. Commit, push, and remote verification are the final delivery actions.

## Verification contract

For preservation edits: assert exact replacement count, original file hash, run formatting, relationship/media hashes, absence of partial output on failure, and refusal to cross unsupported structures. For templates: missing variables fail, XML-sensitive values survive, and repeating rows are populated. For conversions: actual installed binaries must extract/produce content; conversions do not establish layout fidelity. Visual verification is reported separately from ZIP/XML checks.
