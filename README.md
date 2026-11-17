# DOCX Killer

A Codex skill for practical Word document work using **python-docx**, **docxtpl**, **Pandoc**, and **npm Mammoth**, with targeted OOXML editing and explicit fidelity checks.

It supports document inspection, structural editing guidance, conservative cross-run text replacement, strict template rendering, semantic conversion, and package comparison. It does not promise universally lossless editing or implement a Word layout engine.

## Install the skill

Copy `skills/docx-killer` into your runtime's skills directory, normally `~/.codex/skills/docx-killer`. Start a new session if the current session's skill list was already loaded. Invoke `$docx-killer` or ask for a DOCX task that matches its description.

Do not overwrite an existing locally modified skill without comparing it first. The repository contains no user documents or credentials.

## Dependencies

- Python with `python-docx`, `docxtpl`, `lxml`, and `Jinja2`; compatible ranges ship in `skills/docx-killer/requirements.txt`. The root `requirements.txt` includes that file for repository installs.
- Pandoc executable.
- Node.js with `npm install -g mammoth` or a Mammoth package selected with `MAMMOTH_MODULE`. For a project-local installation, set this variable to its absolute package directory; Node resolves modules relative to the skill script, not the working directory.
- An office application for actual layout verification, when required; the four tools alone do not render Word pages.

```powershell
python -m pip install -r skills/docx-killer/requirements.txt
python skills/docx-killer/scripts/doctor.py
python skills/docx-killer/scripts/docx_ops.py --help
python -m unittest discover -s tests -v
```

The integration tests intentionally require all four installed tools. Tests synthesize their own DOCX files in temporary directories.

## Examples

```powershell
python skills/docx-killer/scripts/docx_ops.py inspect input.docx --text
python skills/docx-killer/scripts/docx_ops.py inspect input.docx --summary --text --story word/document.xml --paragraph-range 0:20 --tables --media
python skills/docx-killer/scripts/docx_ops.py replace input.docx edited.docx --old 'Old name' --new 'New name' --expected 2
python skills/docx-killer/scripts/docx_ops.py render template.docx context.json generated.docx
python skills/docx-killer/scripts/docx_ops.py compare input.docx edited.docx
python skills/docx-killer/scripts/docx_ops.py verify generated.docx --render auto --output-dir render-output
node skills/docx-killer/scripts/mammoth_extract.cjs input.docx preview.html
```

For a new or modified DOCX, `verify` checks package structure and reopens it with `python-docx`. Read-only inspection and HTML/Markdown/text conversion should instead verify their actual findings or output artifact. When an Office renderer is available, `verify` also exports a PDF and, when Poppler is installed, creates one PNG preview per page. Each renderer attempt uses a dedicated empty work directory, and an existing destination PDF is never overwritten. On Windows, auto rendering tries `KWPS.Application` before `Word.Application`, then LibreOffice, and records the selected COM ProgID. `--render word` tries only `Word.Application`; use `--render auto` for WPS. The report distinguishes skipped, unavailable, failed, and rendered states and retains failed-attempt reasons after fallback. PDF export does not by itself prove correct pagination or absence of visual overlap.

Read [the skill](skills/docx-killer/SKILL.md) for selection rules and [verification](skills/docx-killer/references/verification.md) for the limits of structural versus visual checks.
