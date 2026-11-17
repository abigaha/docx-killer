# DOCX Killer

English | [简体中文](README.zh-CN.md)

A skill for **Claude Code, Codex, and DSH** that works with Word documents using **python-docx**, **docxtpl**, **Pandoc**, and **npm Mammoth**, with targeted OOXML editing and explicit fidelity checks.

It supports document inspection, structural editing guidance, conservative cross-run text replacement, strict template rendering, semantic conversion, and package comparison. It does not promise universally lossless editing or implement a Word layout engine.

## Install the skill

Clone the repository and run the standard-library Python installer. By default, it installs the local skill into all three agents:

```powershell
git clone https://github.com/abigaha/docx-killer.git
cd docx-killer
python install.py
```

To select one or more agents:

```powershell
python install.py --agent codex
python install.py --agent claude --agent dsh
python install.py --help
```

| Agent | Default destination |
|---|---|
| Claude Code (`claude`) | `~/.claude/skills/docx-killer` |
| Codex (`codex`) | `~/.codex/skills/docx-killer`, or `$CODEX_HOME/skills/docx-killer` when set |
| DSH (`dsh`) | `~/.dsh/skills/docx-killer` |

The installer uses the repository next to `install.py`, so it also works when launched from another directory. It skips matching installations and backs up an existing directory before replacing a changed installation. Backups live in the agent configuration directory's `skill-backups` folder, outside the discoverable `skills` directory; the installer prints their paths. DSH's existing custom description is retained while the skill body and supporting files are updated. DSH descriptions may use one line or an indented `>` / `|` block; multiline quoted descriptions are refused without changing the existing installation. Cache files are excluded.

The installer copies the skill; install the dependencies separately below. Start a new agent session after installation so the updated skill can be loaded, then ask for a matching DOCX task. The repository contains no user documents or credentials.

## Dependencies

- Python with `python-docx`, `docxtpl`, `lxml`, and `Jinja2`; compatible ranges ship in `skills/docx-killer/requirements.txt`. The root `requirements.txt` includes that file for repository installs.
- Pandoc executable.
- Node.js with `npm install -g mammoth` or a Mammoth package selected with `MAMMOTH_MODULE`. For a project-local installation, set this variable to its absolute package directory; Node resolves modules relative to the skill script, not the working directory.
- An office application for actual layout verification, when required; the four tools alone do not render Word pages.

```powershell
python -m pip install -r requirements.txt
python skills/docx-killer/scripts/doctor.py
python skills/docx-killer/scripts/docx_ops.py --help
```

The dependency constraints also ship inside the installed skill as `requirements.txt`. Poppler's `pdfinfo` and `pdftoppm` are needed for PDF metadata and page previews.

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

Use trusted templates and JSON context data. Template rendering rejects missing variables and escapes XML-sensitive values. Mammoth HTML is semantic output, not a visual rendering, and is not sanitized.

## Validation

```powershell
python -m unittest discover -s tests -v
```

The integration tests require all four installed tools and create their own synthetic DOCX files in temporary directories. A real WPS Office 12.1.0.29690 run exported a two-page, two-section document and checked page images, Chinese text, tables, an image, headers, footers, footnotes, and endnotes. This is evidence for that WPS workflow, not a guarantee for every document or a separate Microsoft Word installation.

Read [the skill](skills/docx-killer/SKILL.md) for tool selection, [verification](skills/docx-killer/references/verification.md) for fidelity limits, and [validation evidence](docs/validation.md) for the checks actually performed.
