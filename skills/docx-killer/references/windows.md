# Windows environment and reproducible invocation

Run `python scripts/doctor.py` from the skill directory. It reports the active interpreter, package versions, discovered Pandoc executable, npm Mammoth module, and potential office renderers. It does not install software or change PATH. Use absolute script paths when working elsewhere.

Tested combination: Python 3.14.7, python-docx 1.2.0, docxtpl 0.20.2, lxml 6.1.3, Jinja2 3.1.6, Pandoc 3.11, Node Mammoth 1.12.3. These are verified versions, not a promise that other versions behave identically.

## Python

```powershell
Get-Command python
py -0p
python -m pip show python-docx docxtpl lxml Jinja2
```

Use `python -m pip` with the interpreter that runs the scripts. The template package is named **docxtpl**, not `python-docx-template`. No Python `mammoth` package is required for the included Node wrapper.

If installation is requested on a new computer:

```powershell
python -m pip install -r requirements.txt
```

Run this from the installed skill directory, whose `requirements.txt` ships the compatible ranges. Use a virtual environment for project isolation when appropriate; do not upgrade an existing working global environment unnecessarily.

## Pandoc

After winget installs Pandoc, an existing terminal may retain its old PATH. The doctor probes PATH, then `%LOCALAPPDATA%\Pandoc\pandoc.exe` and Program Files locations. `PANDOC_BINARY` can specify a custom absolute executable path and overrides discovery.

```powershell
$pandocExe = Join-Path $env:LOCALAPPDATA 'Pandoc\pandoc.exe'
& $pandocExe --version
```

If a new installation is requested, `winget install --source winget --exact --id JohnMacFarlane.Pandoc` is an option. A Chocolatey access-denied/lock error is not proof that Pandoc is unavailable elsewhere; inspect the successful installation and do not delete system lock files speculatively.

## Mammoth is npm here

```powershell
npm list -g --depth=0
npm root -g
node scripts/mammoth_extract.cjs --doctor
```

Node's `require('mammoth')` does not normally search npm global packages. The wrapper tries module resolution from the skill script's location, `%APPDATA%\npm\node_modules\mammoth`, and the Node installation directory. A package installed only in the current project's `node_modules` needs an explicit absolute `MAMMOTH_MODULE` path. For a custom global prefix:

```powershell
$npmGlobalRoot = npm root -g
$env:MAMMOTH_MODULE = Join-Path $npmGlobalRoot 'mammoth'
node scripts/mammoth_extract.cjs input.docx preview.html
```

For a project-local installation, set `$env:MAMMOTH_MODULE = (Resolve-Path './node_modules/mammoth').Path` from that project. The override expects a trusted installed package directory, not an arbitrary downloaded JS file. It applies only to that terminal session. Do not add user-specific absolute paths to the shared skill.

## Encoding and output files

Use UTF-8 JSON context files; the template CLI also accepts UTF-8 BOM. Use quoted path arguments with spaces. Use Pandoc `-o` for DOCX output: never send binary DOCX through PowerShell text redirection. Avoid shell interpolation for document text containing `$`, quotes, or backticks; call the Python API from a script for complicated replacement strings.

Python helpers stage output in the destination directory, inspect it, and publish with a hard link without overwriting. Use a local filesystem supporting hard links, such as NTFS. On a filesystem without hard links, choose a local working directory and then copy the completed output intentionally.
