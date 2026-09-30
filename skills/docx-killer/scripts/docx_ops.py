"""Inspect DOCX, replace plain text conservatively, and render trusted templates."""
import argparse
import fnmatch
import json
import locale
import os
import re
import sys
import tempfile
import subprocess
import shutil
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from docx_package import W, NS, compare_packages, inspect_entries, inspect_package, load_package, paragraph_text, parse_xml


def decode_output(data):
    return data.decode(locale.getpreferredencoding(False), errors='replace')


def inspect_document(source, include_text=False, summary=False, stories=None, paragraph_range=None, include_tables=False, include_media=False):
    _, entries = load_package(source)
    report = inspect_entries(entries, include_text or paragraph_range is not None)
    if stories:
        selected = set(stories)
        missing = selected - report['stories'].keys()
        if missing:
            raise ValueError('Story not found: ' + ', '.join(sorted(missing)))
        report['stories'] = {name: value for name, value in report['stories'].items() if name in selected}
    if paragraph_range is not None:
        start, end = paragraph_range
        if start < 0 or end <= start:
            raise ValueError('Paragraph range must be START:END with 0 <= START < END')
        for story in report['stories'].values():
            story['paragraphs'] = [item for item in story.get('paragraphs', []) if start <= item['index'] < end]
    if include_tables:
        for name, story in report['stories'].items():
            root = parse_xml(entries[name])
            story['tables'] = []
            for index, table in enumerate(root.findall('.//w:tbl', NS)):
                rows = table.findall('./w:tr', NS)
                story['tables'].append({
                    'index': index,
                    'row_count': len(rows),
                    'column_counts': [len(row.findall('./w:tc', NS)) for row in rows],
                })
    if include_media:
        items = [
            {'part': name, 'bytes': len(data)}
            for name, data in sorted(entries.items())
            if name.startswith('word/media/') and not name.endswith('/')
        ]
        report['media'] = {'count': len(items), 'items': items}
    if summary:
        report.pop('parts', None)
    return report


def pdf_artifacts(pdf, output_dir):
    pdf = Path(pdf).resolve()
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    pdfinfo = shutil.which('pdfinfo')
    pdftoppm = shutil.which('pdftoppm')
    if not pdfinfo or not pdftoppm:
        return {'page_count': None, 'previews': [], 'creator': None, 'producer': None, 'warning': 'Poppler pdfinfo and pdftoppm are required for page previews'}
    info = subprocess.run([pdfinfo, str(pdf)], capture_output=True, timeout=30)
    info_stdout = decode_output(info.stdout)
    info_stderr = decode_output(info.stderr)
    if info.returncode != 0:
        raise RuntimeError(info_stderr.strip() or 'pdfinfo failed')
    match = re.search(r'^Pages:\s+(\d+)\s*$', info_stdout, re.MULTILINE)
    page_count = int(match.group(1)) if match else None
    prefix = output_dir / 'page'
    for stale in output_dir.glob('page-*.png'):
        stale.unlink()
    raster = subprocess.run([pdftoppm, '-png', '-r', '120', str(pdf), str(prefix)], capture_output=True, timeout=120)
    if raster.returncode != 0:
        raise RuntimeError(decode_output(raster.stderr).strip() or 'pdftoppm failed')
    previews = [str(path) for path in sorted(output_dir.glob('page-*.png'))]
    creator = re.search(r'^Creator:\s*(.+?)\s*$', info_stdout, re.MULTILINE)
    producer = re.search(r'^Producer:\s*(.+?)\s*$', info_stdout, re.MULTILINE)
    return {
        'page_count': page_count,
        'previews': previews,
        'creator': creator.group(1) if creator else None,
        'producer': producer.group(1) if producer else None,
        'warning': None,
    }


def _with_pdf_artifacts(report, output_dir):
    if report.get('status') != 'rendered' or not report.get('pdf'):
        return report
    report['pages'] = pdf_artifacts(report['pdf'], Path(output_dir) / 'pages')
    return report


def _registered_com_progid(progid):
    if os.name != 'nt':
        return False
    try:
        import winreg
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, progid):
            return True
    except OSError:
        return False


def _office_progids(renderer):
    if renderer == 'word':
        return ['Word.Application']
    # WPS exposes a separate automation interface. It avoids the Word.Application
    # compatibility registration hanging on some WPS installations.
    candidates = ['KWPS.Application', 'Word.Application']
    return [progid for progid in candidates if _registered_com_progid(progid)]


def _render_document(source, output_dir, renderer):
    output_dir = Path(output_dir).resolve()
    output_dir.mkdir(parents=True, exist_ok=True)
    source = Path(source).resolve()
    if renderer == 'word' and os.name != 'nt':
        raise RuntimeError('The requested Word-compatible renderer is unavailable on this platform')
    if renderer in ('auto', 'word') and os.name == 'nt':
        powershell = shutil.which('powershell') or shutil.which('pwsh')
        script = Path(__file__).with_name('render_docx.ps1')
        progids = _office_progids(renderer)
        if powershell and script.is_file() and progids:
            failures = []
            for progid in progids:
                try:
                    result = subprocess.run(
                        [powershell, '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', str(script), '-InputPath', str(source), '-OutputDirectory', str(output_dir), '-ApplicationProgId', progid],
                        capture_output=True, text=True, encoding='utf-8', timeout=60,
                    )
                except subprocess.TimeoutExpired:
                    failures.append(f'{progid}: timed out after 60 seconds')
                    continue
                if result.returncode == 0:
                    try:
                        return _with_pdf_artifacts(json.loads(result.stdout), output_dir)
                    except (json.JSONDecodeError, OSError, RuntimeError) as exc:
                        failures.append(f'{progid}: invalid renderer output: {exc}')
                        continue
                failures.append(f'{progid}: {result.stderr.strip() or result.stdout.strip() or "rendering failed"}')
            if renderer == 'word':
                raise RuntimeError('; '.join(failures) or 'Word-compatible renderer failed')
        elif renderer == 'word':
            raise RuntimeError('The requested Word-compatible renderer is unavailable')
    if renderer in ('auto', 'libreoffice'):
        soffice = shutil.which('soffice') or shutil.which('libreoffice')
        if soffice:
            result = subprocess.run(
                [soffice, '--headless', '--convert-to', 'pdf', '--outdir', str(output_dir), str(source)],
                capture_output=True, text=True, encoding='utf-8', timeout=120,
            )
            pdf = output_dir / (source.stem + '.pdf')
            if result.returncode == 0 and pdf.is_file():
                return _with_pdf_artifacts({'status': 'rendered', 'renderer': 'LibreOffice', 'pdf': str(pdf), 'cleanup_warning': None}, output_dir)
            if renderer == 'libreoffice':
                raise RuntimeError(result.stderr.strip() or result.stdout.strip() or 'LibreOffice rendering failed')
        elif renderer == 'libreoffice':
            raise RuntimeError('The requested LibreOffice renderer is unavailable')
    return {'status': 'unavailable', 'renderer': None, 'pdf': None, 'cleanup_warning': None}


def verify_document(source, render='auto', output_dir=None):
    if render not in {'never', 'auto', 'word', 'libreoffice'}:
        raise ValueError('render must be never, auto, word, or libreoffice')
    source = Path(source).resolve()
    structure = inspect_package(source)
    reopen = {'status': 'failed'}
    reopen_error = None
    try:
        from docx import Document
        document = Document(source)
        reopen = {
            'status': 'passed',
            'paragraphs': len(document.paragraphs),
            'tables': len(document.tables),
            'sections': len(document.sections),
            'inline_shapes': len(document.inline_shapes),
        }
    except Exception as exc:
        reopen_error = str(exc)
        reopen['error'] = reopen_error
    if render == 'never':
        rendered = {'status': 'skipped', 'renderer': None, 'pdf': None, 'cleanup_warning': None}
    else:
        destination = output_dir or tempfile.mkdtemp(prefix='docx-killer-render-')
        rendered = _render_document(source, destination, render)
    limitations = []
    if rendered['status'] != 'rendered':
        limitations.append('Layout was not rendered')
    else:
        limitations.append('PDF export does not automatically detect clipping, overlap, font substitution, or incorrect pagination; inspect page images')
    status = 'passed' if not structure['errors'] and reopen_error is None else 'failed'
    return {'status': status, 'structure': structure, 'reopen': reopen, 'render': rendered, 'limitations': limitations}


def parse_paragraph_range(value):
    try:
        start, end = value.split(':', 1)
        return int(start), int(end)
    except (ValueError, AttributeError) as exc:
        raise argparse.ArgumentTypeError('Use START:END, for example 0:20') from exc


def check_output(source, output):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output:
        raise ValueError('Use a separate output; in-place edits are not supported')
    if output.exists():
        raise FileExistsError('Output already exists: ' + str(output))
    if source.suffix.lower() != '.docx' or output.suffix.lower() != '.docx':
        raise ValueError('Only .docx input and output are supported')
    if not output.parent.is_dir():
        raise ValueError('Output directory must already exist')
    return output


def editable_package(source):
    infos, entries = load_package(source)
    if any(name.startswith('_xmlsignatures/') for name in entries):
        raise ValueError('Editing a signed package is not supported')
    if any('vbaproject' in name.lower() for name in entries) or b'macroEnabled' in entries.get('[Content_Types].xml', b''):
        raise ValueError('Macro-enabled packages are not supported')
    report = inspect_entries(entries)
    if report['errors']:
        raise ValueError('; '.join(report['errors']))
    return infos, entries


def publish(output, writer):
    fd, temp_name = tempfile.mkstemp(prefix='.docx-killer-', suffix='.docx', dir=output.parent)
    os.close(fd)
    temp = Path(temp_name)
    try:
        writer(temp)
        report = inspect_package(temp)
        if report['errors']:
            raise ValueError('Generated package failed inspection: ' + '; '.join(report['errors']))
        # Same-directory hard link publishes completely written bytes without overwriting.
        os.link(temp, output)
    finally:
        temp.unlink(missing_ok=True)


def field_paragraphs(root):
    protected, depth = set(), 0
    for node in root.iter():
        if node.tag == f'{{{W}}}fldChar' and node.get(f'{{{W}}}fldCharType') == 'begin':
            depth += 1
        if depth:
            paragraph = node if node.tag == f'{{{W}}}p' else next(node.iterancestors(f'{{{W}}}p'), None)
            if paragraph is not None:
                protected.add(paragraph)
        if node.tag == f'{{{W}}}fldChar' and node.get(f'{{{W}}}fldCharType') == 'end':
            depth = max(0, depth - 1)
    return protected


def plain_text_nodes(paragraph, protected):
    allowed_ancestors = {f'{{{W}}}{name}' for name in ('document', 'body', 'tbl', 'tr', 'tc', 'hdr', 'ftr', 'footnotes', 'footnote', 'endnotes', 'endnote', 'comments', 'comment')}
    if paragraph in protected or any(parent.tag not in allowed_ancestors for parent in paragraph.iterancestors()):
        raise ValueError('Matching paragraph is inside a protected structure')
    nodes = []
    for child in paragraph:
        if child.tag == f'{{{W}}}pPr':
            continue
        if child.tag != f'{{{W}}}r':
            raise ValueError('Matching paragraph contains protected markup')
        for node in child:
            if node.tag == f'{{{W}}}t':
                nodes.append(node)
            elif node.tag != f'{{{W}}}rPr':
                raise ValueError('Matching paragraph contains protected run content')
    if paragraph.findall('.//w:rPrChange', NS) or paragraph.findall('.//w:pPrChange', NS):
        raise ValueError('Matching paragraph contains protected formatting revisions')
    return nodes


def replace_text(source, output, old, new, expected, parts=None):
    output = check_output(source, output)
    if not old or old == new or expected < 1:
        raise ValueError('Require nonempty different text and a positive expected count')
    if any(character in old + new for character in '\r\n\t'):
        raise ValueError('This helper handles text only; paragraph breaks/tabs require structural editing')
    etree.Element('check').text = new
    infos, entries = editable_package(source)
    patterns = parts or ['word/document.xml']
    selected = [name for name in entries if any(fnmatch.fnmatchcase(name, pattern) for pattern in patterns)]
    if not selected:
        raise ValueError('No parts match the requested scope')
    changed, count = [], 0
    for name in selected:
        root = parse_xml(entries[name])
        if root.tag not in {f'{{{W}}}{kind}' for kind in ('document', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments')}:
            raise ValueError('Not a supported text story: ' + name)
        protected = field_paragraphs(root)
        dirty = False
        for index, paragraph in enumerate(root.findall('.//w:p', NS)):
            if old not in paragraph_text(paragraph):
                continue
            try:
                nodes = plain_text_nodes(paragraph, protected)
            except ValueError as exc:
                raise ValueError(f'{name}, paragraph {index}: {exc}') from exc
            text = ''.join(node.text or '' for node in nodes)
            matches = list(re.finditer(re.escape(old), text))
            spans, offset = [], 0
            for node in nodes:
                length = len(node.text or '')
                spans.append((offset, offset + length, node))
                offset += length
            for match in reversed(matches):
                first = True
                for start, end, node in spans:
                    if start >= match.end() or end <= match.start():
                        continue
                    value = node.text or ''
                    lo, hi = max(0, match.start() - start), min(end - start, match.end() - start)
                    node.text = value[:lo] + (new if first else '') + value[hi:]
                    node.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    first = False
            count += len(matches)
            dirty = dirty or bool(matches)
        if dirty:
            entries[name] = etree.tostring(root, encoding='UTF-8', xml_declaration=True)
            changed.append(name)
    if count != expected:
        raise ValueError(f'Replacement count mismatch: expected {expected}, found {count}')

    def write_package(temp):
        with ZipFile(temp, 'w') as archive:
            for info in infos:
                archive.writestr(info, entries[info.filename])

    publish(output, write_package)
    return {'replacements': count, 'changed_parts': sorted(changed), 'output': str(output), 'format_rule': 'Replacement inherits the first matched run; unaffected run properties remain', 'visual_check': 'not performed'}


def render_template(source, output, context):
    from docxtpl import DocxTemplate
    from jinja2 import Environment, StrictUndefined

    output = check_output(source, output)
    editable_package(source)
    if not isinstance(context, dict):
        raise ValueError('Template context must be a JSON object')
    template = DocxTemplate(source)
    environment = Environment(undefined=StrictUndefined, autoescape=True)
    template.render(context, jinja_env=environment, autoescape=True)
    publish(output, template.save)
    return {'output': str(output), 'undefined': 'strict', 'escaping': True, 'visual_check': 'not performed'}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    commands = parser.add_subparsers(dest='command', required=True)
    inspect = commands.add_parser('inspect', help='Inspect package structure; this is not full schema validation')
    inspect.add_argument('input')
    inspect.add_argument('--text', action='store_true', help='Include story text; may contain sensitive document content')
    inspect.add_argument('--summary', action='store_true', help='Omit the per-part hash inventory')
    inspect.add_argument('--story', action='append', help='Include one exact story part; repeatable')
    inspect.add_argument('--paragraph-range', type=parse_paragraph_range, metavar='START:END', help='Include only this half-open paragraph range')
    inspect.add_argument('--tables', action='store_true', help='Include table row and column counts')
    inspect.add_argument('--media', action='store_true', help='Include embedded media names and sizes')
    compare = commands.add_parser('compare', help='Compare uncompressed part payloads')
    compare.add_argument('before')
    compare.add_argument('after')
    replace = commands.add_parser('replace', help='Replace literal plain text across runs, refusing protected paragraphs')
    replace.add_argument('input')
    replace.add_argument('output')
    replace.add_argument('--old', required=True)
    replace.add_argument('--new', required=True)
    replace.add_argument('--expected', required=True, type=int)
    replace.add_argument('--part', action='append', help='Part name/glob; repeatable; default word/document.xml only')
    render = commands.add_parser('render', help='Render a trusted docxtpl template using JSON context')
    render.add_argument('input')
    render.add_argument('context')
    render.add_argument('output')
    verify = commands.add_parser('verify', help='Check package structure, python-docx reopening, and optional office rendering')
    verify.add_argument('input')
    verify.add_argument('--render', choices=('never', 'auto', 'word', 'libreoffice'), default='auto')
    verify.add_argument('--output-dir', help='Directory for rendered PDF; created when needed')
    args = parser.parse_args()
    try:
        if args.command == 'inspect':
            result = inspect_document(args.input, args.text, args.summary, args.story, args.paragraph_range, args.tables, args.media)
        elif args.command == 'compare':
            result = compare_packages(args.before, args.after)
        elif args.command == 'replace':
            result = replace_text(args.input, args.output, args.old, args.new, args.expected, args.part)
        elif args.command == 'render':
            context = json.loads(Path(args.context).read_text(encoding='utf-8-sig'))
            result = render_template(args.input, args.output, context)
        else:
            result = verify_document(args.input, args.render, args.output_dir)
        print(json.dumps(result, ensure_ascii=True, indent=2))
        return 1 if result.get('errors') or result.get('status') == 'failed' else 0
    except Exception as exc:
        print(json.dumps({'error': str(exc)}, ensure_ascii=True), file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
