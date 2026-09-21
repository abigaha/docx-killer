"""Inspect DOCX, replace plain text conservatively, and render trusted templates."""
import argparse
import fnmatch
import json
import os
import re
import sys
import tempfile
from pathlib import Path
from zipfile import ZipFile

from lxml import etree

from docx_package import W, NS, compare_packages, inspect_entries, inspect_package, load_package, paragraph_text, parse_xml


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
    args = parser.parse_args()
    try:
        if args.command == 'inspect':
            result = inspect_package(args.input, args.text)
        elif args.command == 'compare':
            result = compare_packages(args.before, args.after)
        elif args.command == 'replace':
            result = replace_text(args.input, args.output, args.old, args.new, args.expected, args.part)
        else:
            context = json.loads(Path(args.context).read_text(encoding='utf-8-sig'))
            result = render_template(args.input, args.output, context)
        print(json.dumps(result, ensure_ascii=True, indent=2))
        return 1 if result.get('errors') else 0
    except Exception as exc:
        print(json.dumps({'error': str(exc)}, ensure_ascii=True), file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
