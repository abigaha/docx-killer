"""Package-level diagnostics, not a complete OOXML schema validator."""
import hashlib
import posixpath
from pathlib import Path
from urllib.parse import unquote, urlsplit
from zipfile import ZipFile

from lxml import etree

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
REL = 'http://schemas.openxmlformats.org/package/2006/relationships'
NS = {'w': W}
MAX_UNCOMPRESSED = 256 * 1024 * 1024
STORIES = ('document', 'hdr', 'ftr', 'footnotes', 'endnotes', 'comments')


def parse_xml(data):
    root = etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))
    if root.getroottree().docinfo.doctype:
        raise ValueError('DTD declarations are not supported')
    return root


def load_package(path):
    with ZipFile(path) as archive:
        infos = archive.infolist()
        if len({info.filename for info in infos}) != len(infos):
            raise ValueError('Duplicate ZIP entry names')
        if sum(info.file_size for info in infos) > MAX_UNCOMPRESSED:
            raise ValueError('Package exceeds the 256 MiB uncompressed inspection limit')
        entries = {}
        for info in infos:
            name = info.filename
            if name.startswith('/') or '\\' in name or '..' in name.split('/'):
                raise ValueError('Noncanonical ZIP entry path: ' + name)
            entries[name] = archive.read(info)
        return infos, entries


def relationship_target(rels_name, target):
    path = unquote(urlsplit(target).path)
    if path.startswith('/'):
        return posixpath.normpath(path.lstrip('/'))
    base = '' if rels_name == '_rels/.rels' else posixpath.dirname(posixpath.dirname(rels_name))
    return posixpath.normpath(posixpath.join(base, path))


def paragraph_text(paragraph):
    values = []
    for node in paragraph.iter():
        if node.tag not in {f'{{{W}}}t', f'{{{W}}}delText', f'{{{W}}}tab', f'{{{W}}}br'}:
            continue
        owner = next(node.iterancestors(f'{{{W}}}p'), None)
        if owner is paragraph:
            values.append('\t' if node.tag.endswith('}tab') else '\n' if node.tag.endswith('}br') else node.text or '')
    return ''.join(values)


def inspect_entries(entries, include_text=False):
    report = {'errors': [], 'warnings': [], 'parts': {}, 'stories': {}, 'external_relationships': []}
    roots = {}
    for name, data in entries.items():
        report['parts'][name] = {'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}
        if name.endswith(('.xml', '.rels')):
            try:
                roots[name] = parse_xml(data)
            except (etree.XMLSyntaxError, ValueError) as exc:
                report['errors'].append(f'{name}: {exc}')
    for required in ('[Content_Types].xml', '_rels/.rels', 'word/document.xml'):
        if required not in entries:
            report['errors'].append('Required part missing: ' + required)
    main = roots.get('word/document.xml')
    if main is not None and main.tag != f'{{{W}}}document':
        report['errors'].append('Only Transitional OOXML with word/document.xml is supported; convert Strict OOXML with an office application first')
    features = {'fields': 0, 'revisions': 0, 'content_controls': 0, 'drawings': 0, 'text_boxes': 0, 'bookmarks': 0, 'comments': 0, 'equations': 0}
    feature_stories = {key: set() for key in features}
    names = {'fldChar': 'fields', 'fldSimple': 'fields', 'ins': 'revisions', 'del': 'revisions', 'moveFrom': 'revisions', 'moveTo': 'revisions', 'sdt': 'content_controls', 'drawing': 'drawings', 'pict': 'drawings', 'txbxContent': 'text_boxes', 'bookmarkStart': 'bookmarks', 'comment': 'comments'}
    for name, root in roots.items():
        if name.endswith('.rels'):
            seen = set()
            for rel in root.findall(f'{{{REL}}}Relationship'):
                rid, target = rel.get('Id'), rel.get('Target', '')
                if not rid or rid in seen:
                    report['errors'].append(f'{name}: missing or duplicate relationship Id {rid}')
                seen.add(rid)
                if rel.get('TargetMode') == 'External':
                    report['external_relationships'].append({'part': name, 'id': rid, 'target': target})
                elif not target or relationship_target(name, target) not in entries:
                    report['errors'].append(f'{name}: missing relationship target {target}')
        if root.tag in {f'{{{W}}}{kind}' for kind in STORIES}:
            paragraphs = root.findall('.//w:p', NS)
            story = {'paragraph_count': len(paragraphs), 'table_count': len(root.findall('.//w:tbl', NS))}
            if include_text:
                story['paragraphs'] = [{'index': index, 'text': paragraph_text(p)} for index, p in enumerate(paragraphs)]
            report['stories'][name] = story
            for node in root.iter():
                if not isinstance(node.tag, str):
                    continue
                local = etree.QName(node).localname
                if etree.QName(node).namespace == W and local in names:
                    feature = names[local]
                    features[feature] += 1
                    feature_stories[feature].add(name)
                elif local == 'oMath':
                    features['equations'] += 1
                    feature_stories['equations'].add(name)
    report['features'] = features
    report['feature_locations'] = {key: sorted(value) for key, value in feature_stories.items() if value}
    if any(name.startswith('_xmlsignatures/') for name in entries):
        report['warnings'].append('Digitally signed package: edits invalidate signatures')
    if any('vbaproject' in name.lower() for name in entries):
        report['warnings'].append('Macro payload present')
    if any(features.values()):
        details = []
        for feature, count in features.items():
            if count:
                locations = ', '.join(sorted(feature_stories[feature]))
                details.append(f'{feature}={count} ({locations})')
        report['warnings'].append('Advanced structures present: ' + '; '.join(details))
    report['validation_scope'] = 'ZIP reads/CRC, XML well-formedness, conventional DOCX parts, relationship target existence; not XSD validation, rendering, or field correctness'
    return report


def inspect_package(path, include_text=False):
    _, entries = load_package(path)
    return inspect_entries(entries, include_text)


def compare_packages(before, after):
    _, old = load_package(before)
    _, new = load_package(after)
    return {
        'added': sorted(new.keys() - old.keys()),
        'removed': sorted(old.keys() - new.keys()),
        'changed': sorted(name for name in old.keys() & new.keys() if old[name] != new[name]),
        'unchanged_count': sum(old[name] == new[name] for name in old.keys() & new.keys()),
        'comparison_scope': 'Uncompressed part bytes, not ZIP metadata or visual layout',
    }
