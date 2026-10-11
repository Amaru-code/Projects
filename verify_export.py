#!/usr/bin/env python3
"""Verify a local Outlook COM export; never print actual email contents.

Pure standard-library Python 3.9+. No network, no Outlook access, no writes.
Usage: python scripts/verify_export.py path/to/InboxExport_... 
"""
from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
from pathlib import Path

EXPECTED_SCHEMA = 'LENWORK_ORIGINAL_INBOX_EXPORT_V1'
SHA256_RE = re.compile(r'^[0-9a-f]{64}$')


def _safe_path(root: Path, name: str) -> Path:
    path = Path(name)
    if path.is_absolute() or '..' in path.parts or '\\' in name:
        raise ValueError('Unsafe file path in manifest')
    resolved = (root / path).resolve()
    if not resolved.is_relative_to(root.resolve()):
        raise ValueError('File path escapes export directory')
    return resolved


def verify(root: Path) -> dict:
    root = root.resolve()
    manifest_file = root / 'manifest.json'
    if not manifest_file.is_file():
        raise ValueError('Missing manifest.json')
    manifest = json.loads(manifest_file.read_text(encoding='utf-8-sig'))
    if manifest.get('schema') != EXPECTED_SCHEMA:
        raise ValueError('Unexpected schema')
    mode = manifest.get('mode')
    if mode not in ('INBOX_METADATA', 'BODY_SUPPLEMENT'):
        raise ValueError('Unexpected export mode')
    target_status = 'COMPLETE_LOCAL_STORE' if mode == 'INBOX_METADATA' else 'COMPLETE_BODY_SUPPLEMENT'
    if manifest.get('status') != target_status or manifest.get('issue_codes'):
        raise ValueError('Export reported incomplete/errors')
    if manifest.get('source') != 'classic_outlook_local_com' or manifest.get('folder') != 'inbox':
        raise ValueError('Unexpected source/folder')
    if manifest.get('scanned_objects') != manifest.get('inbox_count_at_scan_start'):
        raise ValueError('Inbox scan count mismatch')
    if not isinstance(manifest.get('unique_records'), int) or manifest['unique_records'] < 0:
        raise ValueError('Invalid unique_records')
    start = dt.date.fromisoformat(manifest['scope_start'])
    end = dt.date.fromisoformat(manifest['scope_end_inclusive'])
    if start > end:
        raise ValueError('Invalid scope')
    files = manifest.get('files')
    if not isinstance(files, list):
        raise ValueError('Missing files list')
    expected_batch_count = manifest.get('batch_count')
    if not isinstance(expected_batch_count, int) or expected_batch_count < 0:
        raise ValueError('Invalid batch_count')
    expected_paths = {
        f'batches/mail_batch_{i:04d}.{ext}'
        for i in range(1, expected_batch_count + 1)
        for ext in ('jsonl', 'md')
    }
    if {f.get('path') for f in files} != expected_paths or len(files) != len(expected_paths):
        raise ValueError('Missing/extra/duplicated batch in manifest')
    ids: set[str] = set()
    records = 0
    observed_counts: dict[str, int] = {}
    for f in files:
        name = f['path']
        fmt = f['format']
        if fmt not in ('jsonl', 'md') or not name.endswith('.' + fmt):
            raise ValueError('File format mismatch')
        path = _safe_path(root, name)
        if not path.is_file():
            raise ValueError('Missing batch file')
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if not SHA256_RE.fullmatch(f.get('sha256', '')) or digest != f['sha256']:
            raise ValueError('SHA256 mismatch')
        if not isinstance(f.get('records'), int) or f['records'] < 1:
            raise ValueError('Invalid file record count')
        if fmt == 'jsonl':
            count = 0
            for line in path.read_text(encoding='utf-8').splitlines():
                if not line.strip():
                    continue
                obj = json.loads(line)
                ident = obj.get('message_id')
                if not isinstance(ident, str) or not ident:
                    raise ValueError('Missing message ID')
                if ident in ids:
                    raise ValueError('Duplicate message ID')
                ids.add(ident)
                if obj.get('folder') != 'inbox' or obj.get('source') != 'classic_outlook_local_com':
                    raise ValueError('Out-of-scope mail source')
                received = dt.datetime.fromisoformat(obj['received_at_local']).date()
                if not start <= received <= end:
                    raise ValueError('Out-of-scope mail date')
                if manifest.get('unread_only') and not obj.get('is_unread'):
                    raise ValueError('Unexpected read mail')
                if mode == 'INBOX_METADATA' and 'body' in obj:
                    raise ValueError('Full body present in metadata export')
                if mode == 'BODY_SUPPLEMENT' and 'body' not in obj:
                    raise ValueError('Missing body in supplement')
                count += 1
            records += count
            observed_counts[name] = count
            if count != f['records']:
                raise ValueError('Record count mismatch')
        else:
            # The .md files are human-readable mirrors; SHA binds them to the export.
            if path.read_text(encoding='utf-8').count('```json') != f['records']:
                raise ValueError('Markdown mirror record count mismatch')
    if records != manifest['unique_records']:
        raise ValueError('Total record count mismatch')
    if len(ids) != records:
        raise ValueError('Total ID count mismatch')
    if mode == 'BODY_SUPPLEMENT' and manifest['requested_body_ids_count'] != manifest['found_body_ids_count']:
        raise ValueError('Not all selected bodies exported')
    if records > expected_batch_count * manifest.get('batch_size', 1):
        raise ValueError('Batch size limit violated')
    return {
        'VERIFY_RESULT': 'PASS',
        'MODE': mode,
        'STATUS': manifest['status'],
        'UNIQUE_RECORDS': records,
        'BATCHES': expected_batch_count,
        'FILES_VERIFIED': len(files),
        'ISSUES': 0,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('export_directory', type=Path)
    args = parser.parse_args()
    try:
        result = verify(args.export_directory)
    except (OSError, ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        print('VERIFY_RESULT=FAIL')
        print('REASON=' + str(exc))
        raise SystemExit(1)
    for key, value in result.items():
        print(f'{key}={value}')


if __name__ == '__main__':
    main()
