"""Compact console output with an unchanged, complete JSON artifact."""
import json
from pathlib import Path


def add_summary_argument(parser):
    parser.add_argument('--summary', action='store_true',
                        help='Print a compact summary; requires --json for full details.')


def require_summary_artifact(parser, args):
    if args.summary and not args.json:
        parser.error('--summary requires --json so full details are preserved')


def summarize(result, json_path):
    # Retain every diagnostic and verification boundary. Only bulky successful
    # records (source hashes, paragraph formatting, page text) leave stdout.
    keys = ('status', 'scope', 'issues', 'selected_references', 'environment',
            'input_signature', 'execution_verified', 'counts', 'page_count',
            'paragraph_count', 'table_count', 'figure_caption_count',
            'pending_manual', 'rendered_status', 'complete_verification',
            'visual_review_required', 'evidence_review_required',
            'rendered_review_required', 'manual_review_required')
    summary = {key: result[key] for key in keys if key in result}
    summary['details'] = str(Path(json_path).resolve())
    if 'fingerprints' in result:
        summary['fingerprint_count'] = len(result['fingerprints'])
    if 'conclusions' in result:
        summary['conclusion_count'] = len(result['conclusions'])
    if 'pages' in result and 'page_count' not in result:
        summary['page_count'] = len(result['pages'])
    if 'cache' in result:
        cache = result['cache']
        summary['cache'] = {key: value for key, value in cache.items()
                            if key not in ('added', 'changed', 'removed')}
        for key in ('added', 'changed', 'removed'):
            paths = cache.get(key, [])
            summary['cache'][key + '_count'] = len(paths)
            summary['cache'][key] = paths[:20]
        # All paths, including those beyond this preview, are in details.
    return summary


def emit_result(result, json_path=None, summary=False, print_saved=True):
    if summary and json_path is None:
        raise ValueError('A full JSON artifact is required for summary output')
    payload = json.dumps(result, ensure_ascii=False, indent=2, default=dict)
    if json_path:
        json_path = Path(json_path)
        json_path.parent.mkdir(parents=True, exist_ok=True)
        json_path.write_text(payload, encoding='utf-8')
    if summary:
        print(json.dumps(summarize(result, json_path), ensure_ascii=False))
    elif print_saved or not json_path:
        print(payload)
