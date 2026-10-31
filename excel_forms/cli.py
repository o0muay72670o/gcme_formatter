import argparse
from concurrent.futures import ThreadPoolExecutor
import json
from pathlib import Path
import sys
import time

from .extract import extract, master_context
from .planner import cache_key, create_plan, load_env, validate_plan
from .writer import write_workbook


def save_json(path, data):
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')


def main(argv=None):
    parser = argparse.ArgumentParser(description='Engineering Excel forms: local extraction, AI mapping, deterministic generation.')
    parser.add_argument('--config', default='config.json')
    sub = parser.add_subparsers(dest='command', required=True)
    for name in ['inspect', 'plan', 'convert', 'batch']:
        command = sub.add_parser(name)
        command.add_argument('path', type=Path, help='Source XLSX, or directory for batch')
        command.add_argument('--sheets', nargs='+', help='Optional exact source sheet names for extraction')
        command.add_argument('--output', type=Path, default=Path('output'))
        command.add_argument('--cache', type=Path, default=Path('cache'))
        if name == 'convert':
            command.add_argument('--plan', type=Path, help='Use a saved plan envelope instead of calling AI')
        if name == 'batch':
            command.add_argument('--workers', type=int, default=1, choices=range(1, 5), help='Concurrent files, 1-4')
    args = parser.parse_args(argv)
    load_env()
    try:
        config_path = Path(args.config).resolve()
        config = json.loads(config_path.read_text(encoding='utf-8-sig'))
        config['master_file'] = str((config_path.parent / config['master_file']).resolve())
        if config['last_column'] < 21 or config['last_column'] > 100:
            raise ValueError('last_column must be between 21 and 100')
        if not 0 < config['header_rows'] <= 100 or config['body_font_size'] <= 0 or config['column_width_cm'] <= 0:
            raise ValueError('Header row count, font size and column width must be positive')
        if args.command == 'batch':
            if not args.path.is_dir():
                raise ValueError('batch requires a directory')
            sources = sorted(p for p in args.path.glob('*.xlsx') if not p.name.startswith('~$')
                             and p.resolve() != Path(config['master_file']).resolve())
            if not sources:
                raise ValueError('No input XLSX files found')
        else:
            if not args.path.is_file() or args.path.suffix.lower() != '.xlsx':
                raise ValueError('Input must be an existing .xlsx file')
            sources = [args.path]

        def process(source):
            start = time.monotonic()
            try:
                output = args.output / source.name
                if args.command in ('convert', 'batch') and output.exists():
                    return {'source': str(source), 'status': 'skipped_existing_output'}
                extracted = extract(source, args.sheets)
                if args.command == 'inspect':
                    target = args.output/'reports'/f'{source.stem}.extracted.json'
                    save_json(target, extracted)
                    return {'source': str(source), 'status': 'inspected', 'report': str(target.resolve()),
                            'sheets': len(extracted['sheets']), 'cells': sum(len(s['cells']) for s in extracted['sheets'])}
                scoped = dict(config, source_sheets=args.sheets)
                if getattr(args, 'plan', None):
                    envelope = json.loads(args.plan.read_text(encoding='utf-8'))
                    if envelope.get('source_hash') != cache_key(source, scoped):
                        raise ValueError('Saved plan does not match this source, Master or configuration')
                    from_cache = True
                else:
                    envelope, from_cache = create_plan(source, extracted, scoped, args.cache)
                context = master_context(Path(config['master_file']), config['master_sheet'], config['header_rows'], config['last_column'])
                coverage = validate_plan(envelope['plan'], extracted, context)
                plan_path = args.output/'reports'/f'{source.stem}.plan.json'
                save_json(plan_path, envelope)
                result = {'source': str(source), 'status': 'planned', 'cache_hit': from_cache,
                          'usage': envelope.get('usage', {}), 'coverage': coverage, 'plan': str(plan_path.resolve())}
                if args.command != 'plan':
                    result.update(write_workbook(source, output, envelope['plan'], extracted, config))
                result['seconds'] = round(time.monotonic()-start, 2)
                save_json(args.output/'reports'/f'{source.stem}.report.json', result)
                return result
            except Exception as error:
                result = {'source': str(source), 'status': 'failed', 'error': str(error)}
                save_json(args.output/'reports'/f'{source.stem}.error.json', result)
                return result

        if args.command == 'batch':
            with ThreadPoolExecutor(max_workers=args.workers) as pool:
                results = list(pool.map(process, sources))
            save_json(args.output/'batch-report.json', results)
        else:
            results = [process(sources[0])]
        for result in results:
            print(json.dumps(result, ensure_ascii=True, indent=2))
        return 1 if any(r['status'] == 'failed' for r in results) else 0
    except (ValueError, OSError, KeyError) as error:
        print(f'Error: {error}', file=sys.stderr)
        return 1
