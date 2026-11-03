from pathlib import Path
import json
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook, load_workbook

from excel_forms.cli import main
from excel_forms.extract import extract, master_context
from excel_forms.planner import cache_key, create_plan, validate_plan
from excel_forms.writer import unique_name, write_workbook


class PipelineTests(unittest.TestCase):
    def setUp(self):
        scratch = Path(__file__).resolve().parents[1]/'.test_tmp'
        scratch.mkdir(exist_ok=True)
        location = patch.object(tempfile, 'tempdir', str(scratch))
        location.start()
        self.addCleanup(location.stop)
        self.temp = tempfile.TemporaryDirectory(dir=scratch)
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        self.source = self.root/'Example engineering form with a long name.xlsx'
        wb = Workbook()
        wb.active.title = 'Original'
        wb.active['A1'], wb.active['B1'] = 'Project name', 'Example'
        wb.active['A2'] = '1. Inspect connection'
        wb.save(self.source)
        self.master = self.root/'Master.xlsx'
        wb = Workbook()
        wb.active['A1'] = 'PROJECT'
        wb.active.merge_cells('B1:D1')
        wb.save(self.master)
        self.config = {
            'master_file': str(self.master), 'master_sheet': 'Sheet', 'header_rows': 1,
            'last_column': 42, 'column_width_cm': .44, 'body_font_size': 8,
            'margins_cm': dict(left=1.6, right=.6, top=1.9, bottom=1.9, header=.8, footer=.8),
            'max_input_characters': 100000, 'max_output_tokens': 16000,
        }
        self.plan = {
            'title': 'Example',
            'header_fields': [{'target': 'B1', 'source_ids': ['s0:B1']}],
            'sections': [{'title': '', 'source_ids': [], 'columns': ['Inspection', 'YES', 'NO'],
                          'rows': [{'cells': [{'text': '1. Inspect connection', 'source_ids': ['s0:A2']},
                                             {'text': '', 'source_ids': []}, {'text': '', 'source_ids': []}]}]}],
            'excluded': [{'source_id': 's0:A1', 'reason': 'Label represented in Master header'}],
            'review_notes': [],
        }
        self.extracted = extract(self.source)
        self.context = master_context(self.master, 'Sheet', 1, 42)

    def test_rejects_unaccounted_source_and_header_label_overwrite(self):
        validate_plan(self.plan, self.extracted, self.context)
        self.plan['header_fields'][0]['target'] = 'A1'
        with self.assertRaisesRegex(ValueError, 'header target'):
            validate_plan(self.plan, self.extracted, self.context)
        self.plan['header_fields'][0]['target'] = 'B1'
        self.plan['excluded'] = []
        with self.assertRaisesRegex(ValueError, 'unaccounted'):
            validate_plan(self.plan, self.extracted, self.context)

    def test_saves_original_and_master_layout_and_print_settings(self):
        original = self.source.read_bytes()
        destination = self.root/'out.xlsx'
        result = write_workbook(self.source, destination, self.plan, self.extracted, self.config)
        wb = load_workbook(destination)
        self.addCleanup(wb.close)
        self.assertEqual(wb['Old']['A2'].value, '1. Inspect connection')
        new = wb[result['new_sheet']]
        self.assertTrue(new.sheet_view.tabSelected)
        self.assertTrue(all(not sheet.sheet_view.tabSelected for sheet in wb if sheet is not new))
        self.assertIn('B1:D1', {str(r) for r in new.merged_cells.ranges})
        self.assertEqual(new['A1'].value, 'PROJECT')
        self.assertEqual(new['B1'].value, 'Example')
        self.assertEqual(int(new.page_setup.paperSize), 9)
        self.assertEqual(new.page_setup.fitToHeight, 1)
        self.assertAlmostEqual(new.page_margins.left*2.54, 1.6)
        self.assertEqual(original, self.source.read_bytes())
        with self.assertRaises(FileExistsError):
            write_workbook(self.source, destination, self.plan, self.extracted, self.config)

    def test_rename_rejects_dependent_formulas(self):
        wb = load_workbook(self.source)
        wb.create_sheet('Other')['A1'] = "='Original'!B1"
        wb.save(self.source)
        with self.assertRaisesRegex(ValueError, 'dependents'):
            write_workbook(self.source, self.root/'out.xlsx', self.plan, self.extracted, self.config)

    def test_already_old_sheet_keeps_dependent_references(self):
        wb = load_workbook(self.source)
        wb.active.title = 'Old'
        wb.create_sheet('Other')['A1'] = "='Old'!B1"
        wb.save(self.source)
        result = write_workbook(self.source, self.root/'out.xlsx', self.plan, self.extracted, self.config)
        wb = load_workbook(result['output'])
        try:
            self.assertEqual(wb['Other']['A1'].value, "='Old'!B1")
            self.assertIn('Old', wb.sheetnames)
        finally:
            wb.close()

    def test_cache_avoids_api_and_is_invalidated_by_source_change(self):
        cache = self.root/'cache'
        cache.mkdir()
        import hashlib
        model = 'test-model'
        key = hashlib.sha256((cache_key(self.source, self.config)+model).encode()).hexdigest()
        (cache/f'{key}.json').write_text(json.dumps({'plan': self.plan}), encoding='utf-8')
        with patch.dict('os.environ', {'OPENAI_MODEL': model, 'OPENAI_API_KEY': ''}), patch('excel_forms.planner.urlopen') as call:
            _, hit = create_plan(self.source, self.extracted, self.config, cache)
            self.assertTrue(hit)
            call.assert_not_called()
            wb = load_workbook(self.source)
            wb.active['A3'] = 'Changed'
            wb.save(self.source)
            with self.assertRaisesRegex(ValueError, 'OPENAI_API_KEY'):
                create_plan(self.source, extract(self.source), self.config, cache)

    def test_incomplete_ai_response_does_not_create_plan(self):
        from io import BytesIO
        response = BytesIO(json.dumps({'status': 'incomplete'}).encode())
        with patch.dict('os.environ', {'OPENAI_MODEL': 'test', 'OPENAI_API_KEY': 'fake-test-key'}), \
                patch('excel_forms.planner.urlopen', return_value=response):
            with self.assertRaisesRegex(RuntimeError, 'did not complete'):
                create_plan(self.source, self.extracted, self.config, self.root/'cache')

    def test_inspect_cli_has_no_ai_call(self):
        conf = self.root/'config.json'
        conf.write_text(json.dumps(self.config), encoding='utf-8')
        with patch('excel_forms.planner.urlopen') as call:
            status = main(['--config', str(conf), 'inspect', str(self.source), '--output', str(self.root/'output')])
        self.assertEqual(status, 0)
        call.assert_not_called()

    def test_unique_name_handles_long_names_and_collisions(self):
        self.assertLessEqual(len(unique_name('a'*40, ['a'*31])), 31)
        self.assertEqual(unique_name('Old', ['old']), 'Old_2')

    def test_eight_column_acceptance_table_keeps_all_columns(self):
        section = self.plan['sections'][0]
        section['columns'] = ['Item', 'YES', 'N/A', 'NO', 'Date', 'Contractor', 'GCME', 'Client']
        section['rows'][0]['cells'] = [section['rows'][0]['cells'][0]] + [{'text':'','source_ids':[]} for _ in range(7)]
        validate_plan(self.plan, self.extracted, self.context)
        target = self.root/'eight-columns.xlsx'
        result = write_workbook(self.source, target, self.plan, self.extracted, self.config)
        wb = load_workbook(target)
        try:
            ws = wb[result['new_sheet']]
            headings = [c.value for row in ws.iter_rows(min_row=2, max_row=2) for c in row if c.value]
            self.assertEqual(headings, section['columns'])
        finally:
            wb.close()


if __name__ == '__main__':
    unittest.main()
