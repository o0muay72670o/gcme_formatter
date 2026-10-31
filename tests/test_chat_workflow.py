import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import Workbook

from excel_forms.chat_workflow import prepare, apply_request, save_json
from excel_forms.cli import main


class ChatWorkflowTests(unittest.TestCase):
    def setUp(self):
        scratch = Path(__file__).resolve().parents[1]/'.test_tmp'
        scratch.mkdir(exist_ok=True)
        location = patch.object(tempfile, 'tempdir', str(scratch))
        location.start()
        self.addCleanup(location.stop)
        temp = tempfile.TemporaryDirectory(dir=scratch)
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.source = self.root/'form.xlsx'
        wb = Workbook()
        wb.active['A1'] = 'Inspect connection'
        wb.save(self.source)
        self.master = self.root/'Master.xlsx'
        wb = Workbook()
        wb.active['A1'] = 'PROJECT'
        wb.save(self.master)
        self.config = {
            'master_file': str(self.master), 'master_sheet': 'Sheet', 'header_rows': 1,
            'last_column': 42, 'column_width_cm': .44, 'body_font_size': 8,
            'margins_cm': dict(left=1.6, right=.6, top=1.9, bottom=1.9, header=.8, footer=.8),
        }
        self.jobs, self.output = self.root/'jobs', self.root/'output'
        self.prepared = prepare(self.source, self.config, self.jobs)
        self.request = Path(self.prepared['request'])

    def write_plan(self):
        request = json.loads(self.request.read_text(encoding='utf-8'))
        plan = {
            'title': 'Connection inspection', 'header_fields': [],
            'sections': [{'title': '', 'source_ids': [], 'columns': ['Inspection', 'YES', 'NO'],
                          'rows': [{'cells': [{'text': 'Inspect connection', 'source_ids': ['s0:A1']},
                                              {'text': '', 'source_ids': []}, {'text': '', 'source_ids': []}]}]}],
            'excluded': [], 'review_notes': [],
        }
        envelope = {'source_hash': request['source_hash'], 'author': 'Codex chat', 'plan': plan}
        save_json(self.request.with_name('plan.json'), envelope)
        return envelope

    def test_end_to_end_without_key_or_network(self):
        self.write_plan()
        with patch.dict('os.environ', {'OPENAI_API_KEY': '', 'OPENAI_MODEL': ''}), \
                patch('excel_forms.planner.urlopen', side_effect=AssertionError('Network must not be called')):
            result = apply_request(self.request, self.config, self.output)
        self.assertEqual(result['api_calls'], 0)
        self.assertTrue(Path(result['output']).is_file())
        self.assertEqual(result['coverage']['represented'], 1)

    def test_missing_plan_waits_without_producing_workbook(self):
        result = apply_request(self.request, self.config, self.output)
        self.assertEqual(result['status'], 'waiting_for_chat')
        self.assertFalse((self.output/self.source.name).exists())

    def test_changed_source_and_changed_settings_are_rejected(self):
        self.write_plan()
        altered = dict(self.config, body_font_size=9)
        with self.assertRaisesRegex(ValueError, 'changed since prepare'):
            apply_request(self.request, altered, self.output)
        wb = Workbook()
        wb.active['A1'] = 'Different source'
        wb.save(self.source)
        with self.assertRaisesRegex(ValueError, 'changed since prepare'):
            apply_request(self.request, self.config, self.output)

    def test_prepare_keeps_existing_plan(self):
        self.write_plan()
        before = self.request.with_name('plan.json').read_bytes()
        prepare(self.source, self.config, self.jobs)
        self.assertEqual(before, self.request.with_name('plan.json').read_bytes())

    def test_plan_cannot_claim_wrong_request(self):
        envelope = self.write_plan()
        envelope['source_hash'] = 'wrong'
        save_json(self.request.with_name('plan.json'), envelope)
        with self.assertRaisesRegex(ValueError, 'fingerprint'):
            apply_request(self.request, self.config, self.output)

    def test_batch_default_prepares_without_api(self):
        inputs = self.root/'input'
        inputs.mkdir()
        self.source.rename(inputs/self.source.name)
        config = self.root/'config.json'
        save_json(config, self.config)
        with patch('excel_forms.planner.urlopen', side_effect=AssertionError('Network must not be called')):
            status = main(['--config', str(config), 'batch', str(inputs), '--jobs', str(self.jobs)])
        self.assertEqual(status, 0)

    def test_unaccounted_and_malformed_plan_rejected(self):
        envelope = self.write_plan()
        envelope['plan']['sections'][0]['rows'][0]['cells'][0]['source_ids'] = []
        save_json(self.request.with_name('plan.json'), envelope)
        with self.assertRaisesRegex(ValueError, 'unaccounted'):
            apply_request(self.request, self.config, self.output)
        envelope['plan']['sections'][0]['columns'] = 'not an array'
        save_json(self.request.with_name('plan.json'), envelope)
        with self.assertRaisesRegex(ValueError, 'expected an array'):
            apply_request(self.request, self.config, self.output)

    def test_chat_title_fills_configured_blank_master_cell(self):
        from openpyxl import load_workbook
        config = dict(self.config, title_cell='B1')
        prepared = prepare(self.source, config, self.jobs)
        self.request = Path(prepared['request'])
        self.write_plan()
        result = apply_request(self.request, config, self.output)
        wb = load_workbook(result['output'])
        try:
            self.assertEqual(wb[result['new_sheet']]['B1'].value, 'Connection inspection')
        finally:
            wb.close()


if __name__ == '__main__':
    unittest.main()
