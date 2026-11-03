"""Regression cases for the reviewer's Body identification rule."""
import copy
import unittest

from excel_forms.planner import validate_plan, INSTRUCTIONS


class IdentificationTests(unittest.TestCase):
    def setUp(self):
        self.plan = dict(title='Inspection', header_fields=[], excluded=[], review_notes=[],
                         sections=[dict(title='Checks', source_ids=[], columns=['Inspection'],
                                        rows=[dict(cells=[dict(text='Inspect equipment', source_ids=['s0:A1'])])])])
        self.source = {'sheets': [{'cells': [{'id': 's0:A1', 'value': 'Inspect equipment'}]}]}
        self.context = {'allowed_targets': []}

    def test_duplicate_aliases_rejected_everywhere(self):
        for text in ['GCME Contract No. :', 'SERVICE', 'Service No', 'Contractor',
                     'EQUIPMENT NO:', 'Equipment Number: P-100']:
            for location in ['title', 'column', 'cell']:
                with self.subTest(text=text, location=location):
                    plan = copy.deepcopy(self.plan)
                    section = plan['sections'][0]
                    if location == 'title':
                        section['title'] = text
                    elif location == 'column':
                        section['columns'][0] = text
                    else:
                        section['rows'][0]['cells'][0]['text'] = text
                    with self.assertRaisesRegex(ValueError, 'Duplicate Body identification'):
                        validate_plan(plan, self.source, self.context)

    def test_completion_columns_and_role_rows_preserved(self):
        section = self.plan['sections'][0]
        section['columns'] = ['Contractor', 'GCME', 'Client (if applicable)']
        section['rows'][0]['cells'] += [dict(text='', source_ids=[]), dict(text='', source_ids=[])]
        validate_plan(self.plan, self.source, self.context)
        section['columns'] = ['Company', 'Company', 'Company']
        for cell, text in zip(section['rows'][0]['cells'], ['Contractor', 'GCME', 'Client']):
            cell['text'] = text
        validate_plan(self.plan, self.source, self.context)

    def test_engineering_and_acceptance_wording_preserved(self):
        for text in ['Electrical & Building Services', 'Ancillary equipment supplied',
                     'ACCEPTED Contractor', 'Building Contractor', 'Contractor Date/Sign']:
            self.plan['sections'][0]['rows'][0]['cells'][0]['text'] = text
            validate_plan(self.plan, self.source, self.context)

    def test_excluded_duplicate_still_requires_source_coverage(self):
        self.plan['sections'][0]['rows'][0]['cells'][0]['source_ids'] = []
        with self.assertRaisesRegex(ValueError, 'unaccounted'):
            validate_plan(self.plan, self.source, self.context)
        self.plan['excluded'] = [{'source_id': 's0:A1', 'reason': 'Duplicate identification in Master header'}]
        validate_plan(self.plan, self.source, self.context)

    def test_instructions_include_review_and_data_preservation(self):
        self.assertIn('Never recreate identification fields', INSTRUCTIONS)
        self.assertIn('never silently discard a value', INSTRUCTIONS)


if __name__ == '__main__':
    unittest.main()
