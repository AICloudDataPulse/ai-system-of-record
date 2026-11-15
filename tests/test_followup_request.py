import json
import unittest
from datetime import date
from io import StringIO
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from openpyxl import Workbook
from networking.followups import handle_followup_request


class FollowupRequestTests(unittest.TestCase):
    def test_invalid_shapes_rejected_before_delegation(self):
        cases = [
            None, True, 1, [], [('date', '2026-11-14')],
            '{"date":"2026-11-14"}', {}, {'include_earlier': True},
            {'date': '2026-11-14', 'workbook_path': '/private/path.xlsx'},
            {'date': '2026-11-14', 'unexpected': 'private value'},
            {'date': '2026-11-14', 1: True},
        ]
        for request in cases:
            with self.subTest(request=request), patch(
                'networking.followups.get_followups_response'
            ) as wrapper:
                self.assertEqual(handle_followup_request(request), {
                    'ok': False, 'error': {
                        'code': 'INVALID_INPUT',
                        'message': 'Provide a request dictionary with date and optional include_earlier only.',
                    },
                })
                wrapper.assert_not_called()

    def test_invalid_fields_rejected_before_configuration_or_access(self):
        cases = [{'date': value} for value in (
            None, True, 123, [], {}, date(2026, 11, 14),
            '', 'private-invalid-date', '2026-02-30', '20261114',
        )] + [{'date': '2026-11-14', 'include_earlier': value}
             for value in (None, 0, 1, 'true', [], {})]
        for request in cases:
            with (
                self.subTest(request=request),
                patch('networking.followups.get_workbook_path') as config,
                patch('networking.followups.load_workbook') as loader,
            ):
                result = handle_followup_request(request)
                self.assertFalse(result['ok'])
                self.assertEqual(result['error']['code'], 'INVALID_INPUT')
                self.assertNotIn('private-invalid-date', json.dumps(result))
                config.assert_not_called()
                loader.assert_not_called()

    def test_delegates_without_mutating_request_or_response(self):
        for fields, flag in (({}, False), ({'include_earlier': False}, False),
                             ({'include_earlier': True}, True)):
            request = {'date': ' 2026-11-14 ', **fields}
            original = dict(request)
            for response in ({'ok': True}, {'ok': False, 'error': {'code': 'CONFIGURATION_ERROR'}}):
                with patch('networking.followups.get_followups_response', return_value=response) as wrapper:
                    self.assertIs(handle_followup_request(request), response)
                    wrapper.assert_called_once_with(' 2026-11-14 ', include_earlier=flag)
                    self.assertEqual(request, original)

    def test_real_workbook_results_and_no_console_io(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'synthetic.xlsx'
            workbook = Workbook()
            try:
                sheet = workbook.active
                sheet.title = 'Contacts'
                sheet.append(['Name', 'Company', 'Next Action', 'Follow-up Due'])
                sheet.append(['Alex Example', None, 'Send update', date(2026, 11, 14)])
                sheet.append(['Sam Sample', 'Sample Works', 'Call', date(2026, 11, 13)])
                workbook.save(path)
            finally:
                workbook.close()
            original = path.read_bytes()
            for value, fields, count, mode in (
                (' 2026-11-14 ', {}, 1, 'exact'),
                ('2026-11-14', {'include_earlier': False}, 1, 'exact'),
                ('2026-11-14', {'include_earlier': True}, 2, 'due_through'),
                ('2024-02-29', {}, 0, 'exact'),
            ):
                with (
                    self.subTest(value=value, fields=fields),
                    patch.dict('os.environ', {'NETWORKING_WORKBOOK': str(path)}, clear=True),
                    patch('builtins.input') as prompt,
                    patch('sys.stdout', new_callable=StringIO) as stdout,
                    patch('sys.stderr', new_callable=StringIO) as stderr,
                ):
                    result = handle_followup_request({'date': value, **fields})
                    self.assertEqual(set(result), {'ok', 'query_date', 'mode', 'contacts', 'total'})
                    self.assertTrue(result['ok'])
                    self.assertEqual(result['query_date'], value.strip())
                    self.assertEqual(result['mode'], mode)
                    self.assertEqual(result['total'], count)
                    self.assertEqual([c['name'] for c in result['contacts']],
                                     ['Alex Example', 'Sam Sample'][:count])
                    self.assertEqual(json.loads(json.dumps(result)), result)
                    self.assertEqual(path.read_bytes(), original)
                    prompt.assert_not_called()
                    self.assertEqual(stdout.getvalue(), '')
                    self.assertEqual(stderr.getvalue(), '')

    def test_configuration_failure_and_unexpected_errors(self):
        with patch.dict('os.environ', {}, clear=True):
            result = handle_followup_request({'date': '2026-11-14'})
            self.assertEqual(result['error']['code'], 'CONFIGURATION_ERROR')
            self.assertFalse(result['ok'])
        for error in (TypeError('bug'), ValueError('bug'), RuntimeError('bug')):
            with patch('networking.followups.get_followups_response', side_effect=error):
                with self.assertRaises(type(error)) as raised:
                    handle_followup_request({'date': '2026-11-14'})
                self.assertIs(raised.exception, error)


if __name__ == '__main__':
    unittest.main()
