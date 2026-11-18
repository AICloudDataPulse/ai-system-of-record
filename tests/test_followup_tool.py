import json
import unittest
from pathlib import Path
from unittest.mock import patch

from networking.followups import dispatch_tool_call, handle_followup_json


class FollowupToolTests(unittest.TestCase):
    def test_published_argument_contract(self):
        path = Path(__file__).resolve().parents[1] / "networking" / "get_followups.tool.json"
        tool = json.loads(path.read_text(encoding="utf-8"))
        self.assertEqual(tool["name"], "get_followups")
        self.assertTrue(tool["description"].strip())

        schema = tool["input_schema"]
        self.assertEqual(schema["type"], "object")
        self.assertEqual(schema["required"], ["date"])
        self.assertIs(schema["additionalProperties"], False)
        fields = schema["properties"]
        self.assertEqual(set(fields), {"date", "include_earlier"})
        self.assertEqual(fields["date"]["type"], "string")
        self.assertEqual(fields["date"]["format"], "date")
        self.assertEqual(fields["include_earlier"]["type"], "boolean")
        self.assertIs(fields["include_earlier"]["default"], False)

    def test_unknown_tool_fails_before_argument_handling(self):
        expected = {"ok": False, "error": {"code": "UNKNOWN_TOOL", "message": "Unknown tool name."}}
        with patch("networking.followups.handle_followup_json") as handler:
            for name in ("", "GET_FOLLOWUPS", "get_followups ", "other", None):
                with self.subTest(name=name):
                    self.assertEqual(dispatch_tool_call(name, "{"), expected)
            handler.assert_not_called()

    def test_known_tool_preserves_invalid_input_response(self):
        for arguments in ("{", '{"date":"2026-02-30"}', '{"workbook_path":"private"}'):
            with self.subTest(arguments=arguments):
                result = dispatch_tool_call("get_followups", arguments)
                self.assertEqual(result, handle_followup_json(arguments))
                self.assertEqual(result["error"]["code"], "INVALID_INPUT")
