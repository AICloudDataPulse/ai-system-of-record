import json
import unittest
from pathlib import Path


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
