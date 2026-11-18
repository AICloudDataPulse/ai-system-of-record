import io
import json
import os
import unittest
from contextlib import redirect_stdout
from unittest.mock import patch

from examples.live_followups import request_call, run


def completed_response(output):
    body = io.BytesIO(json.dumps({"status": "completed", "output": output}).encode())
    body.status = 200
    return body


class LiveExampleTests(unittest.TestCase):
    @patch("examples.live_followups.HTTPSConnection")
    def test_two_prompts_dispatch_locally(self, connect):
        connection = connect.return_value
        connection.getresponse.side_effect = [
            completed_response([{
                "type": "function_call", "name": "get_followups",
                "arguments": json.dumps({"date": "2026-11-14", **fields}),
            }]) for fields in ({}, {"include_earlier": True})
        ]
        stdout = io.StringIO()
        with patch.dict(os.environ, {"NETWORKING_WORKBOOK": "private.xlsx"}), redirect_stdout(stdout):
            run("test-model", "test-key")
            self.assertEqual(os.environ["NETWORKING_WORKBOOK"], "private.xlsx")
        results = [json.loads(line)["result"] for line in stdout.getvalue().splitlines()]
        self.assertEqual([result["total"] for result in results], [1, 2])
        self.assertEqual(connection.request.call_count, 2)
        for request, phrase in zip(connection.request.call_args_list, ("on", "through")):
            self.assertEqual(request.args, ("POST", "/v1/responses"))
            body = json.loads(request.kwargs["body"])
            self.assertEqual(set(body), {"model", "input", "tools", "store", "parallel_tool_calls", "tool_choice"})
            self.assertEqual(body["input"], f"Who is due {phrase} November 14, 2026?")
            self.assertEqual(body["tool_choice"], "auto")
            self.assertEqual(body["tools"][0]["name"], "get_followups")
            for private in ("private.xlsx", "Alex Example", "Sam Sample", "test-key"):
                self.assertNotIn(private, json.dumps(body))

    @patch("examples.live_followups.HTTPSConnection")
    def test_zero_or_multiple_calls(self, connect):
        for output in ([], [{"type": "function_call"}] * 2):
            connect.return_value.getresponse.return_value = completed_response(output)
            with self.assertRaisesRegex(ValueError, "exactly one"):
                request_call("test-model", "test-key", "prompt", {})

    @patch("examples.live_followups.request_call")
    def test_incorrect_interpretation_fails_acceptance(self, call):
        call.return_value = {
            "name": "get_followups",
            "arguments": '{"date":"2026-11-14","include_earlier":true}',
        }
        with self.assertRaisesRegex(ValueError, "did not match"):
            run("test-model", "test-key")
