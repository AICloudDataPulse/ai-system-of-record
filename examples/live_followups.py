"""Run two live tool-selection checks against a temporary fictional workbook."""

import argparse
import json
import os
from http.client import HTTPException, HTTPSConnection
from pathlib import Path

from examples.followup_response import fictional_workbook
from networking.followups import dispatch_tool_call


SCENARIOS = (
    ("Who is due on November 14, 2026?", "exact", ["Alex Example"]),
    ("Who is due through November 14, 2026?", "due_through", ["Alex Example", "Sam Sample"]),
)


def request_call(model, key, prompt, tool):
    payload = {
        "model": model,
        "input": prompt,
        "tools": [tool],
        "tool_choice": "auto",
        "parallel_tool_calls": False,
        "store": False,
    }
    connection = HTTPSConnection("api.openai.com", timeout=45)
    try:
        connection.request(
            "POST", "/v1/responses", body=json.dumps(payload).encode("utf-8"),
            headers={"Authorization": f"Bearer {key}", "Content-Type": "application/json"},
        )
        response = connection.getresponse()
        if response.status != 200:
            raise ValueError(f"OpenAI request failed (HTTP {response.status}).")
        data = json.load(response)
    finally:
        connection.close()
    if not isinstance(data, dict) or data.get("status") != "completed":
        raise ValueError("OpenAI did not return a completed response.")
    output = data.get("output")
    if not isinstance(output, list):
        raise ValueError("OpenAI response is missing its output list.")
    calls = [item for item in output if isinstance(item, dict) and item.get("type") == "function_call"]
    if len(calls) != 1:
        raise ValueError("Expected exactly one function call.")
    return calls[0]


def run(model, key):
    path = Path(__file__).resolve().parents[1] / "networking" / "get_followups.tool.json"
    definition = json.loads(path.read_text(encoding="utf-8"))
    tool = {
        "type": "function", "name": definition["name"], "description": definition["description"],
        "parameters": definition["input_schema"], "strict": False,
    }
    with fictional_workbook():
        for prompt, mode, names in SCENARIOS:
            call = request_call(model, key, prompt, tool)
            result = dispatch_tool_call(call.get("name"), call.get("arguments"))
            if not result["ok"]:
                raise ValueError(f"{prompt} Lookup failed ({result['error']['code']}).")
            if (result["query_date"] != "2026-11-14" or result["mode"] != mode
                    or [row["name"] for row in result["contacts"]] != names):
                raise ValueError(f"{prompt} The lookup did not match the expected date, mode, and contacts.")
            print(json.dumps({"prompt": prompt, "result": result}, ensure_ascii=False))


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--model", required=True, help="OpenAI model supporting function calls")
    args = parser.parse_args()
    key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not key:
        parser.exit(1, "Set OPENAI_API_KEY in your local environment.\n")
    try:
        run(args.model, key)
    except (ValueError, OSError, HTTPException) as error:
        parser.exit(1, f"{error}\n")
