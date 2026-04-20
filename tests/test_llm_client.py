import os
import sys
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import Mock, patch

SCRIPTS_DIR = Path(__file__).resolve().parents[1] / "scripts"
if str(SCRIPTS_DIR) not in sys.path:
    sys.path.append(str(SCRIPTS_DIR))

from llm_client import call_gemini_for_json


class LlmClientTests(unittest.TestCase):
    def test_non_retryable_error_fails_fast(self):
        client = Mock()
        client.models.generate_content = Mock(side_effect=RuntimeError("invalid schema"))
        with patch("llm_client._candidate_models", return_value=["gemini-a"]), \
                patch.dict(os.environ, {"GEMINI_MAX_RETRIES": "3", "GEMINI_INITIAL_BACKOFF_SECONDS": "0"}, clear=False), \
                patch("llm_client.time.sleep") as sleep_mock:
            with self.assertRaises(RuntimeError):
                call_gemini_for_json(client, ["prompt"])
        self.assertEqual(client.models.generate_content.call_count, 1)
        sleep_mock.assert_not_called()

    def test_retryable_error_retries_and_succeeds(self):
        client = Mock()
        client.models.generate_content = Mock(
            side_effect=[RuntimeError("503 UNAVAILABLE"), SimpleNamespace(text='{"ok": true}')]
        )
        with patch("llm_client._candidate_models", return_value=["gemini-a"]), \
                patch.dict(os.environ, {"GEMINI_MAX_RETRIES": "3", "GEMINI_INITIAL_BACKOFF_SECONDS": "0"}, clear=False), \
                patch("llm_client.time.sleep") as sleep_mock:
            result = call_gemini_for_json(client, ["prompt"])
        self.assertEqual(result, {"ok": True})
        self.assertEqual(client.models.generate_content.call_count, 2)
        sleep_mock.assert_called_once_with(0.0)


if __name__ == "__main__":
    unittest.main()
