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
                patch.dict(
                    os.environ,
                    {
                        "GEMINI_MAX_RETRIES_PER_MODEL": "3",
                        "GEMINI_INITIAL_BACKOFF_SECONDS": "0",
                        "GEMINI_MAX_BACKOFF_SECONDS": "1",
                    },
                    clear=False,
                ), \
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
                patch.dict(
                    os.environ,
                    {
                        "GEMINI_MAX_RETRIES_PER_MODEL": "3",
                        "GEMINI_INITIAL_BACKOFF_SECONDS": "0",
                        "GEMINI_MAX_BACKOFF_SECONDS": "1",
                    },
                    clear=False,
                ), \
                patch("llm_client.time.sleep") as sleep_mock:
            result = call_gemini_for_json(client, ["prompt"])
        self.assertEqual(result, {"ok": True})
        self.assertEqual(client.models.generate_content.call_count, 2)
        sleep_mock.assert_called_once_with(0.0)

    def test_retryable_error_switches_model_when_exhausted(self):
        client = Mock()
        client.models.generate_content = Mock(
            side_effect=[
                RuntimeError("503 UNAVAILABLE"),
                RuntimeError("503 UNAVAILABLE"),
                RuntimeError("503 UNAVAILABLE"),
                SimpleNamespace(text='{"ok": true}'),
            ]
        )
        with patch("llm_client._candidate_models", return_value=["gemini-a", "gemini-b"]), \
                patch.dict(
                    os.environ,
                    {
                        "GEMINI_MAX_RETRIES_PER_MODEL": "3",
                        "GEMINI_INITIAL_BACKOFF_SECONDS": "0",
                        "GEMINI_MAX_BACKOFF_SECONDS": "1",
                    },
                    clear=False,
                ), \
                patch("llm_client.time.sleep") as sleep_mock:
            result = call_gemini_for_json(client, ["prompt"])
        self.assertEqual(result, {"ok": True})
        self.assertEqual(client.models.generate_content.call_count, 4)
        self.assertEqual(sleep_mock.call_count, 2)
        self.assertEqual(
            [call.kwargs["model"] for call in client.models.generate_content.call_args_list],
            ["gemini-a", "gemini-a", "gemini-a", "gemini-b"],
        )

    def test_fallback_can_be_disabled(self):
        with patch.dict(
            os.environ,
            {
                "GEMINI_MODEL": "gemini-main",
                "GEMINI_FALLBACK_MODELS": "gemini-main,gemini-x,gemini-y",
                "GEMINI_ENABLE_MODEL_FALLBACK": "false",
            },
            clear=False,
        ):
            from llm_client import _candidate_models
            self.assertEqual(_candidate_models(), ["gemini-main"])

    def test_all_models_exhausted_raises_consolidated_error(self):
        client = Mock()
        client.models.generate_content = Mock(side_effect=RuntimeError("503 UNAVAILABLE"))
        with patch("llm_client._candidate_models", return_value=["gemini-a", "gemini-b"]), \
                patch.dict(
                    os.environ,
                    {
                        "GEMINI_MAX_RETRIES_PER_MODEL": "2",
                        "GEMINI_INITIAL_BACKOFF_SECONDS": "0",
                        "GEMINI_MAX_BACKOFF_SECONDS": "1",
                    },
                    clear=False,
                ), \
                patch("llm_client.time.sleep"):
            with self.assertRaises(RuntimeError) as ctx:
                call_gemini_for_json(client, ["prompt"])
        self.assertIn("agotó todos los modelos", str(ctx.exception))
        self.assertEqual(client.models.generate_content.call_count, 4)

    def test_invalid_retry_env_value_raises_descriptive_error(self):
        client = Mock()
        with patch("llm_client._candidate_models", return_value=["gemini-a"]), \
                patch.dict(os.environ, {"GEMINI_MAX_RETRIES_PER_MODEL": "abc"}, clear=False):
            with self.assertRaises(RuntimeError) as ctx:
                call_gemini_for_json(client, ["prompt"])
        self.assertIn("GEMINI_MAX_RETRIES_PER_MODEL", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
