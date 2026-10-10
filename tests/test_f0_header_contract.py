import pathlib
import re
import unittest

ROOT = pathlib.Path(__file__).resolve().parents[1]
SCRIPTS = {
    "email": ROOT / "app/scripts/email.js",
    "sms": ROOT / "app/scripts/sms.js",
    "history": ROOT / "app/scripts/historico.js",
    "config": ROOT / "app/scripts/config.js",
}


class F0HeaderContractTests(unittest.TestCase):
    def test_normalization_preserves_accents(self):
        for name, path in SCRIPTS.items():
            with self.subTest(script=name):
                source = path.read_text(encoding="utf-8")
                self.assertNotIn('.normalize("NFD")', source)
                self.assertNotRegex(source, r"\\u0300-\\u036f")

    def test_duplicate_headers_are_rejected(self):
        for name in ("email", "sms", "history", "config"):
            with self.subTest(script=name):
                source = SCRIPTS[name].read_text(encoding="utf-8")
                self.assertIn("Cabeçalho duplicado ou ambíguo", source)

    def test_email_and_sms_do_not_create_missing_schema(self):
        for name in ("email", "sms"):
            with self.subTest(script=name):
                source = SCRIPTS[name].read_text(encoding="utf-8")
                self.assertIn("ID REGISTRO está ausente", source)
                self.assertIn("Estrutura da aba", source)
                self.assertNotIn("Se ID REGISTRO ainda não existir, ele é criado automaticamente", source)

    def test_sms_and_email_do_not_retry_uncertain_confirmation(self):
        for name in ("email", "sms"):
            with self.subTest(script=name):
                source = SCRIPTS[name].read_text(encoding="utf-8")
                self.assertIn("tentativa <= 1", source)
                self.assertNotIn("tentativa <= 3", source)

    def test_config_write_paths_guard_missing_or_ambiguous_headers(self):
        source = SCRIPTS["config"].read_text(encoding="utf-8")
        self.assertGreaterEqual(source.count("Cabeçalhos do Config ausentes ou ambíguos"), 2)
        self.assertIn("Cabeçalho ausente ou ambíguo", source)


if __name__ == "__main__":
    unittest.main(verbosity=2)
