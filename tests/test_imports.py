import importlib
import unittest


class AppImportTests(unittest.TestCase):
    def test_jhanel_imports(self):
        module = importlib.import_module("jhanel")
        self.assertIsNotNone(module)


if __name__ == "__main__":
    unittest.main()
