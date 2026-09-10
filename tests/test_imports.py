import importlib
import unittest

from word_puzzle import WORD_PUZZLES, check_guess


class AppImportTests(unittest.TestCase):
    def test_jhanel_imports(self):
        module = importlib.import_module("jhanel")
        self.assertIsNotNone(module)

    def test_word_titles_follow_adjective_title_format(self):
        self.assertTrue(WORD_PUZZLES)
        for puzzle in WORD_PUZZLES:
            title = puzzle["word"]
            self.assertIn(" — ", title)
            left, right = title.split(" — ", 1)
            self.assertTrue(left.strip())
            self.assertTrue(right.strip())
            self.assertNotEqual(left, left.lower())
            self.assertNotEqual(right, right.lower())

    def test_phrase_guessing_ignores_spacing_and_punctuation(self):
        self.assertTrue(check_guess(1, "Ethnocentric — Sci-Fi Executive", 1))


if __name__ == "__main__":
    unittest.main()
