import unittest

from processor import apply_quote_style_to_segments, apply_quote_style_to_text, normalize_run_text


class NormalizeRunTextTests(unittest.TestCase):
    def test_keeps_outer_spaces_around_quotes(self):
        text = 'strankama "plavih" i "zelenih"'
        self.assertEqual(normalize_run_text(text, "academic"), text)

    def test_removes_inner_spaces_inside_quotes(self):
        text = 'strankama " plavih " i " zelenih "'
        expected = 'strankama "plavih" i "zelenih"'
        self.assertEqual(normalize_run_text(text, "academic"), expected)

    def test_fixes_space_before_comma_and_period(self):
        text = "Ovo je test , a ovo druga recenica ."
        expected = "Ovo je test, a ovo druga recenica."
        self.assertEqual(normalize_run_text(text, "academic"), expected)

    def test_joins_broken_hyphenated_word(self):
        self.assertEqual(normalize_run_text("pro- gram", "academic"), "program")
        self.assertEqual(normalize_run_text("pro - gram", "academic"), "program")

    def test_removes_false_spacing_in_numbers(self):
        self.assertEqual(normalize_run_text("1 000 i 12 345", "academic"), "1000 i 12345")

    def test_normalizes_ligatures(self):
        self.assertEqual(normalize_run_text("ofﬁce ﬂow", "academic"), "office flow")

    def test_normalizes_pdf_quotes(self):
        self.assertEqual(normalize_run_text("``tekst''", "academic"), '"tekst"')
        self.assertEqual(normalize_run_text("„tekst”", "academic"), '"tekst"')
        self.assertEqual(normalize_run_text("„razbojnik11", "academic"), '"razbojnik"')

    def test_does_not_turn_regular_number_11_into_quote(self):
        self.assertEqual(normalize_run_text('clan 11 stupa na snagu', "academic"), "clan 11 stupa na snagu")

    def test_fixes_ocr_quote_11_and_missing_space_before_quote(self):
        text = 'Jer nazivi "ubica" i "razbojnik11 bili su od njih cenjeni i odgovarali su nazivu"energičan"'
        expected = 'Jer nazivi "ubica" i "razbojnik" bili su od njih cenjeni i odgovarali su nazivu "energičan"'
        self.assertEqual(normalize_run_text(text, "academic"), expected)

    def test_normalizes_duplicate_punctuation(self):
        self.assertEqual(normalize_run_text("ovo .. ,,, test", "academic"), "ovo., test")

    def test_slash_spacing_depends_on_profile(self):
        self.assertEqual(normalize_run_text("i / ili", "academic"), "i/ili")
        self.assertEqual(normalize_run_text("i / ili", "legal"), "i / ili")

    def test_uniform_quote_style_english(self):
        self.assertEqual(apply_quote_style_to_text('"Crvena zvezda"', "english-double"), '"Crvena zvezda"')

    def test_uniform_quote_style_english_single(self):
        self.assertEqual(apply_quote_style_to_text('"Crvena zvezda"', "english-single"), "'Crvena zvezda'")

    def test_uniform_quote_style_serbian(self):
        self.assertEqual(apply_quote_style_to_text('"Crvena zvezda"', "serbian"), "„Crvena zvezda”")

    def test_uniform_quote_style_german(self):
        self.assertEqual(apply_quote_style_to_text('"Crvena zvezda"', "german"), "„Crvena zvezda“")

    def test_uniform_quote_style_serbian_across_segments(self):
        segments = ['"', "crveni", '"', ", ", '"', "zeleni", '"']
        expected = ["„", "crveni", "”", ", ", "„", "zeleni", "”"]
        self.assertEqual(apply_quote_style_to_segments(segments, "serbian"), expected)


if __name__ == "__main__":
    unittest.main()
