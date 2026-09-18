import unittest

from docx import Document
from docx.enum.section import WD_SECTION_START

from processor import (
    apply_quote_style_to_segments,
    apply_quote_style_to_text,
    delete_empty_paragraphs,
    normalize_run_text,
)
from audit_log import AuditLog
from processor import _run_fix_broken_sentences, fix_broken_sentences


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

    def test_keeps_an_ellipsis(self):
        self.assertEqual(
            normalize_run_text("trebalo je imati vere... vere koju", "academic"),
            "trebalo je imati vere... vere koju",
        )

    def test_tidies_a_spaced_or_overlong_ellipsis(self):
        self.assertEqual(normalize_run_text("vere. . . vere", "academic"), "vere... vere")
        self.assertEqual(normalize_run_text("vere..... vere", "academic"), "vere... vere")

    def test_keeps_an_en_dash_between_words(self):
        self.assertEqual(
            normalize_run_text("znamo da to mozemo – kao sto znamo", "academic"),
            "znamo da to mozemo – kao sto znamo",
        )

    def test_reduces_an_em_dash_to_an_en_dash(self):
        self.assertEqual(normalize_run_text("mozemo — kao sto", "academic"), "mozemo – kao sto")
        self.assertEqual(normalize_run_text("— Zdravo", "academic"), "– Zdravo")

    def test_spaces_out_a_dash_glued_between_words(self):
        self.assertEqual(normalize_run_text("rec—rec", "academic"), "rec – rec")

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


class DeleteEmptyParagraphsTests(unittest.TestCase):
    @staticmethod
    def _clean(doc):
        return delete_empty_paragraphs(doc, lambda message: None)

    def test_keeps_last_paragraph_in_otherwise_empty_table_cell(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=2)
        table.cell(0, 0).text = "sadrzaj"
        table.cell(0, 1).text = ""

        self._clean(doc)

        self.assertEqual(len(table.cell(0, 1).paragraphs), 1)

    def test_removes_extra_blank_paragraphs_inside_table_cell(self):
        doc = Document()
        table = doc.add_table(rows=1, cols=1)
        cell = table.cell(0, 0)
        cell.text = ""
        cell.add_paragraph("")
        cell.add_paragraph("")

        removed = self._clean(doc)

        self.assertEqual(len(cell.paragraphs), 1)
        self.assertEqual(removed, 2)

    def test_keeps_a_blank_paragraph_that_carries_a_section_break(self):
        doc = Document()
        doc.add_paragraph("prvi")
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        doc.add_paragraph("drugi")

        self._clean(doc)

        self.assertEqual(len(doc.sections), 2)
        self.assertEqual([p.text for p in doc.paragraphs if p.text.strip()], ["prvi", "drugi"])

    def test_removes_all_blank_paragraphs_from_document_body(self):
        doc = Document()
        doc.add_paragraph("prvi")
        doc.add_paragraph("")
        doc.add_paragraph("")
        doc.add_paragraph("drugi")

        removed = self._clean(doc)

        self.assertEqual([p.text for p in doc.paragraphs], ["prvi", "drugi"])
        self.assertEqual(removed, 2)


class MergeAuditTests(unittest.TestCase):
    @staticmethod
    def _audit():
        return AuditLog("src.docx", "dst.docx", "academic", "serbian", ".docx", {})

    def test_records_one_change_per_merge_and_nothing_else(self):
        doc = Document()
        doc.add_paragraph("Prva recenica koja se nastavlja")
        doc.add_paragraph("nastavak prve recenice.")
        doc.add_paragraph("Druga recenica koja se nastavlja")
        doc.add_paragraph("nastavak druge recenice.")
        doc.add_paragraph("Peta recenica stoji sama.")
        doc.add_paragraph("Sesta recenica stoji sama.")
        audit = self._audit()

        _run_fix_broken_sentences(doc, lambda message: None, "academic", None, None, audit)

        self.assertEqual(audit.stats["merged_paragraph_pairs"], 2)
        self.assertEqual(len(audit.changes["paragraph_merge"]), 2)

    def test_recorded_change_shows_the_two_merged_paragraphs(self):
        doc = Document()
        doc.add_paragraph("Prva recenica koja se nastavlja")
        doc.add_paragraph("nastavak prve recenice.")
        audit = self._audit()

        _run_fix_broken_sentences(doc, lambda message: None, "academic", None, None, audit)

        change = audit.changes["paragraph_merge"][0]
        self.assertIn("Prva recenica koja se nastavlja", change["before"])
        self.assertIn("nastavak prve recenice.", change["before"])
        self.assertEqual(
            change["after"],
            "Prva recenica koja se nastavlja nastavak prve recenice.",
        )


class MergeAcrossSectionBreakTests(unittest.TestCase):
    @staticmethod
    def _merge(doc):
        return fix_broken_sentences(doc, lambda message: None, "academic")

    def test_merges_a_sentence_split_across_a_section_break(self):
        doc = Document()
        doc.add_paragraph("Prva recenica koja se nastavlja")
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        doc.add_paragraph("nastavak prve recenice.")

        merges = self._merge(doc)

        self.assertEqual(merges, 1)
        self.assertIn(
            "Prva recenica koja se nastavlja nastavak prve recenice.",
            [p.text for p in doc.paragraphs],
        )

    def test_the_section_break_survives_the_merge(self):
        doc = Document()
        doc.add_paragraph("Prva recenica koja se nastavlja")
        doc.add_section(WD_SECTION_START.NEW_PAGE)
        doc.add_paragraph("nastavak prve recenice.")

        self._merge(doc)

        self.assertEqual(len(doc.sections), 2)

    def test_does_not_merge_across_a_plain_blank_paragraph(self):
        doc = Document()
        doc.add_paragraph("Prva recenica koja se nastavlja")
        doc.add_paragraph("")
        doc.add_paragraph("nastavak prve recenice.")

        self.assertEqual(self._merge(doc), 0)


if __name__ == "__main__":
    unittest.main()
