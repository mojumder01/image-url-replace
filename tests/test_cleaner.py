"""Tests for the rule-based cleaner (one test per prompt rule)."""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from catprep.cleaner import check_row, clean_rows, parse, split_plain_bullets  # noqa: E402

IMG = "https://sl-dev-s3.s3.amazonaws.com/product/a.webp"
IMG2 = "https://sl-dev-s3.s3.amazonaws.com/product/b.webp"


def clean(name, he, hb, de, db):
    out, stats, notes = clean_rows([(name, he, hb, de, db)])
    return out[0], notes[0]


class CleanerTest(unittest.TestCase):
    def test_strips_attributes_and_wrappers(self):
        (he, hb, de, db), _ = clean(
            "Phone Case",
            '<div style="color:red"><ul><li class="x"><span>Soft <b>silicone</b></span></li>'
            '<li><a href="http://x">Shock proof</a></li></ul></div>', None,
            '<h1 id="t">Phone Case</h1><p style="margin:0">Fits well.</p><img class="i" src="%s">' % IMG, None)
        self.assertEqual(he, "<ul><li>Soft silicone</li><li>Shock proof</li></ul>")
        self.assertEqual(hb, he)  # one highlight empty -> copied
        self.assertEqual(de, f'<h2>Phone Case</h2><p>Fits well.</p><img src="{IMG}"/>')

    def test_highlights_from_description(self):
        (he, hb, de, db), notes = clean(
            "Lamp", "<p><img src='%s'></p>" % IMG, "  ",
            "<h2>Lamp</h2><ul><li>Bright LED light</li><li>USB charging port</li></ul>", None)
        self.assertEqual(he, "<ul><li>Bright LED light</li><li>USB charging port</li></ul>")
        self.assertIn("highlights_from_description", notes)
        self.assertIn(IMG, de)  # image from the 'empty' highlight is kept

    def test_description_from_highlights(self):
        (he, hb, de, db), notes = clean("Mug", "<ul><li>Ceramic mug 300ml</li></ul>", None, "<p> </p>", "")
        self.assertEqual(de, "<h2>Mug</h2><ul><li>Ceramic mug 300ml</li></ul>")
        self.assertIn("description_from_highlights", notes)

    def test_paragraph_and_plain_text_bullets(self):
        self.assertEqual(split_plain_bullets("Product details$ Type: Bag$ Brand: X"),
                         ["Product details", "Type: Bag", "Brand: X"])
        self.assertEqual(split_plain_bullets("* Soft fabric * Machine washable"),
                         ["Soft fabric", "Machine washable"])
        self.assertEqual(split_plain_bullets("Size: 54 * 45 * 20 (L*W*H)"), ["Size: 54 * 45 * 20 (L*W*H)"])
        (he, *_), _ = clean("Bag", "• Strong zipper • Two pockets • Waterproof", None, "<p>A bag.</p>", None)
        self.assertEqual(he, "<ul><li>Strong zipper</li><li>Two pockets</li><li>Waterproof</li></ul>")

    def test_junk_and_duplicate_bullets(self):
        (he, *_), _ = clean("Toy", "<ul><li>a</li><li>Made of strong plastic</li><li>Made of strong plastic</li>"
                                   "<li>200 g</li><li>Suitable for kids above three years</li>"
                                   "<li>Suitable for kids above years</li></ul>", None, "<p>Toy.</p>", None)
        self.assertEqual(he, "<ul><li>Made of strong plastic</li><li>200 g</li>"
                             "<li>Suitable for kids above three years</li></ul>")

    def test_concatenated_bullet_removed(self):
        (he, *_), _ = clean("Pen", "<ul><li>Blue ink</li><li>Smooth writing tip</li>"
                                   "<li>Blue ink Smooth writing tip</li></ul>", None, "<p>Pen.</p>", None)
        self.assertEqual(he, "<ul><li>Blue ink</li><li>Smooth writing tip</li></ul>")

    def test_repeated_paragraph_dropped_and_layout_order(self):
        (_, _, de, _), _ = clean(
            "Fan", None, None,
            '<img src="%s"><p>Quiet motor with three speed settings.</p><h2>Desk Fan</h2>'
            '<ul><li>Quiet motor</li><li>Three speed settings</li></ul>'
            '<p>Comes with a two year warranty from the brand.</p>'
            '<table><tr><td>Power</td><td>45W</td><td></td></tr></table>' % IMG, None)
        self.assertNotIn("<p>Quiet motor", de)
        self.assertIn("<p>Comes with a two year warranty from the brand.</p>", de)
        self.assertTrue(de.index("<ul>") < de.index("<p>") < de.index("<table>") < de.index("<img"))
        self.assertIn("<tr><td>Power</td><td>45W</td></tr>", de)

    def test_corrupted_markup_and_symbols(self):
        (he, *_), _ = clean("Cap", 'Cotton cap" style="font-size:12px;">Breathable ✔ 😀 א', None,
                            "<p>Cap.</p>", None)
        self.assertNotIn("style", he)
        self.assertNotIn(">", he.replace("<ul>", "").replace("<li>", "").replace("</li>", "").replace("</ul>", ""))
        self.assertNotIn("😀", he)
        self.assertNotIn("א", he)

    def test_bangla_kept_and_not_junk(self):
        (_, hb, *_), _ = clean("Toy", "<ul><li>Red</li></ul>", "<ul><li>রঙ: লাল এবং নীল।</li></ul>", "<p>x y z</p>", None)
        self.assertEqual(hb, "<ul><li>রঙ: লাল এবং নীল।</li></ul>")

    def test_all_images_kept(self):
        row = ("Shoe", f'<img src="{IMG2}"><ul><li>Rubber sole</li></ul>', None,
               f'<h2>Shoe</h2><p>Comfortable walking shoe.</p><img src="{IMG}">', None)
        out, _, _ = clean_rows([row])
        self.assertEqual(check_row(row[1:], out[0]), [])
        self.assertIn(IMG, out[0][2])
        self.assertIn(IMG2, out[0][2])

    def test_generic_highlights_enriched(self):
        generic = "<ul><li>High quality product</li><li>Best price</li></ul>"
        rows = [(f"P{i}", generic, None,
                 f"<h2>P{i}</h2><ul>" + "".join(f"<li>Specific feature {i}-{k}</li>" for k in range(4)) + "</ul>",
                 None) for i in range(2)]
        out, stats, _ = clean_rows(rows)
        self.assertEqual(stats["highlights_enriched"], 2)
        self.assertIn("Specific feature 0-0", out[0][0])

    def test_parse_survives_broken_html(self):
        blocks, images = parse('<ul><li>One<li>Two</ul><p>Unclosed <b>bold<td>cell</ul></table>')
        self.assertEqual(blocks[0], ("list", ["One", "Two"]))


if __name__ == "__main__":
    unittest.main()
