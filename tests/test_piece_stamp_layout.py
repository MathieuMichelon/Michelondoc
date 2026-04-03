import unittest
from pathlib import Path
import tempfile

import fitz

from app_v9_rotationfixed_stampfix2 import FileOps, A4_WIDTH, A4_HEIGHT


class PieceStampLayoutTests(unittest.TestCase):
    def test_layout_respects_bounds_on_mixed_sizes(self):
        page_sizes = [
            (A4_WIDTH, A4_HEIGHT),            # A4 portrait
            (A4_HEIGHT, A4_WIDTH),            # A4 paysage
            (300, 300),                       # petite page carrée
            (2480, 3508),                     # scan "pixel-like" très grand
            (1200, 800),                      # format intermédiaire
        ]

        for width, height in page_sizes:
            with self.subTest(width=width, height=height):
                layout = FileOps._compute_piece_stamp_layout(fitz.Rect(0, 0, width, height))
                outer = layout["outer"]

                self.assertGreaterEqual(outer.x0, FileOps.STAMP_MIN_MARGIN)
                self.assertGreaterEqual(outer.y0, FileOps.STAMP_MIN_MARGIN)
                self.assertLessEqual(outer.x1, width - FileOps.STAMP_MIN_MARGIN + 1e-6)
                self.assertLessEqual(outer.y1, height - FileOps.STAMP_MIN_MARGIN + 1e-6)
                self.assertGreater(outer.width, 0)
                self.assertGreater(outer.height, 0)

                self.assertGreaterEqual(outer.width, min(FileOps.STAMP_MIN_WIDTH, width - (2 * FileOps.STAMP_MIN_MARGIN)))
                self.assertLessEqual(outer.width, min(FileOps.STAMP_MAX_WIDTH, width - (2 * FileOps.STAMP_MIN_MARGIN)) + 1e-6)

                self.assertGreaterEqual(layout["title_font"], 9.0)
                self.assertGreaterEqual(layout["piece_font"], 8.5)

    def test_add_piece_stamp_normalizes_rotation_and_keeps_stamp_in_top_right(self):
        with tempfile.TemporaryDirectory() as tmpdir:
            src = Path(tmpdir) / "in.pdf"
            dst = Path(tmpdir) / "out.pdf"

            doc = fitz.open()
            for rotation in (0, 90, 180, 270):
                page = doc.new_page(width=A4_WIDTH, height=A4_HEIGHT)
                if rotation:
                    page.set_rotation(rotation)
            doc.save(src)
            doc.close()

            FileOps.add_piece_stamp(src, dst, piece_label="12", stamp_title="Cabinet Test")

            stamped = fitz.open(dst)
            self.assertEqual(len(stamped), 4)

            for page in stamped:
                self.assertEqual(page.rotation, 0)
                blocks = [b for b in page.get_text("blocks") if isinstance(b[4], str)]
                self.assertTrue(any("Cabinet Test" in b[4] for b in blocks))
                self.assertTrue(any("Pièce n° : 12" in b[4] for b in blocks))

                title_block = next(b for b in blocks if "Cabinet Test" in b[4])
                # bloc attendu près du haut droit visible
                self.assertGreater(title_block[0], page.rect.width * 0.60)
                self.assertLess(title_block[1], page.rect.height * 0.20)

            stamped.close()


if __name__ == "__main__":
    unittest.main()
