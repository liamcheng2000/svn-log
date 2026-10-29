"""日历弹窗定位自测：python -m unittest discover -s tests -v"""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from app.ui import _clamp_popup_position  # noqa: E402


class PopupPositionTests(unittest.TestCase):
    def test_below_anchor_when_room(self) -> None:
        x, y = _clamp_popup_position(500, 300, 60, 24, 240, 220, 1920, 1080)
        self.assertEqual(x, 500 + (60 - 240) // 2)
        self.assertEqual(y, 300 + 24 + 4)

    def test_flip_above_near_bottom(self) -> None:
        x, y = _clamp_popup_position(500, 900, 60, 24, 240, 220, 1920, 1080)
        self.assertEqual(y, 900 - 220 - 4)
        self.assertLessEqual(y + 220, 1080 - 8)

    def test_clamp_on_small_screen(self) -> None:
        screen_w, screen_h = 1024, 728
        x, y = _clamp_popup_position(900, 700, 60, 24, 240, 220, screen_w, screen_h)
        self.assertGreaterEqual(x, 8)
        self.assertLessEqual(x + 240, screen_w - 8)
        self.assertGreaterEqual(y, 8)
        self.assertLessEqual(y + 220, screen_h - 8)

    def test_clamp_right_edge(self) -> None:
        x, y = _clamp_popup_position(1900, 100, 60, 24, 240, 220, 1920, 1080)
        self.assertLessEqual(x + 240, 1920 - 8)
        self.assertGreaterEqual(x, 8)

    def test_never_negative(self) -> None:
        x, y = _clamp_popup_position(-50, -40, 60, 24, 240, 220, 1920, 1080)
        self.assertGreaterEqual(x, 8)
        self.assertGreaterEqual(y, 8)


if __name__ == "__main__":
    unittest.main()
