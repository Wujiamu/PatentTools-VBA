"""Static regression guards for the VBA source.

These checks verify the source wiring and encoding-sensitive constants. They do
not replace importing, compiling, or running the modules in Word.
"""

from pathlib import Path
import re
import unittest


ROOT = Path(__file__).resolve().parents[1]
SRC = ROOT / "src"


class SourceGuards(unittest.TestCase):
    def read(self, name: str) -> str:
        return (SRC / name).read_text(encoding="utf-8")

    def test_source_files_are_strict_utf8(self) -> None:
        for path in SRC.iterdir():
            if path.suffix.lower() in {".bas", ".cls", ".frm"}:
                path.read_bytes().decode("utf-8")

    def test_compat_ranges_use_unsigned_decimal_values(self) -> None:
        text = self.read("modPatentCompat.bas")
        self.assertIn("value >= 13312", text)
        self.assertIn("value <= 40959", text)
        self.assertIn("value >= 65296", text)
        self.assertNotRegex(text, r"&H(?:9FFF|FF(?:10|19|21|3A|41|5A))")

    def test_figure_modules_use_pure_vba_fallback(self) -> None:
        for name in ("modEmbodimentFigureTag.bas", "modClaimFigureTag.bas"):
            text = self.read(name)
            self.assertIn("ParseFigureTagText", text)
            self.assertIn("CreateStringMap", text)
            self.assertNotIn("VBScript.RegExp", text)

    def test_revision_guards_are_wired(self) -> None:
        for name in (
            "modEmbodimentFigureTag.bas",
            "modClaimFigureTag.bas",
            "modClaimToInvention.bas",
            "modHeaderSync.bas",
        ):
            self.assertIn("RangeTouchesExistingRevision", self.read(name)) if name in {
                "modEmbodimentFigureTag.bas",
                "modClaimFigureTag.bas",
            } else self.assertIn("RangeHasExistingRevision", self.read(name))

        lite = self.read("modCleanFormatTrackLite.bas")
        self.assertIn("oldStatusBar", lite)
        self.assertIn("EndCustomUndoRecord undoStarted", lite)
        full = self.read("modCleanFormatFull.bas")
        self.assertIn("DocumentHasExistingRevisions", full)

    def test_dispatcher_entries_still_exist(self) -> None:
        dispatcher = self.read("modToolDispatcher.bas")
        for entry in (
            "权利要求标号",
            "权利要求转发明内容",
            "具体实施方式标号",
            "页眉双模式同步",
            "一键改格式_全文修改_不兼容修订",
            "一键改格式_选中部分修改_兼容修订",
        ):
            self.assertIn(entry, dispatcher)


if __name__ == "__main__":
    unittest.main()
