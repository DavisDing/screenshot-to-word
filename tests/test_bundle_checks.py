from __future__ import annotations

from pathlib import Path
import sys
import unittest


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.check_bundle import check_contents, REQUIRED_MODULES


class BundleChecksTests(unittest.TestCase):
    def paths(self, platform: str) -> list[str]:
        prefix = "_internal/" if platform == "win32" else "Contents/Resources/"
        paths = [prefix + name for name in (
            "icon.ico", "assets/icon.png", "assets/chevron-down.svg",
            "docx/templates/default.docx",
        )]
        suffix = ".pyd" if platform == "win32" else ".abi3.so"
        paths.extend(prefix + name + suffix for name in (
            "PySide6/QtCore", "PySide6/QtGui", "PySide6/QtWidgets", "PySide6/QtSvg",
            "shiboken6/Shiboken", "lxml/etree",
        ))
        plugins = ("platforms/qwindows.dll", "imageformats/qsvg.dll") if platform == "win32" else (
            "platforms/libqcocoa.dylib", "imageformats/libqsvg.dylib",
        )
        paths.extend(prefix + "PySide6/Qt/plugins/" + name for name in plugins)
        return paths

    def test_both_platform_layouts(self):
        for platform in ("win32", "darwin"):
            with self.subTest(platform=platform):
                self.assertEqual(check_contents(self.paths(platform), REQUIRED_MODULES, platform), [])

    def test_missing_runtime_resources_or_plugins_fail(self):
        for missing in ("chevron-down.svg", "default.docx", "qwindows.dll", "qsvg.dll", "Shiboken.pyd"):
            with self.subTest(missing=missing):
                paths = [path for path in self.paths("win32") if not path.endswith(missing)]
                self.assertTrue(check_contents(paths, REQUIRED_MODULES, "win32"))

    def test_missing_delayed_import_or_pillow_fails(self):
        for missing in ("qt_app.services.word_service", "PIL.Image", "PySide6.support.deprecated"):
            with self.subTest(missing=missing):
                errors = check_contents(self.paths("darwin"), REQUIRED_MODULES - {missing}, "darwin")
                self.assertIn(f"Missing bundled module: {missing}", errors)

    def test_unused_stacks_fail_for_dll_framework_and_qml_data(self):
        for extra in (
            "_internal/PySide6/Qt6WebEngineCore.dll",
            "Contents/Frameworks/QtQuick.framework/Versions/A/QtQuick",
            "_internal/PySide6/QtMultimedia.pyd",
            "_internal/PySide6/Qt6Qml.dll",
            "_internal/PySide6/Qt3DCore.pyd",
            "_internal/PySide6/Qt/plugins/platforminputcontexts/qtvirtualkeyboardplugin.dll",
            "_internal/PySide6/Qt/qml/Qt/labs/qmldir",
        ):
            with self.subTest(extra=extra):
                errors = check_contents(self.paths("win32") + [extra], REQUIRED_MODULES, "win32")
                self.assertTrue(any("Unused Qt stacks" in error for error in errors))


if __name__ == "__main__":
    unittest.main()
