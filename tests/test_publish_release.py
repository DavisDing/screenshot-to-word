"""Exercise the real release shell in an isolated repo with a recording gh stub."""
from __future__ import annotations

import os
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
import unittest


ROOT = Path(__file__).resolve().parents[1]


@unittest.skipUnless(sys.platform != "win32" and shutil.which("bash") and shutil.which("git"), "release shell runs on POSIX")
class PublishReleaseTests(unittest.TestCase):
    def setUp(self):
        self.temporary = TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.base = Path(self.temporary.name)
        self.git("init", "-q")
        self.git("-c", "user.name=Synthetic Test", "-c", "user.email=synthetic@example.invalid",
                 "commit", "--allow-empty", "-m", "isolated synthetic commit", "-q")
        self.sha = self.git("rev-parse", "HEAD").strip()
        assets = self.base / "release-assets"
        assets.mkdir()
        for platform in ("windows", "macos"):
            (assets / f"DesktopTestTool-{platform}.zip").write_bytes(b"synthetic package for isolated test")
        self.recording = self.base / "gh-args.txt"
        bin_path = self.base / "bin"
        bin_path.mkdir()
        stub = bin_path / "gh"
        stub.write_text('#!/usr/bin/env bash\nprintf "%s\\n" "$@" > "$RELEASE_TEST_RECORD"\n', encoding="utf-8")
        stub.chmod(0o755)
        self.environment = dict(os.environ, PATH=str(bin_path) + os.pathsep + os.environ["PATH"],
                                RELEASE_TEST_RECORD=str(self.recording), RELEASE_TAG="v0.0.37", RELEASE_SHA=self.sha)

    def git(self, *args):
        return subprocess.check_output(["git", "-c", "commit.gpgsign=false", "-c", "core.hooksPath=" + str(self.base / "no-hooks"), *args],
                                       cwd=self.base, text=True, stderr=subprocess.STDOUT)

    def publish(self):
        return subprocess.run(["bash", str(ROOT / "scripts/publish_release.sh")], cwd=self.base,
                              env=self.environment, text=True, capture_output=True, timeout=15)

    def test_new_release_supplies_both_packages_and_exact_commit(self):
        result = self.publish()
        self.assertEqual(result.returncode, 0, result.stderr)
        arguments = self.recording.read_text().splitlines()
        self.assertEqual(arguments[:3], ["release", "create", "v0.0.37"])
        self.assertIn("release-assets/DesktopTestTool-windows.zip", arguments)
        self.assertIn("release-assets/DesktopTestTool-macos.zip", arguments)
        self.assertIn("release-assets/SHA256SUMS.txt", arguments)
        manifest = (self.base / "release-assets/SHA256SUMS.txt").read_text()
        self.assertIn("DesktopTestTool-windows.zip", manifest)
        self.assertIn("DesktopTestTool-macos.zip", manifest)
        self.assertEqual(arguments[arguments.index("--target") + 1], self.sha)

    def test_wrong_commit_tag_stops_before_publishing(self):
        self.git("tag", "v0.0.37")
        self.git("-c", "user.name=Synthetic Test", "-c", "user.email=synthetic@example.invalid",
                 "commit", "--allow-empty", "-m", "second isolated commit", "-q")
        self.environment["RELEASE_SHA"] = self.git("rev-parse", "HEAD").strip()
        result = self.publish()
        self.assertNotEqual(result.returncode, 0)
        self.assertIn("different commit", result.stderr)
        self.assertFalse(self.recording.exists())

    def test_matching_tag_and_missing_package(self):
        self.git("tag", "v0.0.37")
        result = self.publish()
        self.assertEqual(result.returncode, 0, result.stderr)
        self.recording.unlink()
        (self.base / "release-assets/DesktopTestTool-macos.zip").unlink()
        self.assertNotEqual(self.publish().returncode, 0)
        self.assertFalse(self.recording.exists())


if __name__ == "__main__":
    unittest.main()
