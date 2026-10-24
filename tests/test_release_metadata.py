from __future__ import annotations

from pathlib import Path
import sys
import unittest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from scripts.prepare_build import build_metadata


class ReleaseMetadataTests(unittest.TestCase):
    SHA = "843a811ac4f391a3c91c11408a3006ead496fbb4"

    def test_branch_pr_and_manual_builds_do_not_publish(self):
        for event, ref in (("push", "refs/heads/main"), ("push", "refs/heads/feature/test"),
                           ("pull_request", "refs/pull/1/merge"), ("workflow_dispatch", "refs/heads/main")):
            with self.subTest(event=event, ref=ref):
                self.assertEqual(build_metadata(event, ref, self.SHA, "main"),
                                 {"version": "dev-843a811", "release_tag": ""})

    def test_tag_push_and_explicit_manual_release_share_packaged_version(self):
        for tag in ("v0.0.37", "v1.0.0-rc.1"):
            for event, ref, manual in (("push", f"refs/tags/{tag}", ""),
                                       ("workflow_dispatch", "refs/heads/main", tag)):
                with self.subTest(event=event, tag=tag):
                    self.assertEqual(build_metadata(event, ref, self.SHA, "main", manual),
                                     {"version": tag, "release_tag": tag})

    def test_manual_release_requires_default_branch(self):
        with self.assertRaisesRegex(ValueError, "default branch"):
            build_metadata("workflow_dispatch", "refs/heads/feature", self.SHA, "main", "v0.0.37")
        self.assertEqual(build_metadata("workflow_dispatch", "refs/heads/trunk", self.SHA, "trunk", "v0.0.37")["release_tag"],
                         "v0.0.37")

    def test_invalid_versions_fail_before_packaging(self):
        for tag in ("v", "v1", "v1.2.3/other", "v1.2.3\nversion=other", "v1.2.3$(command)", "v1.2.3 "):
            with self.subTest(tag=tag), self.assertRaisesRegex(ValueError, "Release version"):
                build_metadata("workflow_dispatch", "refs/heads/main", self.SHA, "main", tag)


if __name__ == "__main__":
    unittest.main()
