"""Resolve one application version and optional release tag for both builds."""
from __future__ import annotations

import os
from pathlib import Path
import re


def build_metadata(event: str, ref: str, sha: str, default_branch: str, manual_tag: str = "") -> dict[str, str]:
    tag = ""
    if event == "push" and ref.startswith("refs/tags/v"):
        tag = ref.removeprefix("refs/tags/")
    elif event == "workflow_dispatch" and manual_tag:
        if ref != f"refs/heads/{default_branch}":
            raise ValueError("Manual releases must run on the repository default branch")
        tag = manual_tag
    if tag and not re.fullmatch(r"v[0-9]+\.[0-9]+\.[0-9]+(?:-[0-9A-Za-z]+(?:[.-][0-9A-Za-z]+)*)?", tag):
        raise ValueError("Release version must look like v0.0.37 or v0.0.37-rc.1")
    return {"version": tag or f"dev-{sha[:7]}", "release_tag": tag}


def main() -> int:
    metadata = build_metadata(
        os.environ["BUILD_EVENT"], os.environ["BUILD_REF"], os.environ["BUILD_SHA"],
        os.environ["DEFAULT_BRANCH"], os.environ.get("MANUAL_RELEASE_TAG", ""),
    )
    with Path(os.environ["GITHUB_OUTPUT"]).open("a", encoding="utf-8") as output:
        for name, value in metadata.items():
            output.write(f"{name}={value}\n")
    print(f"Build version: {metadata['version']}; release: {metadata['release_tag'] or 'artifacts only'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
