#!/usr/bin/env bash
set -euo pipefail

: "${RELEASE_TAG:?Release tag is required}"
: "${RELEASE_SHA:?Release commit is required}"

test -s release-assets/DesktopTestTool-windows.zip
test -s release-assets/DesktopTestTool-macos.zip

# An existing version must identify the exact commit used for both builds.
if git show-ref --verify --quiet "refs/tags/$RELEASE_TAG"; then
  if [ "$(git rev-parse "refs/tags/$RELEASE_TAG^{commit}")" != "$RELEASE_SHA" ]; then
    echo "Release tag $RELEASE_TAG points to a different commit" >&2
    exit 1
  fi
fi

gh release create "$RELEASE_TAG" \
  release-assets/DesktopTestTool-windows.zip \
  release-assets/DesktopTestTool-macos.zip \
  --target "$RELEASE_SHA" \
  --title "Release $RELEASE_TAG" \
  --notes "Automated desktop release for Windows and macOS."
