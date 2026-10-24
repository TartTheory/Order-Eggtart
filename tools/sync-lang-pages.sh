#!/bin/sh
# ch/ and en/ are copies of index.html (same file; the language comes from the URL path). Run before every release.
cd "$(dirname "$0")/.." || exit 1
mkdir -p ch en && cp index.html ch/index.html && cp index.html en/index.html
