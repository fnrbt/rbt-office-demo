#!/bin/bash
# Build the demo into docs/ (GitHub Pages serves master:/docs).
#   OFFICE_DIST=<rbt-office checkout>/dist ./build.sh
# The samples come from office-hype's generators (docs/make_word.py, docs/make_excel.py)
# with the public-domain photos in samples/photos.
set -euo pipefail
cd "$(dirname "$0")"
: "${OFFICE_DIST:?set OFFICE_DIST to an rbt-office dist folder}"
npx vite build
touch docs/.nojekyll
echo "built docs/ against $OFFICE_DIST"
