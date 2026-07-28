#!/usr/bin/env bash
# Runs the shipWithReport contact-matching tests.
#
# Office Scripts has no test runner, so the script and the test file are
# concatenated (Office Scripts share one file scope) and compiled with tsc.
# Pass a path to test a different revision, e.g.:
#   git show HEAD~1:office-scripts/shipWithReport.ts > /tmp/old.ts
#   office-scripts/tests/run.sh /tmp/old.ts
set -euo pipefail
here="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
script="${1:-$here/../shipWithReport.ts}"
out="$(mktemp -d)"
trap 'rm -rf "$out"' EXIT

cat "$script" "$here/shipWithReport.test.ts" > "$out/bundle.ts"
# The test harness uses console; the script itself must stay DOM-free, which is
# verified separately below.
tsc --target es2019 --lib es2019,dom --ignoreDeprecations 6.0 --outFile "$out/bundle.js" \
    "$here/excelscript.d.ts" "$out/bundle.ts"
node "$out/bundle.js"

# The Office Script itself may only use what Office Scripts provides.
tsc --noEmit --target es2019 --lib es2019 "$here/excelscript.d.ts" "$script"
echo "typecheck clean (no console/DOM dependency in the script)"
