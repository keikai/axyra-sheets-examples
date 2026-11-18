#!/usr/bin/env bash
# Runs every test class in its own process.
#
# The evaluation build of Axyra.Sheets allows 100 opens/saves/renders per process, so the
# test classes cannot share one process (the Java project gets the same isolation from
# Surefire forking one JVM per test class). xunit v3 test projects are executables, so this
# builds once and runs the test executable once per class.
#
#   ./run-tests.sh                 # every class except those with the known-issue trait
#   ./run-tests.sh --known-issues  # only tests with [Trait("Category", "known-issue")]
#   ./run-tests.sh SomeTestClass   # one class (simple or full name)
set -uo pipefail
cd "$(dirname "$0")"
export PATH="$HOME/.dotnet:$PATH" DOTNET_NOLOGO=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
PROJECT=tests/Axyra.Sheets.Examples.Tests
dotnet build "$PROJECT" -v:q -nologo || exit 1
EXE=$(ls "$PROJECT"/bin/Debug/net*/Axyra.Sheets.Examples.Tests.dll | head -1)

TRAIT=(-trait- "Category=known-issue")
if [[ "${1:-}" == "--known-issues" ]]; then TRAIT=(-trait "Category=known-issue"); shift; fi

if [[ $# -gt 0 ]]; then
  CLASSES=("$@")
else
  CLASSES=($(dotnet "$EXE" -list classes "${TRAIT[@]}" 2>/dev/null | grep -E '^Axyra\.' | sort -u))
fi

if [[ ${#CLASSES[@]} -eq 0 ]]; then
  echo "No test classes discovered." >&2
  exit 1
fi

failed=()
for cls in "${CLASSES[@]}"; do
  [[ "$cls" == *.* ]] || cls=$(dotnet "$EXE" -list classes 2>/dev/null | grep -E "\.${cls}\$" | head -1)
  if [[ -z "$cls" ]]; then echo "Unknown test class." >&2; exit 1; fi
  echo "=== $cls"
  if ! dotnet "$EXE" -class "$cls" "${TRAIT[@]}" -noLogo -noColor; then failed+=("$cls"); fi
done
echo
if [[ ${#failed[@]} -eq 0 ]]; then echo "All ${#CLASSES[@]} test classes passed."; else
  echo "FAILED classes (${#failed[@]}/${#CLASSES[@]}):"; printf '  %s\n' "${failed[@]}"; exit 1; fi
