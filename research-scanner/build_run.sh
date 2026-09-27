#!/usr/bin/env bash
set -e

PAYLOAD_DIR="payload"
ENTRYPOINT="entrypoint.sh"
OUTPUT_RUN="scan_research_data.run"

if ! command -v makeself &>/dev/null; then
    echo "Error: 'makeself' utility is not installed."
    echo "Install it on macOS via: brew install makeself"
    exit 1
fi

chmod +x "$PAYLOAD_DIR/$ENTRYPOINT"

echo "Building self-extracting archive using makeself..."

# Pass --format=posix (or --numeric-owner) to tar via makeself's --tar-extra flag
makeself --sha256 --nomd5 --nocrc --tar-format posix "$PAYLOAD_DIR" "$OUTPUT_RUN" "Archaeological Literature Data Extractor" ./"$ENTRYPOINT"

echo "Success! Archive built: $OUTPUT_RUN"
