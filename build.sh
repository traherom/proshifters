#!/usr/bin/env bash
DIR=$(cd `dirname $0` && pwd)
echo "$DIR"

OUTPUT_DIR="./_output"
mkdir -p "$OUTPUT_DIR" || exit 1

for rid in linux-x64 win-x64; do
    echo "Building $rid"
    cd "$MAIN_WD"
    dotnet publish -c Release --runtime "$rid" --self-contained -o "$OUTPUT_DIR" Proshifters || exit 1
done

cd "$OUTPUT_DIR" || exit 1
rm -rf *.pdb || exit 1
ls -l 