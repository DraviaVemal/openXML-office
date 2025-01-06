#!/bin/bash

# Default to an empty string (no --release)
release_flag=""
win_binary_dir="target/x86_64-pc-windows-gnu/debug"
linux_binary_dir="target/x86_64-unknown-linux-gnu/debug"
target_dir="target"

# Check if --release is passed as an argument
for arg in "$@"; do
  if [ "$arg" = "--release" ]; then
    echo "Running Release Build"
    release_flag="--release"
    win_binary_dir="target/x86_64-pc-windows-gnu/release"
    linux_binary_dir="target/x86_64-unknown-linux-gnu/release"
    break
  fi
done

# Define the source and target directories
SOURCE_DIR="fbs"
RUST_FFI_DIR="rs_ffi/src"
C_SHARP_DIR="cs/openxml_office/fbs"
JAVA_DIR="java/draviavemal_openxml_office/src/main/java"
PYTHON_DIR="python"
GO_DIR="go/openxml_office/src"

# Flatbuffer Generated code setup
rm -rf "$C_SHARP_DIR/openxml_office_ffi"
rm -rf "$GO_DIR/openxml_office_ffi"
rm -rf "$JAVA_DIR/openxml_office_ffi"
rm -rf "$PYTHON_DIR/openxml_office_ffi"

# Remove Previous Build Results
rm -rf target

# Find and compile each .fbs file
find "$SOURCE_DIR" -name "*.fbs" | while read -r fbs_file; do
  # Get the directory of the .fbs file relative to SOURCE_DIR
  relative_dir=$(dirname "$fbs_file" | sed "s|$SOURCE_DIR||")

  # Create the corresponding output directory in TARGET_DIR
  mkdir -p "$C_SHARP_DIR"
  mkdir -p "$GO_DIR"
  mkdir -p "$JAVA_DIR/openxml_office_ffi"
  mkdir -p "$PYTHON_DIR/openxml_office_ffi"

  # Compile the .fbs file with flatc, targeting the appropriate directory
  flatc -n -o "$C_SHARP_DIR" "$fbs_file"
  flatc -g -o "$GO_DIR" "$fbs_file"
  flatc -j -o "$JAVA_DIR" "$fbs_file"
done
# Rust code get complied as one source to maintain the modular hierarchy
flatc -p --gen-all -o "$PYTHON_DIR" "fbs/consolidated.fbs"
flatc -r --gen-all -o "$RUST_FFI_DIR" "fbs/consolidated.fbs"

# Build Rust Core Libraries

# Build Rust dynamic link file

# Prepare Build Result Directory
rm -rf cs/openxml_office/lib && mkdir -p cs/openxml_office/lib
rm -rf python/lib && mkdir -p python/lib
rm -rf java/draviavemal_openxml_office/src/lib && mkdir -p java/draviavemal_openxml_office/src/main/resources/lib
rm -rf go/openxml_office/src/lib && mkdir -p go/openxml_office/src/lib

# Cargo Build FFI binary
# Clear build history
cargo clean

# Windows
cargo build $release_flag --target x86_64-pc-windows-gnu

# Linux
cargo build $release_flag --target x86_64-unknown-linux-gnu

# Mac osX
# cargo build --release --target x86_64-apple-darwin

# Copy Result binary to CS targets
cp $win_binary_dir/draviavemal_openxml_office_ffi.dll cs/openxml_office/lib/draviavemal_openxml_office_ffi.dll
cp $linux_binary_dir/libdraviavemal_openxml_office_ffi.so cs/openxml_office/lib/draviavemal_openxml_office_ffi.so

# Copy Result binary to Python targets
cp $win_binary_dir/draviavemal_openxml_office_ffi.dll python/lib/draviavemal_openxml_office_ffi.dll
cp $linux_binary_dir/libdraviavemal_openxml_office_ffi.so python/lib/libdraviavemal_openxml_office_ffi.so
cp $target_dir/bindings.h python/lib/bindings.h

# Copy Result binary to Java targets
cp $win_binary_dir/draviavemal_openxml_office_ffi.dll java/draviavemal_openxml_office/src/main/resources/lib/draviavemal_openxml_office_ffi.dll
cp $linux_binary_dir/libdraviavemal_openxml_office_ffi.so java/draviavemal_openxml_office/src/main/resources/lib/libdraviavemal_openxml_office_ffi.so

# Copy Result binary to Go targets
cp $win_binary_dir/draviavemal_openxml_office_ffi.dll go/openxml_office/src/lib/draviavemal_openxml_office_ffi.dll
cp $linux_binary_dir/libdraviavemal_openxml_office_ffi.so go/openxml_office/src/lib/libdraviavemal_openxml_office_ffi.so

# Build wrapper library using link files

# C# Project Build

cd cs

if [ "$release_flag" = "--release" ]; then
  dotnet build --configuration Release
else
  dotnet build
fi

cd ..

# Python Build
cd python

python3 setup.py bdist_wheel

cd ..

# cd python

# python3 setup.py clean --all

# python3 setup.py build

# cd ..
