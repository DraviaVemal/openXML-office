#!/bin/bash

# Default to an empty string (no --release)
release_flag=""
win_binary_dir="target/x86_64-pc-windows-gnu/debug"
linux_binary_dir="target/x86_64-unknown-linux-gnu/debug"

# Check if --release is passed as an argument
for arg in "$@"; do
  if [ "$arg" == "--release" ]; then
    release_flag="--release"
    win_binary_dir="target/x86_64-pc-windows-gnu/release"
    linux_binary_dir="target/x86_64-unknown-linux-gnu/release"
    break
  fi
done

# Define the source and target directories
SOURCE_DIR="fbs"
RUST_FFI_DIR="rs-ffl/fbs/src"
C_SHARP_DIR="cs/Fbs"
GO_DIR="go/fbs/src"
JAVA_DIR="java/fbs/src/main/java/com/draviavemal"

# Flatbuffer Generated code setup
rm -rf "$C_SHARP_DIR/openxmloffice_fbs"
rm -rf "$GO_DIR/openxmloffice_fbs"
rm -rf "$JAVA_DIR/openxmloffice_fbs"

# Find and compile each .fbs file
find "$SOURCE_DIR" -name "*.fbs" | while read -r fbs_file; do
  # Get the directory of the .fbs file relative to SOURCE_DIR
  relative_dir=$(dirname "$fbs_file" | sed "s|$SOURCE_DIR||")

  # Create the corresponding output directory in TARGET_DIR
  mkdir -p "$C_SHARP_DIR"
  mkdir -p "$GO_DIR"
  mkdir -p "$JAVA_DIR"

  # Compile the .fbs file with flatc, targeting the appropriate directory
  flatc -n -o "$C_SHARP_DIR" "$fbs_file"
  flatc -g -o "$GO_DIR" "$fbs_file"
  flatc -j -o "$JAVA_DIR" "$fbs_file"
done
# Rust code get complied as one source to maintain the modular hierarchy
flatc -r --gen-all -o "$RUST_FFI_DIR" "fbs/consolidated.fbs"

# Build Rust Core Libraries

# Build Rust dynamic link file

cd rs-ffl

# Prepare Build Result Directory
rm -rf ../cs/Spreadsheet/Lib && mkdir -p ../cs/Spreadsheet/Lib
rm -rf ../cs/Presentation/Lib && mkdir -p ../cs/Presentation/Lib
rm -rf ../cs/Document/Lib && mkdir -p ../cs/Document/Lib

# Cargo Build FFI binary
# Clear build history
cargo clean

# Windows
cargo build $release_flag --target x86_64-pc-windows-gnu

# Linux
cargo build $release_flag --target x86_64-unknown-linux-gnu

# Mac osX
# cargo build --release --target x86_64-apple-darwin

# Copy Result binary to targets
cp $win_binary_dir/openxmloffice_ffi.dll ../cs/Spreadsheet/Lib/openxmloffice_ffi.dll
cp $win_binary_dir/openxmloffice_ffi.dll ../cs/Presentation/Lib/openxmloffice_ffi.dll
cp $win_binary_dir/openxmloffice_ffi.dll ../cs/Document/Lib/openxmloffice_ffi.dll
cp $linux_binary_dir/libopenxmloffice_ffi.so ../cs/Spreadsheet/Lib/openxmloffice_ffi.so
cp $linux_binary_dir/libopenxmloffice_ffi.so ../cs/Presentation/Lib/openxmloffice_ffi.so
cp $linux_binary_dir/libopenxmloffice_ffi.so ../cs/Document/Lib/openxmloffice_ffi.so

cd ..

# Build wrapper library using link files

# C# Project Build

cd cs

dotnet build

cd ..
