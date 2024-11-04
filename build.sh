# Define the source and target directories
SOURCE_DIR="fbs"
RUST_DIR="rs/fbs/src"
C_SHARP_DIR="cs/Fbs"
GO_DIR="go/fbs/src"
JAVA_DIR="java/fbs/src/main/java/com/draviavemal"

rm -rf "$C_SHARP_DIR/openxmloffice"
rm -rf "$GO_DIR/openxmloffice"
rm -rf "$JAVA_DIR/openxmloffice"

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
  flatc -r --gen-all -o "$RUST_DIR" "fbs/consolidated.fbs"
