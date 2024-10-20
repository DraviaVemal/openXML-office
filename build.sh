# Build Core components
cargo build  --manifest-path openxmloffice-core/xml/Cargo.toml
cargo build  --manifest-path openxmloffice-core/global/Cargo.toml
cargo build  --manifest-path openxmloffice-core/spreadsheet/Cargo.toml
cargo build  --manifest-path openxmloffice-core/presentation/Cargo.toml
cargo build  --manifest-path openxmloffice-core/document/Cargo.toml
cargo build  --manifest-path openxmloffice-core/efi/Cargo.toml

# Build Rust Wrapper
cargo build  --manifest-path openxmloffice-rs/Cargo.toml

# Build Rust API Container
cargo build  --manifest-path openxmloffice-rs-api/Cargo.toml

# Build C# Wrapper
dotnet build openxmloffice-cs/openXML-Office.sln

# Build Java Wrapper
mvn clean install -f openxmloffice-java/spreadsheet/pom.xml
mvn clean install -f openxmloffice-java/presentation/pom.xml
mvn clean install -f openxmloffice-java/document/pom.xml

# Build Go Wrapper
cd openxmloffice-go/spreadsheet && go build && cd ../..
cd openxmloffice-go/presentation && go build && cd ../..
cd openxmloffice-go/document && go build && cd ../..

# Build TS Wrapper
cd openxmloffice-ts/document && napi build && cd ../..
cd openxmloffice-ts/presentation && napi build && cd ../..
cd openxmloffice-ts/spreadsheet && napi build && cd ../..