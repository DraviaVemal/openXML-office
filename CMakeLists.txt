cmake_minimum_required(VERSION 3.15)
project(RustMultiProjectBuilder)

# Step 1: Define paths and directory structure.
set(PROTO_DIR ${CMAKE_SOURCE_DIR}/proto)            # Root directory for proto files.
set(RS_PROJECT ${CMAKE_SOURCE_DIR}/rs)              # Rust project root
set(CS_PROJECT ${CMAKE_SOURCE_DIR}/cs)              # CS project root
set(GO_PROJECT ${CMAKE_SOURCE_DIR}/go)              # GO project root
set(JAVA_PROJECT ${CMAKE_SOURCE_DIR}/java)          # Java project root
set(BINARY_DIR ${CMAKE_BINARY_DIR}/output)          # Output directory for final artifacts.

# Step 2: Create necessary directories.
file(MAKE_DIRECTORY ${BINARY_DIR})

# Step 3: Define Rust projects and their build order.
set(RUST_PROJECTS 
    "xml"
    "global"
    "spreadsheet"
    "presentation"
    "document"
    "ffi"
)

# Step 4: Clean the schema directory before regenerating Protobufs.
function(clean_schema_dir project_name)
    set(SCHEMA_DIR ${RS_PROJECT}/${project_name}/schema)
    if(EXISTS ${SCHEMA_DIR})
        file(REMOVE_RECURSE ${SCHEMA_DIR})
    endif()
    file(MAKE_DIRECTORY ${SCHEMA_DIR})
endfunction()

# Step 5: Generate Rust structs from Protobufs for each project.
foreach(project ${RUST_PROJECTS})
    clean_schema_dir(${project})  # Clean schema directory before generating new code.

    set(PROJECT_PROTO_DIR ${PROTO_DIR}/${project})   # Path to proto files.
    set(GENERATED_DIR ${RS_PROJECT}/${project}/schema)  # Output schema dir.

    add_custom_target(
        GenerateProtos_${project} ALL
        COMMAND protoc --rust_out=${GENERATED_DIR} --proto_path=${PROJECT_PROTO_DIR} ${PROJECT_PROTO_DIR}/*.proto
        COMMENT "Generating Rust code from ${PROJECT_PROTO_DIR}/*.proto into ${GENERATED_DIR}"
    )
endforeach()

# Step 6: Macro to build Rust projects using Cargo.
macro(build_rust_project project_name)
    add_custom_target(
        ${project_name}_build ALL
        COMMAND cargo build --manifest-path ${CMAKE_SOURCE_DIR}/rs/${project_name}/Cargo.toml --release
        COMMENT "Building Rust project: ${project_name}"
        DEPENDS GenerateProtos_${project_name}
    )
endmacro()

# Step 7: Build all Rust projects in the specified order.
foreach(project ${RUST_PROJECTS})
    build_rust_project(${project})
endforeach()

# Step 8: Install the build artifacts to the output directory.
foreach(project ${RUST_PROJECTS})
    install(DIRECTORY ${CMAKE_SOURCE_DIR}/rs/${project}/target/release/
            DESTINATION ${BINARY_DIR}/${project})
endforeach()

# Step 9: Completion message.
add_custom_target(
    CompletionMessage ALL
    COMMAND ${CMAKE_COMMAND} -E echo "All Rust projects built successfully!"
    COMMENT "Build completed."
)
