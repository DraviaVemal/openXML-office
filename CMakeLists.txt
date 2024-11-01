cmake_minimum_required(VERSION 3.15)

project(OpenXxlOffice)

# Master Variables
set(ROOT_PATH ${CMAKE_CURRENT_LIST_DIR})
set(PROTO_PATH "proto")
set(RUST_PATH "rs")
set(C_SHARP_PATH "cs")
set(GO_PATH "go")
set(JAVA_PATH "java")
set(GLOBAL_PATH "global")
set(SPREADSHEET_PATH "spreadsheet")
set(PRESENTATION_PATH "presentation")
set(DOCUMENT_PATH "document")

# Step 1: Prepare the proto buf structs
message("Proto Buffer build Start")

set(PROTOC_COMMANDS)
set(PROTOC_OUTPUTS)

# Glob recursively for all .proto files in the specified PROTO_PATH
file(GLOB_RECURSE PROTO_FILES "${PROTO_PATH}/*.proto")

# Iterate over each found .proto file
foreach(PROTO_FILE ${PROTO_FILES})
    # Get the relative path of the .proto file
    file(RELATIVE_PATH RELATIVE_PROTO_FILE ${ROOT_PATH}/proto ${PROTO_FILE})

    # Extract the directory name of the .proto file
    get_filename_component(DIRECTORY_NAME ${RELATIVE_PROTO_FILE} DIRECTORY)

    # Extract the file name of the .proto file
    get_filename_component(FILE_NAME ${RELATIVE_PROTO_FILE} NAME)

    # Extract the file name without extension
    get_filename_component(FILE_NAME_NO_EXTENSION ${RELATIVE_PROTO_FILE} NAME_WE)

    # Replace '/' with ';' in the directory name for list processing
    string(REPLACE "/" ";" DIR_LIST ${DIRECTORY_NAME})

    # Get the first directory from the list
    list(GET DIR_LIST 0 FIRST_DIR)

    # Create the remaining path after the first directory
    string(REPLACE "${FIRST_DIR}/" "" REST_PATH ${DIRECTORY_NAME})

    # Set the output directory for Rust files
    set(OUTPUT_RUST_DIR "${RUST_PATH}/${FIRST_DIR}/src/schema/${REST_PATH}")

    # Get the first letter of the first directory
    string(SUBSTRING "${FIRST_DIR}" 0 1 FIRST_LETTER)

    # Convert the first letter to uppercase
    string(TOUPPER "${FIRST_LETTER}" FIRST_LETTER_UPPER)

    # Get the remaining string after the first letter
    string(SUBSTRING "${FIRST_DIR}" 1 -1 REST_STRING)

    # Combine the uppercase first letter with the remaining string
    set(UPDATED_STRING "${FIRST_LETTER_UPPER}${REST_STRING}")

    # Set the output directory for C# files
    set(OUTPUT_C_SHARP_DIR "${C_SHARP_PATH}/${UPDATED_STRING}/schema/${REST_PATH}")

    # Set the output directory for Go files
    set(OUTPUT_GO_DIR "${GO_PATH}/${FIRST_DIR}")

    # Set the output directory for Java files
    set(OUTPUT_JAVA_DIR "${JAVA_PATH}/${FIRST_DIR}/src/schema/${REST_PATH}")

    # Set the output file path for the Rust generated file
    set(OUTPUT_FILE "${OUTPUT_RUST_DIR}/${FILE_NAME_NO_EXTENSION}.rs")

    # Create the necessary output directories
    file(MAKE_DIRECTORY "${OUTPUT_RUST_DIR}")
    file(MAKE_DIRECTORY "${OUTPUT_C_SHARP_DIR}")
    file(MAKE_DIRECTORY "${OUTPUT_GO_DIR}")
    file(MAKE_DIRECTORY "${OUTPUT_JAVA_DIR}")

    # Append protoc command
    list(APPEND PROTOC_COMMANDS "protoc --proto_path=. --rs_out=${OUTPUT_RUST_DIR} --csharp_out=${OUTPUT_C_SHARP_DIR} --go_out=${OUTPUT_GO_DIR} --java_out=${OUTPUT_JAVA_DIR} proto/${DIRECTORY_NAME}/*.proto")
    list(APPEND PROTOC_OUTPUTS "proto/${DIRECTORY_NAME}/*.proto")
endforeach()

set(UNIQUE_COMMANDS)
set(UNIQUE_OUTPUTS)

set(CURRENT_INDEX 0)
foreach(CMD ${PROTOC_COMMANDS})
    list(FIND UNIQUE_COMMANDS "${CMD}" INDEX)
    if(INDEX EQUAL -1)
        list(GET PROTOC_OUTPUTS ${CURRENT_INDEX} OUTPUT)
        list(APPEND UNIQUE_COMMANDS "${CMD}")
        list(APPEND UNIQUE_OUTPUTS "${OUTPUT}")
    endif()
    math(EXPR CURRENT_INDEX "${CURRENT_INDEX} + 1")
endforeach()

set(CURRENT_INDEX 0)
foreach(CMD ${UNIQUE_COMMANDS})
    list(GET UNIQUE_OUTPUTS ${CURRENT_INDEX} OUTPUT)
    add_custom_command(
        OUTPUT "pwd && ${OUTPUT}"
        COMMAND ${CMD}
        COMMENT "Executing: ${CMD}"
        VERBATIM
    )
    # Add a custom target that depends on this command
    add_custom_target("generate_protobuf_${CURRENT_INDEX}" ALL DEPENDS "${OUTPUT}")
    math(EXPR CURRENT_INDEX "${CURRENT_INDEX} + 1")
endforeach()

message("Proto Buffer build Complete")


