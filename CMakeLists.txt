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
    set(OUTPUT_C_SHARP_DIR "${C_SHARP_PATH}/${UPDATED_STRING}/Schema/${REST_PATH}")

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

    # Add a custom command to generate the code using protoc
    add_custom_command(
        OUTPUT ${OUTPUT_FILE}
        COMMAND protoc --proto_path=proto/${DIRECTORY_NAME} --rs_out=${OUTPUT_RUST_DIR} --csharp_out=${OUTPUT_C_SHARP_DIR} --go_out=${OUTPUT_GO_DIR} --java_out=${OUTPUT_JAVA_DIR} ${FILE_NAME}
        DEPENDS "proto/${RELATIVE_PROTO_FILE}"
        COMMENT "Generating code for ${FILE_NAME}"
        COMMAND_ERROR_IS_FATAL ANY
    )

    # Create a custom target that depends on the generated file
    add_custom_target(generate_${FILE_NAME} DEPENDS ${OUTPUT_FILE})
endforeach()


message("Proto Buffer build Complete")


message("Rust build Start")
message("Rust build Complete")
message("CSharp build Start")
message("CSharp build Complete")
message("Gol lang build Start")
message("Gol lang build Complete")
message("Java build Start")
message("Java build Complete")

