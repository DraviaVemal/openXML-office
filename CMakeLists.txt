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


message("Proto Buffer build Start")

# Step 1: Prepare the Rust structs
message("Proto Buffer build Rust Proto Preparation")

file(GLOB_RECURSE PROTO_FILES "${PROTO_PATH}/*.proto")
foreach(PROTO_FILE ${PROTO_FILES})
    # Get Relative File Path
    file(RELATIVE_PATH RELATIVE_PROTO_FILE ${ROOT_PATH}/proto ${PROTO_FILE})
    # Extract just the file name
    get_filename_component(DIRECTORY_NAME ${RELATIVE_PROTO_FILE} DIRECTORY)
    get_filename_component(FILE_NAME ${RELATIVE_PROTO_FILE} NAME)
    get_filename_component(FILE_NAME_NO_EXTENSION ${RELATIVE_PROTO_FILE} NAME_WE)
    message(STATUS "Proto File Found : proto/${RELATIVE_PROTO_FILE}")
    string(REPLACE "/" ";" DIR_LIST ${DIRECTORY_NAME})
    list(GET DIR_LIST 0 FIRST_DIR)
    string(REPLACE "${FIRST_DIR}/" "" REST_PATH ${DIRECTORY_NAME})
    set(OUTPUT_DIR "${FIRST_DIR}/src/schema/${REST_PATH}")
    set(OUTPUT_FILE "${OUTPUT_DIR}/${FILE_NAME_NO_EXTENSION}.rs")
    # Ensure the output directory exists
    file(MAKE_DIRECTORY "${OUTPUT_DIR}")
    message(STATUS "Rust Output File  : ${OUTPUT_FILE}")
    message(STATUS "protoc --proto_path=proto/${DIRECTORY_NAME} --rs_out=${OUTPUT_DIR} ${FILE_NAME}")
    add_custom_command(
        OUTPUT ${OUTPUT_FILE}
        COMMAND protoc --proto_path=proto/${DIRECTORY_NAME} --rs_out=${OUTPUT_DIR} ${FILE_NAME}
        DEPENDS "proto/${RELATIVE_PROTO_FILE}"
        COMMENT "Generating code for ${FILE_NAME}"
        COMMAND_ERROR_IS_FATAL ANY
    )
    add_custom_target(generate_${FILE_NAME} DEPENDS ${OUTPUT_FILE})
endforeach()
# Step 2: Prepare the C# Class
message("Proto Buffer build C# Proto Preparation")
# Step 3: Prepare the Go interface
message("Proto Buffer build Go Proto Preparation")
# Step 4: Prepare the Java Class
message("Proto Buffer build Java Proto Preparation")
# Step 5: Copy schema's to targets
message("Proto Buffer copy results to target directory")

message("Proto Buffer build Complete")


message("Rust build Start")
message("Rust build Complete")
message("CSharp build Start")
message("CSharp build Complete")
message("Gol lang build Start")
message("Gol lang build Complete")
message("Java build Start")
message("Java build Complete")

