-- query : create_archive_table# Create initial blob archive table
CREATE TABLE
    archive (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        file_name TEXT NOT NULL UNIQUE, -- Name of the file including directory
        compressed_file_size INTEGER, -- Size of compressed file in bytes
        uncompressed_file_size INTEGER, -- Size of uncompressed file in bytes
        compression_level INTEGER NOT NULL, -- File Compression level can be adjusted to adjust CPU load
        compression_type TEXT NOT NULL, -- File Compression type
        content BLOB -- File content as a BLOB
    );

-- query : insert_archive_table# Create initial blob archive table
INSERT INTO
    archive (
        file_name,
        compressed_file_size,
        uncompressed_file_size,
        compression_level,
        compression_type,
        content
    )
VALUES
    (?, ?, ?, ?, ?, ?) ON CONFLICT (file_name) DO
UPDATE
SET
    compressed_file_size = excluded.compressed_file_size,
    uncompressed_file_size = excluded.uncompressed_file_size,
    compression_level = excluded.compression_level,
    compression_type = excluded.compression_type,
    content = excluded.content
WHERE
    file_name = excluded.file_name;

-- query : select_all_archive_table# Get All content from archive table
SELECT
    id,
    file_name,
    compressed_file_size,
    uncompressed_file_size,
    compression_level,
    compression_type,
    content
FROM
    archive
ORDER BY
    id;

-- query : select_archive_content# select and pull workbook blob content
SELECT
    content
FROM
    archive
WHERE
    file_name = ?;