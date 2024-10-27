-- query : #create_archive_table# Create initial blob archive table
CREATE TABLE
    archive (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        file_name TEXT NOT NULL, -- Name of the file including directory
        compressed_file_size INTEGER, -- Size of compressed file in bytes
        uncompressed_file_size INTEGER, -- Size of uncompressed file in bytes
        compression_level INTEGER NOT NULL, -- File Compression level can be adjusted to adjust CPU load
        compression_type TEXT NOT NULL, -- File Compression type
        content BLOB -- File content as a BLOB
    );