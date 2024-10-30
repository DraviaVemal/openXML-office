-- query : select_workbook# select and pull workbook blob content
SELECT
    content
FROM
    archive
WHERE
    file_name = ?;

-- query : insert_workbook# select and pull workbook blob content
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
    (? , ?, ? , ? , ? , ? );