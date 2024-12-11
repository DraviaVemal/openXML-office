-- query : create_dynamic_sheet# 
CREATE TABLE
    ? (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        row_id INTEGER NOT NULL,
        col_id INTEGER NOT NULL,
        cell_placeholder TEXT,
        cell_metadata TEXT,
        cell_value_metadata TEXT,
        cell_formula TEXT,
        cell_value TEXT,
        cell_style INTEGER,
        cell_type TEXT,
    );

-- query : insert_dynamic_sheet# 
INSERT INTO
    ? (
        row_id,
        col_id,
        cell_placeholder,
        cell_metadata,
        cell_value_metadata,
        cell_formula,
        cell_value,
        cell_style,
        cell_type,
    )
VALUES
    (?, ?, ?, ?, ?, ?, ?, ?, ?);

-- query : select_all_dynamic_sheet# 
SELECT
    row_id,
    col_id,
    cell_placeholder,
    cell_metadata,
    cell_value_metadata,
    cell_formula,
    cell_value,
    cell_style,
    cell_type,
FROM
    ?
ORDER BY
    row_id,
    col_id;

-- query : select_range_dynamic_sheet# 
SELECT
    row_id,
    col_id,
    cell_placeholder,
    cell_metadata,
    cell_value_metadata,
    cell_formula,
    cell_value,
    cell_style,
    cell_type,
FROM
    ?
WHERE
    row_id >= ?
    AND col_id >= ?
    AND row_id <= ?
    AND col_id <= ?
ORDER BY
    row_id,
    col_id;