-- query : create_dynamic_sheet# 
CREATE TABLE
    {} (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        row_index INTEGER NOT NULL,
        col_index INTEGER NOT NULL,
        cell_style_id INTEGER,
        cell_value TEXT,
        cell_formula TEXT,
        cell_type TEXT,
        cell_metadata TEXT,
        cell_place_holder TEXT,
        cell_comment_id INTEGER,
        UNIQUE (row_index, col_index)
    );

-- query : create_dynamic_sheet_row# 
CREATE TABLE
    {} (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        row_index INTEGER NOT NULL UNIQUE,
        row_hide INTEGER,
        row_span TEXT,
        row_height INTEGER,
        row_style_id INTEGER,
        row_thick_top INTEGER,
        row_thick_bottom INTEGER,
        row_group_level INTEGER,
        row_collapsed INTEGER,
        row_place_holder TEXt
    );

-- query : insert_dynamic_sheet# 
INSERT INTO
    {} (
        row_index,
        col_index,
        cell_style_id,
        cell_value,
        cell_formula,
        cell_type,
        cell_metadata,
        cell_place_holder,
        cell_comment_id
    )
VALUES
    (
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?
    );

-- query : insert_conflict_dynamic_sheet# 
INSERT INTO
    {} (
        row_index,
        col_index,
        cell_style_id,
        cell_value,
        cell_formula,
        cell_type,
        cell_metadata,
        cell_place_holder,
        cell_comment_id
    )
VALUES
    (
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?
    ) ON CONFLICT (row_index, col_index) DO
UPDATE
SET
    cell_style_id = excluded.cell_style_id,
    cell_value = excluded.cell_value,
    cell_formula = excluded.cell_formula,
    cell_type = excluded.cell_type,
    cell_metadata = excluded.cell_metadata,
    cell_place_holder = excluded.cell_place_holder,
    cell_comment_id = excluded.cell_comment_id;

-- query : insert_dynamic_sheet_row# 
INSERT INTO
    {} (
        row_index,
        row_hide,
        row_span,
        row_height,
        row_style_id,
        row_thick_top,
        row_thick_bottom,
        row_group_level,
        row_collapsed,
        row_place_holder
    )
VALUES
    (
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?
    );

-- query : insert_conflict_dynamic_sheet_row# 
INSERT INTO
    {} (
        row_index,
        row_hide,
        row_span,
        row_height,
        row_style_id,
        row_thick_top,
        row_thick_bottom,
        row_group_level,
        row_collapsed,
        row_place_holder
    )
VALUES
    (
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?,
        ?
    ) ON CONFLICT (row_index) DO
UPDATE
SET
    row_hide = excluded.row_hide,
    row_span = excluded.row_span,
    row_height = excluded.row_height,
    row_style_id = excluded.row_style_id,
    row_thick_top = excluded.row_thick_top,
    row_thick_bottom = excluded.row_thick_bottom,
    row_group_level = excluded.row_group_level,
    row_collapsed = excluded.row_collapsed,
    row_place_holder = excluded.row_place_holder;

-- query : delete_row_record# 

DELETE FROM {}
WHERE row_index = ?;

-- query : select_all_dynamic_sheet# 
SELECT
    a.row_index,
    a.row_hide,
    a.row_span,
    a.row_height,
    a.row_style_id,
    a.row_thick_top,
    a.row_thick_bottom,
    a.row_group_level,
    a.row_collapsed,
    a.row_place_holder,
    b.col_index,
    b.cell_style_id,
    b.cell_value,
    b.cell_formula,
    b.cell_type,
    b.cell_metadata,
    b.cell_place_holder,
    b.cell_comment_id
FROM
    {0} as a
LEFT JOIN
    {} as b
ON
    a.row_index = b.row_index
ORDER BY
    a.row_index DESC,
    b.col_index DESC;

-- query : select_range_dynamic_sheet# 
SELECT
    *
FROM
    {}
WHERE
    row_index >= ?
    AND col_index >= ?
    AND row_index <= ?
    AND col_index <= ?
ORDER BY
    row_index DESC,
    col_index DESC;


-- query : select_dimension_dynamic_sheet# 
SELECT 
    MIN(row_index) as start_row_id,
    MIN(col_index) as start_col_id,
    MAX(row_index) as end_row_id,
    MAX(col_index) as end_col_id
FROM
    {};