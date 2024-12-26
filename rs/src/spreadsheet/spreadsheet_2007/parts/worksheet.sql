-- query : create_dynamic_sheet# 
CREATE TABLE
    {} (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        row_id INTEGER NOT NULL,
        row_hide INTEGER,
        row_span TEXT,
        row_height INTEGER,
        row_style_id INTEGER,
        row_thick_top INTEGER,
        row_thick_bottom INTEGER,
        row_group_level INTEGER,
        row_collapsed INTEGER,
        row_place_holder TEXt,
        col_id INTEGER NOT NULL,
        col_style_id INTEGER,
        cell_value TEXT,
        cell_formula TEXT,
        cell_type TEXT,
        cell_metadata TEXT,
        cell_place_holder TEXT,
        cell_comment_id INTEGER,
        UNIQUE (row_id, col_id)
    );

-- query : insert_dynamic_sheet# 
INSERT INTO
    {} (
        row_id,
        row_hide,
        row_span,
        row_height,
        row_style_id,
        row_thick_top,
        row_thick_bottom,
        row_group_level,
        row_collapsed,
        row_place_holder,
        col_id,
        col_style_id,
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
    ) ON CONFLICT (row_id, col_id) DO
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
    row_place_holder = excluded.row_place_holder,
    col_style_id = excluded.col_style_id,
    cell_value = excluded.cell_value,
    cell_formula = excluded.cell_formula,
    cell_type = excluded.cell_type,
    cell_metadata = excluded.cell_metadata,
    cell_place_holder = excluded.cell_place_holder,
    cell_comment_id = excluded.cell_comment_id;

-- query : select_all_dynamic_sheet# 
SELECT
    row_id,
    row_hide,
    row_span,
    row_height,
    row_style_id,
    row_thick_top,
    row_thick_bottom,
    row_group_level,
    row_collapsed,
    row_place_holder,
    col_id,
    col_style_id,
    cell_value,
    cell_formula,
    cell_type,
    cell_metadata,
    cell_place_holder,
    cell_comment_id
FROM
    {}
ORDER BY
    row_id DESC,
    col_id DESC;

-- query : select_range_dynamic_sheet# 
SELECT
    row_id,
    row_hide,
    row_span,
    row_height,
    row_style_id,
    row_thick_top,
    row_thick_bottom,
    row_group_level,
    row_collapsed,
    row_place_holder,
    col_id,
    col_style_id,
    cell_value,
    cell_formula,
    cell_type,
    cell_metadata,
    cell_place_holder,
    cell_comment_id
FROM
    {}
WHERE
    row_id >= ?
    AND col_id >= ?
    AND row_id <= ?
    AND col_id <= ?
ORDER BY
    row_id DESC,
    col_id DESC;
