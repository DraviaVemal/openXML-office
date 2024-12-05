-- query : create_number_format_table# 
CREATE TABLE
    number_formats (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        format_id INTEGER NOT NULL UNIQUE, -- Number Format code
        format_code TEXT NOT NULL -- Number Format code
    );

-- query : create_font_style_table# 
CREATE TABLE
    font_style (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        font_name TEXT NOT NULL, -- Style Name
        color_type TEXT, -- Color Setting
        color_value TEXT, -- Color Setting
        family_id INTEGER, -- Font Family Id
        font_size INTEGER, -- Font Size
        font_scheme TEXT, -- Font Schema
        is_bold INTEGER, -- Is Bold BOOL
        is_italic INTEGER, -- Is Italic BOOL
        is_underline INTEGER, -- Is Underline BOOL
        is_double_underline INTEGER -- Is Double Underline BOOL
    );

-- query : create_fill_style_table# 
CREATE TABLE
    fill_style (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        background_color_setting TEXT, -- Background Color Setting JSON
        foreground_color_setting TEXT, -- Foreground Color Setting JSON
        pattern_type TEXT -- Pattern Type Enum
    );

-- query : create_border_style_table# 
CREATE TABLE
    border_style (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        left_border TEXT, -- Left border setting JSON
        top_border TEXT, -- Top border setting JSON
        right_border TEXT, -- Right border setting JSON
        bottom_border TEXT, -- Bottom border setting JSON
        diagonal_border TEXT -- Bottom border setting JSON
    );

-- query : create_cell_style_table# 
CREATE TABLE
    cell_xfs (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        format_id INTEGER, -- Format Id
        number_format_id INTEGER, -- Number Format Id
        font_id INTEGER, -- Font Id
        fill_id INTEGER, -- Fill Id
        border_id INTEGER, -- Border Id
        is_apply_font INTEGER, -- Apply font flag BOOL
        is_apply_alignment INTEGER, -- Apply alignment flag BOOL
        is_apply_fill INTEGER, -- Apply fill flag BOOL
        is_apply_border INTEGER, -- Apply border flag BOOL
        is_apply_number_format INTEGER, -- Apply number format flag BOOL
        is_apply_protection INTEGER, -- Apply protection flag BOOL
        is_wrap_text INTEGER, -- Wrap cell text flag BOOL
        horizontal_alignment TEXT, -- Horizontal alignment setting JSON
        vertical_alignment TEXT -- Vertical alignment setting JSON
    );

-- query : insert_number_format_table# 
INSERT INTO
    number_formats (format_code, format_id)
VALUES
    (?, ?);

-- query : insert_font_style_table# 
INSERT INTO
    font_style (
        font_name,
        color_type,
        color_value,
        family_id,
        font_size,
        font_scheme,
        is_bold,
        is_italic,
        is_underline,
        is_double_underline
    )
VALUES
    (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);

-- query : insert_fill_style_table# 
INSERT INTO
    fill_style (
        background_color_setting,
        foreground_color_setting,
        pattern_type
    )
VALUES
    (?, ?, ?);

-- query : insert_border_style_table# 
INSERT INTO
    border_style (
        left_border,
        top_border,
        right_border,
        bottom_border,
        diagonal_border
    )
VALUES
    (?, ?, ?, ?, ?);

-- query : insert_cell_style_table# 
INSERT INTO
    cell_xfs (
        format_id,
        number_format_id,
        font_id,
        fill_id,
        border_id,
        is_apply_font,
        is_apply_alignment,
        is_apply_fill,
        is_apply_border,
        is_apply_number_format,
        is_apply_protection,
        is_wrap_text,
        horizontal_alignment,
        vertical_alignment
    )
VALUES
    (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);

-- query : select_share_string_table# 
SELECT
    format_code,
    format_id
FROM
    number_formats
ORDER BY
    id;

-- query : select_share_string_table# 
SELECT
    font_name,
    color_theme,
    color_rgb,
    family_id,
    font_size,
    font_scheme,
    is_bold,
    is_italic,
    is_underline,
    is_double_underline
FROM
    font_style
ORDER BY
    id;

-- query : select_share_string_table# 
SELECT
    background_color_setting,
    foreground_color_setting,
    pattern_type
FROM
    fill_style
ORDER BY
    id;

-- query : select_share_string_table# 
SELECT
    left_border,
    top_border,
    right_border,
    bottom_border,
    diagonal_border
FROM
    border_style
ORDER BY
    id;

-- query : select_cell_style_table# 
SELECT
    format_id,
    number_format_id,
    font_id,
    fill_id,
    border_id,
    is_apply_font,
    is_apply_alignment,
    is_apply_fill,
    is_apply_border,
    is_apply_number_format,
    is_apply_protection,
    is_wrap_text,
    horizontal_alignment,
    vertical_alignment
FROM
    cell_xfs
ORDER BY
    id;