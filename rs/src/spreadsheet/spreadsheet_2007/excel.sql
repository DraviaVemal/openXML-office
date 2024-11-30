-- query : create_number_format_table# Create initial blob archive table
CREATE TABLE
    number_formats (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        format_code TEXT NOT NULL -- Number Format code
    );

-- query : create_font_style_table# Create initial blob archive table
CREATE TABLE
    font_style (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        name TEXT NOT NULL, -- Style Name
        color_setting TEXT, -- Color Setting JSON
        family_id INTEGER, -- Font Family Id
        font_size INTEGER, -- Font Size
        font_scheme TEXT, -- Font Schema JSON
        is_bold INTEGER, -- Is Bold BOOL
        is_italic INTEGER, -- Is Italic BOOL
        is_underline INTEGER, -- Is Underline BOOL
        is_double_underline INTEGER -- Is Double Underline BOOL
    );

-- query : create_fill_style_table# Create initial blob archive table
CREATE TABLE
    fill_style (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        background_color_setting TEXT, -- Background Color Setting JSON
        foreground_color_setting TEXT, -- Foreground Color Setting JSON
        pattern_type TEXT -- Pattern Type JSON
    );

-- query : create_border_style_table# Create initial blob archive table
CREATE TABLE
    border_style (
        id INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
        left_border TEXT, -- Left border setting JSON
        top_border TEXT, -- Top border setting JSON
        right_border TEXT, -- Right border setting JSON
        bottom_border TEXT -- Bottom border setting JSON
    );

-- query : create_cell_style_table# Create initial blob archive table
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