-- query : create_share_string_table# 
CREATE TABLE
    share_string
(
    id     INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
    string TEXT NOT NULL                      -- Common String
);

-- query : insert_share_string_table# 
INSERT INTO share_string (string -- Common String
)
VALUES (?);

-- query : select_share_string_table# 
SELECT string
FROM share_string
ORDER BY id;