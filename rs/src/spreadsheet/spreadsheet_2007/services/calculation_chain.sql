-- query : create_calculation_chain_table# 
CREATE TABLE
    calculation_chain
(
    id       INTEGER PRIMARY KEY AUTOINCREMENT, -- Unique ID for each file
    cell     TEXT NOT NULL UNIQUE,              -- Excel Cell ID String
    sheet_id TEXT NOT NULL                      -- Sheet of the Cell String
);

-- query : insert_calculation_chain_table# 
INSERT INTO calculation_chain (cell, -- Excel Cell ID String
                               sheet_id -- Sheet of the Cell String
)
VALUES (?, ?);

-- query : select_calculation_chain_table# 
SELECT cell,
       sheet_id
FROM calculation_chain
ORDER BY id;