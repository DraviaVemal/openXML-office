-- query : #select_workbook# select and pull workbook blob content
SELECT
    content
FROM
    archive
WHERE
    file_name = "xl/workbook.xml";