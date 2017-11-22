IMPORT FGL fgl_excel

MAIN
DEFINE sql STRING
DEFINE filename STRING
DEFINE header BOOLEAN
DEFINE preview BOOLEAN
DEFINE result BOOLEAN

    DEFER INTERRUPT
    DEFER QUIT
    OPTIONS INPUT WRAP
    OPTIONS FIELD ORDER FORM

    -- Create and populate test database
    CONNECT TO ":memory:+driver='dbmsqt'"
    CALL populate()
    
    OPEN WINDOW w WITH FORM "fgl_excel_generic_test"

    -- Default values
    LET sql = "SELECT * FROM test_data"
    LET filename = "fgl_excel_generic_test.xlsx"
    LET header = TRUE
    LET preview = TRUE
    
    INPUT BY NAME sql, filename, header, preview ATTRIBUTES(UNBUFFERED, WITHOUT DEFAULTS=TRUE, ACCEPT=FALSE, CANCEL=FALSE)
        ON ACTION excel ATTRIBUTES(TEXT="Generate Excel", IMAGE="fa-file-excel-o")
            IF sql_to_excel(sql, filename, header) THEN
                IF preview THEN
                    CALL fgl_putfile(filename, filename)
                    CALL ui.Interface.frontCall("standard","shellExec", filename, result)
                ELSE
                    MESSAGE "Spreadsheet created"
                END IF
            ELSE
                ERROR "Something went wrong"
            END IF
        
        ON ACTION close
            EXIT INPUT

    END INPUT
END MAIN


FUNCTION sql_to_excel(sql, filename, header)
DEFINE hdl base.SqlHandle
DEFINE sql STRING
DEFINE filename STRING
DEFINE header BOOLEAN
DEFINE row_idx, col_idx INTEGER 

DEFINE workbook     fgl_excel.workbookType 
DEFINE sheet        fgl_excel.sheetType  
DEFINE row          fgl_excel.rowType  
DEFINE cell         fgl_excel.cellType 
DEFINE header_style fgl_excel.cellStyleType
DEFINE header_font  fgl_excel.fontType

DEFINE datatype STRING
    
    LET hdl = base.SqlHandle.create()
    TRY
        CALL hdl.prepare(sql)
        CALL hdl.open()
    CATCH
        RETURN FALSE
    END TRY

    CALL fgl_excel.workbook_create() RETURNING workbook

    -- create a worksheet
    CALL fgl_excel.workbook_createsheet(workbook) RETURNING sheet

    -- create data rows
    LET row_idx = 0 
    
    WHILE TRUE
        CALL hdl.fetch()
        IF STATUS=NOTFOUND THEN
            EXIT WHILE
        END IF
        LET row_idx = row_idx + 1

        IF row_idx = 1 AND header THEN
            -- create a font, will be used in header
            CALL fgl_excel.font_create(workbook) RETURNING header_font
            CALL fgl_excel.font_set(header_font, "weight", "bold")

            -- create a style, will be used in header
            CALL fgl_excel.style_create(workbook) RETURNING header_style
            CALL fgl_excel.style_set(header_style, "alignment","center")
            CALL fgl_excel.style_font_set(header_style, header_font)
   
            -- Add column headers
            CALL fgl_excel.sheet_createrow(sheet, 0) RETURNING row
            FOR col_idx = 1 TO hdl.getResultCount()
                CALL fgl_excel.row_createcell(row, col_idx-1) RETURNING cell
                CALL fgl_excel.cell_value_set(cell, hdl.getResultName(col_idx))
                CALL fgl_excel.cell_style_set(cell, header_style)
            END FOR
        END IF
        CALL fgl_excel.sheet_createrow(sheet, IIF(header,row_idx, row_idx-1)) RETURNING row

        FOR col_idx = 1 TO hdl.getResultCount()
            CALL fgl_excel.row_createcell(row, col_idx-1) RETURNING cell
            LET datatype = hdl.getResultType(col_idx) 
            CASE 
                WHEN datatype =  "INTEGER" -- TODO check logic
                  OR datatype MATCHES "DECIMAL*"
                  OR datatype MATCHES "FLOAT*"
                  OR datatype MATCHES "*INT*"
                    CALL fgl_excel.cell_number_set(cell, hdl.getResultValue(col_idx))
                OTHERWISE
                    CALL fgl_excel.cell_value_set(cell, hdl.getResultValue(col_idx))
            END CASE
        END FOR
    END WHILE

    -- TODO this code should automatically size the columns
    -- However it is very very slow for reasons I can't determine
    -- Uncomment and test at your leisure
    --IF hdl IS NOT NULL THEN
        --FOR col_idx = 1 TO hdl.getResultCount()
            --CALL fgl_excel.sheet_autosizecolumn(sheet, col_idx-1)
        --END FOR
    --END IF

    -- Write to File
    CALL fgl_excel.workbook_writeToFile(workbook, filename)

    RETURN TRUE   
END FUNCTION



FUNCTION populate()
DEFINE idx INTEGER
DEFINE rec RECORD
    integer_type INTEGER,
    date_type DATE,
    char_type CHAR(20),
    float_type FLOAT
END RECORD

    CREATE TABLE test_data 
        (integer_type INTEGER,
         date_type DATE,
         char_type VARCHAR(20),
         float_type FLOAT)

    FOR idx = 1 TO 26
        LET rec.integer_type = idx
        LET rec.date_type = TODAY+idx
        LET rec.char_type = ASCII(64+idx)
        LET rec.float_type = 1/idx
        
        INSERT INTO test_data VALUES(rec.*)
    END FOR
END FUNCTION
