IMPORT FGL fgl_excel
IMPORT util

MAIN
DEFINE workbook     fgl_excel.workbookType 
DEFINE sheet        fgl_excel.sheetType  
DEFINE row          fgl_excel.rowType  
DEFINE cell         fgl_excel.cellType 
DEFINE header_style fgl_excel.cellStyleType
DEFINE header_font  fgl_excel.fontType 

DEFINE i INTEGER

DEFINE result INTEGER

    -- create workbook
    CALL fgl_excel.workbook_create() RETURNING workbook

    -- create a worksheet
    CALL fgl_excel.workbook_createsheet(workbook) RETURNING sheet

    -- add heading row
    
    -- create a font, will be used in header
    CALL fgl_excel.font_create(workbook) RETURNING header_font
    CALL fgl_excel.font_set(header_font, "weight", "bold")

    -- create a style, will be used in header
    CALL fgl_excel.style_create(workbook) RETURNING header_style
    CALL fgl_excel.style_set(header_style, "alignment","center")
    CALL fgl_excel.style_font_set(header_style, header_font)
    CALL fgl_excel.sheet_createrow(sheet, 0) RETURNING row
    
    CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
    CALL fgl_excel.cell_value_set(cell, "Name")
    CALL fgl_excel.cell_style_set(cell, header_style)

    CALL fgl_excel.row_createcell(row, column2row("B")) RETURNING cell
    CALL fgl_excel.cell_value_set(cell, "Qty")
    CALL fgl_excel.cell_style_set(cell, header_style)

    -- create data rows
    FOR i = 1 TO 10
        CALL fgl_excel.sheet_createrow(sheet, i) RETURNING row
        
        CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
        CALL fgl_excel.cell_value_set(cell,SFMT("Item #%1",i))

        CALL fgl_excel.row_createcell(row, column2row("B")) RETURNING cell
        CALL fgl_excel.cell_number_set(cell,util.math.rand(100))
        #CALL cell.
    END FOR

    -- create footer row
    CALL fgl_excel.sheet_createrow(sheet, i) RETURNING row
        
    CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
    CALL fgl_excel.cell_value_set(cell,"Total")

    CALL fgl_excel.row_createcell(row, column2row("B")) RETURNING cell
    CALL fgl_excel.cell_formula_set(cell,"SUM(B2:B11)")  

    CALL fgl_excel.sheet_createrow(sheet, i+1) RETURNING row
    CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
    CALL fgl_excel.cell_value_set(cell,SFMT("This document created on %1 at %2", TODAY, CURRENT HOUR TO SECOND))

    -- Write to File
    CALL fgl_excel.workbook_writeToFile(workbook, "fgl_excel_test.xlsx");
END MAIN