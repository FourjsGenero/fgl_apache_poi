IMPORT FGL fgl_excel

MAIN
DEFINE workbook     fgl_excel.workbookType 
DEFINE sheet        fgl_excel.sheetType  
DEFINE row          fgl_excel.rowType  
DEFINE cell         fgl_excel.cellType 

DEFINE i INTEGER
    
    -- create workbook
    CALL fgl_excel.workbook_create() RETURNING workbook

    -- SIMPLE test
    
    -- create a worksheet
    CALL fgl_excel.workbook_createsheet(workbook) RETURNING sheet
    -- create row
    CALL fgl_excel.sheet_createrow(sheet, 0) RETURNING row

    CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
    CALL fgl_excel.cell_number_set(cell,-1000000 )

    CALL fgl_excel.row_createcell(row, column2row("B")) RETURNING cell
    CALL fgl_excel.cell_number_set(cell,550000 )

    CALL fgl_excel.row_createcell(row, column2row("C")) RETURNING cell
    CALL fgl_excel.cell_number_set(cell,550000 )

    CALL fgl_excel.row_createcell(row, column2row("D")) RETURNING cell
    CALL fgl_excel.cell_formula_set(cell,"IRR(A1:C1,0.1)")  

    DISPLAY SFMT("Value in cell D1 = %1",  workbook.getCreationHelper().createFormulaEvaluator().evaluate(cell).getNumberValue())


    -- COMPLEX test
    CALL fgl_excel.workbook_createsheet(workbook) RETURNING sheet
    CALL fgl_excel.sheet_createrow(sheet, 0) RETURNING row

    CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
    CALL fgl_excel.cell_number_set(cell,-1000000 )

    FOR i = 1 TO 360
        CALL fgl_excel.sheet_createrow(sheet, i) RETURNING row
        CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
        CALL fgl_excel.cell_number_set(cell,8775.72 )
    END FOR
    CALL fgl_excel.sheet_createrow(sheet, i) RETURNING row
    CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
    CALL fgl_excel.cell_formula_set(cell,SFMT("IRR(A1:A%1,0.01)", i USING "<<<<"))  
    DISPLAY SFMT("Value in cell A362 = %1",  12*workbook.getCreationHelper().createFormulaEvaluator().evaluate(cell).getNumberValue())
END MAIN

    
    
