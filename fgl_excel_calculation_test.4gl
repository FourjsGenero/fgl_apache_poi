IMPORT FGL fgl_excel

MAIN
DEFINE workbook     fgl_excel.workbookType 
DEFINE sheet        fgl_excel.sheetType  
DEFINE row          fgl_excel.rowType  
DEFINE cell         fgl_excel.cellType 

    -- create workbook
    CALL fgl_excel.workbook_create() RETURNING workbook

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
   
END MAIN

    
    
