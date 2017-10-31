IMPORT FGL fgl_excel
IMPORT util

MAIN
DEFINE workbook     fgl_excel.workbookType 
DEFINE sheet        fgl_excel.sheetType  
DEFINE row          fgl_excel.rowType  
DEFINE cell         fgl_excel.cellType

    -- Read fgl_excel_test.xlsx
    LET workbook = fgl_excel.workbook_open("fgl_excel_test.xlsx")

    -- Number of sheets
    DISPLAY SFMT("Number of sheets = %1",workbook.getNumberOfSheets())

    -- Number of rows in sheet
    LET sheet = workbook.getSheetAt(0)
    DISPLAY SFMT("Number of rows in sheet = %1", sheet.getPhysicalNumberOfRows())

    -- Get row 12, java index 0 based so subtract 1
    LET row = sheet.getRow(11)
    DISPLAY SFMT("Last cell in row 12 = %1", row.getPhysicalNumberOfCells())

    -- Get cell A12,"Total" java index 0 based so subtract 1
    -- Java strongly typed so have to check Type and then use appropriate method to return
    LET cell = row.getCell(0)
    DISPLAY SFMT("Celltype in cell A12 = %1", cell.getCellType())
    DISPLAY SFMT("Value in cell A12 = %1", cell.getStringCellValue())

    -- Get cell B12, "Formula with total", java index 0 based so subtract 1 
    -- Java strongly typed so have to check Type and then use appropriate method to return
    -- As formula, should be able to get cached result but this not working
    -- Have to evaluate formula
    LET cell = row.getCell(1)
    DISPLAY SFMT("Celltype in cell B12 = %1", cell.getCellType())
    DISPLAY SFMT("Formula in cell B12 = %1", cell.getCellFormula())

    DISPLAY ""
    DISPLAY "Not sure why these return 0"
    DISPLAY SFMT("Get cached formula result type in cell B12 = %1", cell.getCachedFormulaResultType())
    DISPLAY SFMT("Value in cell B12 = %1", cell.getNumericCellValue())

    DISPLAY ""
    DISPLAY "This works but I should be able to read last value rather than evaluating"
    DISPLAY SFMT("Evaluated value in cell B12 = %1",  workbook.getCreationHelper().createFormulaEvaluator().evaluate(cell).getNumberValue())
    

END MAIN



