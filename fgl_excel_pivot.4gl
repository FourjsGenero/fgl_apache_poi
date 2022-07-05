IMPORT FGL fgl_excel
IMPORT util

MAIN
    DEFINE workbook fgl_excel.workbookType
    DEFINE data_sheet, pivot_sheet fgl_excel.sheetType
    DEFINE pivot_table fgl_excel.pivotTableType
    DEFINE row fgl_excel.rowType
    DEFINE cell fgl_excel.cellType
    DEFINE header_style fgl_excel.cellStyleType
    DEFINE header_font fgl_excel.fontType

    DEFINE i INTEGER

    -- create workbook
    CALL fgl_excel.workbook_create() RETURNING workbook

    -- create a worksheet
    CALL fgl_excel.workbook_createsheet(workbook) RETURNING data_sheet
    CALL workbook.setSheetName(0,"Data")

    -- create a font, will be used in header
    CALL fgl_excel.font_create(workbook) RETURNING header_font
    CALL fgl_excel.font_set(header_font, "weight", "bold")

    -- create a style, will be used in header
    CALL fgl_excel.style_create(workbook) RETURNING header_style
    CALL fgl_excel.style_set(header_style, "alignment", "center")
    CALL fgl_excel.style_font_set(header_style, header_font)
    CALL fgl_excel.sheet_createrow(data_sheet, 0) RETURNING row

    -- add heading row
    CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
    CALL fgl_excel.cell_value_set(cell, "Product")
    CALL fgl_excel.cell_style_set(cell, header_style)

    CALL fgl_excel.row_createcell(row, column2row("B")) RETURNING cell
    CALL fgl_excel.cell_value_set(cell, "Customer")
    CALL fgl_excel.cell_style_set(cell, header_style)

    CALL fgl_excel.row_createcell(row, column2row("C")) RETURNING cell
    CALL fgl_excel.cell_value_set(cell, "Store")
    CALL fgl_excel.cell_style_set(cell, header_style)

    CALL fgl_excel.row_createcell(row, column2row("D")) RETURNING cell
    CALL fgl_excel.cell_value_set(cell, "Month")
    CALL fgl_excel.cell_style_set(cell, header_style)

    CALL fgl_excel.row_createcell(row, column2row("E")) RETURNING cell
    CALL fgl_excel.cell_value_set(cell, "Qty")
    CALL fgl_excel.cell_style_set(cell, header_style)

    -- add 10000 rows of random data
    FOR i = 1 TO 10000
        CALL fgl_excel.sheet_createrow(data_sheet, i) RETURNING row

        CALL fgl_excel.row_createcell(row, column2row("A")) RETURNING cell
        CALL fgl_excel.cell_value_set(cell, ASCII (util.Math.rand(26) + 65))

        CALL fgl_excel.row_createcell(row, column2row("B")) RETURNING cell
        CALL fgl_excel.cell_value_set(cell, ASCII (util.Math.rand(26) + 65))

        CALL fgl_excel.row_createcell(row, column2row("C")) RETURNING cell
        CALL fgl_excel.cell_value_set(cell, ASCII (util.Math.rand(10) + 65))

        CALL fgl_excel.row_createcell(row, column2row("D")) RETURNING cell
        CALL fgl_excel.cell_number_set(cell, util.Math.rand(12) + 1)

        CALL fgl_excel.row_createcell(row, column2row("E")) RETURNING cell
        CALL fgl_excel.cell_number_set(cell, util.Math.rand(100) + 1)
    END FOR

    -- Create Pivot on new second sheet

    LET pivot_sheet = fgl_excel.workbook_createsheet(workbook)
     CALL workbook.setSheetName(1,"Pivot")
    LET pivot_table = fgl_excel.pivot_table_create(workbook, data_sheet, "A1:E10001", pivot_sheet, "A3")
    CALL fgl_excel.pivot_table_add_row(pivot_table, 2)
    CALL fgl_excel.pivot_table_add_column(pivot_table, 3)
    CALL fgl_excel.pivot_table_add_data_sum(pivot_table, 4)

    -- Write to File
    CALL fgl_excel.workbook_writeToFile(workbook, "fgl_excel_pivot.xlsx");

END MAIN
