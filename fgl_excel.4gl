IMPORT JAVA java.io.FileOutputStream
IMPORT JAVA java.io.FileInputStream

IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFWorkbook
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFSheet
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFPivotTable
IMPORT JAVA org.apache.poi.ss.usermodel.Row
IMPORT JAVA org.apache.poi.ss.usermodel.Cell
IMPORT JAVA org.apache.poi.ss.usermodel.CellStyle
IMPORT JAVA org.apache.poi.ss.usermodel.HorizontalAlignment
IMPORT JAVA org.apache.poi.ss.usermodel.Font
IMPORT JAVA org.apache.poi.ss.usermodel.PrintSetup

IMPORT JAVA org.apache.poi.ss.SpreadsheetVersion
IMPORT JAVA org.apache.poi.ss.usermodel.DataConsolidateFunction

--This java entry is needed for the STFieldSortType to sort PivotTable columns/rows.  It is outside of the ApachePOI download
-- IMPORT JAVA org.openxmlformats.schemas.spreadsheetml.x2006.main.STFieldSortType

IMPORT JAVA org.apache.poi.ss.util.CellUtil
IMPORT JAVA org.apache.poi.ss.util.CellReference
IMPORT JAVA org.apache.poi.ss.util.AreaReference
IMPORT JAVA org.apache.poi.ss.usermodel.IndexedColors
IMPORT JAVA org.apache.poi.ss.usermodel.BorderStyle
IMPORT JAVA java.util.HashMap
IMPORT JAVA java.lang.Integer

PUBLIC TYPE workbookType XSSFWorkbook
PUBLIC TYPE sheetType XSSFSheet
PUBLIC TYPE pivotTableType XSSFPivotTable
PUBLIC TYPE rowType Row
PUBLIC TYPE cellType Cell
PUBLIC TYPE cellStyleType CellStyle
PUBLIC TYPE fontType Font

FUNCTION workbook_create() RETURNS workbookType
    RETURN XSSFWorkbook.create()
END FUNCTION

FUNCTION workbook_writeToFile(w, filename)
    DEFINE w workbookType
    DEFINE filename STRING
    DEFINE fo FileOutputStream

    LET fo = FileOutputStream.create(filename)
    CALL w.write(fo)
    CALL fo.close()
END FUNCTION

FUNCTION workbook_open(filename) RETURNS workbookType
    DEFINE filename STRING
    DEFINE fi FileInputStream
    DEFINE w workbookType

    LET fi = FileInputStream.create(filename)
    LET w = XSSFWorkbook.create(fi)
    RETURN w
END FUNCTION

FUNCTION workbook_createsheet(w) RETURNS sheetType
    DEFINE w workbookType
    DEFINE s sheetType

    LET s = w.createSheet()
    RETURN s
END FUNCTION

FUNCTION sheet_createrow(s, idx) RETURNS rowType
    DEFINE s sheetType
    DEFINE idx INTEGER
    DEFINE r rowType
    LET r = s.createRow(idx)
    RETURN r
END FUNCTION

FUNCTION sheet_autosizecolumn(s, c)
    DEFINE s sheetType
    DEFINE c INTEGER

    CALL s.autoSizeColumn(c)
END FUNCTION

FUNCTION sheet_columnwidth_set(s, c, wf)
    DEFINE s sheetType
    DEFINE c INTEGER
    DEFINE wf FLOAT
    DEFINE wi INTEGER

    LET wi = 256.0 * wf
    CALL s.setColumnWidth(c, wi)
END FUNCTION

FUNCTION row_createcell(r, idx) RETURNS cellType
    DEFINE r rowType
    DEFINE idx INTEGER
    DEFINE c cellType
    LET c = r.createCell(idx)
    RETURN c
END FUNCTION

FUNCTION cell_value_set(c, v)
    DEFINE c cellType
    DEFINE v STRING
    CALL c.setCellValue(v)
END FUNCTION

FUNCTION cell_number_set(c, v)
    DEFINE c cellType
    DEFINE v FLOAT

    CALL c.setCellValue(v)
END FUNCTION

FUNCTION cell_formula_set(c, v)
    DEFINE c cellType
    DEFINE v STRING
    CALL c.setCellFormula(v)
END FUNCTION

-- Map A to 0 B to 1, Z to 25, AA to 26, AZ to 51
FUNCTION column2row(col)
    DEFINE col STRING

    CASE
        WHEN col MATCHES "[A-Z]"
            RETURN ORD(col) - 65
        WHEN col MATCHES "[A-Z][A-Z]"
            RETURN ((ORD(col.subString(1, 1)) - 65) * 26) + (ORD(col.subString(2, 2)) - 65)
    END CASE
    RETURN -1
END FUNCTION

FUNCTION cell_style_set(c, s)
    DEFINE c cellType
    DEFINE s cellStyleType
    CALL c.setCellStyle(s)
END FUNCTION

FUNCTION font_create(w) RETURNS fontType
    DEFINE w workBookType
    DEFINE f fontType
    LET f = w.createFont()
    RETURN f
END FUNCTION

FUNCTION font_set(f, a, v)
    DEFINE f fontType
    DEFINE a STRING
    DEFINE v STRING

    CASE
        WHEN a = "weight" AND v = "bold"
            CALL f.setBold(TRUE)
        WHEN a = "weight" AND v = "normal"
            CALL f.setBold(FALSE)
            -- add more as required
    END CASE
END FUNCTION

FUNCTION style_create(w) RETURNS cellStyleType
    DEFINE w workBookType
    DEFINE s cellStyleType
    LET s = w.createCellStyle()
    RETURN s
END FUNCTION

FUNCTION style_set(s, a, v)
    DEFINE s cellStyleType
    DEFINE a STRING
    DEFINE v STRING

    CASE
        WHEN a = "alignment" AND v = "center"
            CALL s.setAlignment(HorizontalAlignment.CENTER)

        WHEN a = "alignment" AND v = "left"
            CALL s.setAlignment(HorizontalAlignment.LEFT)

        WHEN a = "alignment" AND v = "right"
            CALL s.setAlignment(HorizontalAlignment.RIGHT)

        WHEN a = "alignment" AND v = "justify"
            CALL s.setAlignment(HorizontalAlignment.JUSTIFY)

        WHEN a = "alignment" AND v = "general"
            CALL s.setAlignment(HorizontalAlignment.GENERAL)

            -- add more as required
    END CASE
END FUNCTION

FUNCTION style_font_set(s, f)
    DEFINE s cellStyleType
    DEFINE f fontType
    CALL s.setFont(f)
END FUNCTION

FUNCTION workbook_fit_to_page(old_filename STRING, new_filename STRING)
    DEFINE w workbookType
    DEFINE s sheetType

    DEFINE ps PrintSetup
    CONSTANT SHORT_ONE SMALLINT = 1

    LET w = workbook_open(old_filename)
    LET s = w.getSheetAt(0)
    CALL s.setFitToPage(TRUE)
    CALL s.setAutobreaks(TRUE)
    LET ps = s.getPrintSetup()
    CALL ps.setFitWidth(SHORT_ONE)
    CALL ps.setFitHeight(SHORT_ONE)

    CALL workbook_writeToFile(w, new_filename)
END FUNCTION

FUNCTION cellutil_set_cell_property_experiment(c cellType)

    DEFINE h java.util.HashMap

    LET h = java.util.HashMap.create()
    #CALL h.put( CellUtil.FILL_BACKGROUND_COLOR, java.lang.Integer.create(IndexedColors.RED.index))
    #CALL h.put( CellUtil.FILL_FOREGROUND_COLOR, java.lang.Integer.create(IndexedColors.BLUE.index))

    CALL h.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM)
    CALL h.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM)
    CALL h.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM)
    CALL h.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM)
    CALL h.put(CellUtil.TOP_BORDER_COLOR, java.lang.Integer.create(IndexedColors.RED.getIndex()))
    CALL h.put(CellUtil.BOTTOM_BORDER_COLOR, java.lang.Integer.create(IndexedColors.BLUE.getIndex()))
    CALL h.put(CellUtil.LEFT_BORDER_COLOR, java.lang.Integer.create(IndexedColors.YELLOW.getIndex()))
    CALL h.put(CellUtil.RIGHT_BORDER_COLOR, java.lang.Integer.create(IndexedColors.GREEN.getIndex()))

    CALL CellUtil.setCellStyleProperties(c, h)
END FUNCTION

FUNCTION pivot_table_create(w workbookType, data_sheet sheetType, area_ref STRING, pivot_sheet sheetType, cell_ref STRING)
    RETURNS pivotTableType
    DEFINE pivot_table pivotTableType

    LET pivot_table =
        pivot_sheet.createPivotTable(
            AreaReference.create(area_ref, SpreadsheetVersion.EXCEL2007), CellReference.create(cell_ref), data_sheet)
    RETURN pivot_table
END FUNCTION

FUNCTION pivot_table_add_column(pt pivotTableType, idx INTEGER)
    CALL pt.addColLabel(idx)

    --This line uses STFieldSortType to sort PivotTable columns/rows.  It is outside of the ApachePOI download
    -- CALL pt.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(idx).setSortType(STFieldSortType.ASCENDING)
END FUNCTION

FUNCTION pivot_table_add_row(pt pivotTableType, idx INTEGER)
    CALL pt.addRowLabel(idx)

    --This line uses STFieldSortType to sort PivotTable columns/rows.  It is outside of the ApachePOI download
    -- CALL pt.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(idx).setSortType(STFieldSortType.ASCENDING)
END FUNCTION

-- Add variations of this function to do different operations
FUNCTION pivot_table_add_data_sum(pt pivotTableType, idx INTEGER)
    CALL pt.addColumnLabel(DataConsolidateFunction.SUM, idx)
END FUNCTION
