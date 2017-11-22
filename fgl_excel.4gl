IMPORT JAVA java.io.FileOutputStream
IMPORT JAVA java.io.FileInputStream

IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFWorkbook
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFSheet
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFRow
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFCell
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFCellStyle
IMPORT JAVA org.apache.poi.xssf.usermodel.XSSFFont


PUBLIC TYPE workbookType XSSFWorkbook
PUBLIC TYPE sheetType XSSFSheet
PUBLIC TYPE rowType XSSFRow
PUBLIC TYPE cellType XSSFCell
PUBLIC TYPE cellStyleType XSSFCellStyle
PUBLIC TYPE fontType XSSFFont

FUNCTION workbook_create()
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


FUNCTION workbook_open(filename)
DEFINE filename STRING
DEFINE fi FileInputStream
DEFINE w workbookType

    LET fi = FileInputStream.create(filename)
    LET w = XSSFWorkbook.create(fi)
    RETURN w
END FUNCTION



FUNCTION workbook_createsheet(w)
DEFINE w workbookType
DEFINE s sheetType

    LET s= w.createSheet()
    RETURN s
END FUNCTION



FUNCTION sheet_createrow(s,idx)
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

FUNCTION sheet_columnwidth_set(s, c, w)
DEFINE s sheetType
DEFINE c INTEGER
DEFINE w INTEGER

    CALL s.setColumnWidth(c,w)
END FUNCTION



FUNCTION row_createcell(r,idx)
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
    CALL c.setCellType(XSSFCell.CELL_TYPE_NUMERIC)
    CALL c.setCellValue(v)
END FUNCTION



FUNCTION cell_formula_set(c, v)
DEFINE c cellType
DEFINE v STRING
    CALL c.setCellType(XSSFCell.CELL_TYPE_FORMULA)
    CALL c.setCellFormula(v)
END FUNCTION



-- Map A to 0 B to 1, Z to 25, AA to 26, AZ to 51
FUNCTION column2row(col)
DEFINE col STRING

    CASE
        WHEN col MATCHES "[A-Z]" 
            RETURN ORD(col) - 65
        WHEN col MATCHES "[A-Z][A-Z]"
            RETURN ((ORD(col.subString(1,1))-65)*26) + (ORD(col.subString(2,2)) - 65)
    END CASE
    RETURN -1
END FUNCTION
    



FUNCTION cell_style_set(c, s)
DEFINE c cellType
DEFINE s cellStyleType
    CALL c.setCellStyle(s)
END FUNCTION



FUNCTION font_create(w)
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
        WHEN a="weight" AND v="bold"
            CALL f.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD)
        WHEN a="weight" AND v="normal"
            CALL f.setBoldweight(XSSFFont.BOLDWEIGHT_NORMAL)
        -- add more as required
    END CASE
END FUNCTION



FUNCTION style_create(w)
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
        WHEN a="alignment" AND v="center"
            CALL s.setAlignment(XSSFCellStyle.ALIGN_CENTER)
        WHEN a="alignment" AND v="left"
            CALL s.setAlignment(XSSFCellStyle.ALIGN_LEFT)
        WHEN a="alignment" AND v="right"
            CALL s.setAlignment(XSSFCellStyle.ALIGN_RIGHT)
        WHEN a="alignment" AND v="justify"
            CALL s.setAlignment(XSSFCellStyle.ALIGN_JUSTIFY)
        WHEN a="alignment" AND v="general"
            CALL s.setAlignment(XSSFCellStyle.ALIGN_GENERAL)
        -- add more as required
    END CASE
END FUNCTION



FUNCTION style_font_set(s,f)
DEFINE s cellStyleType
DEFINE f fontType
    CALL s.setFont(f)
END FUNCTION