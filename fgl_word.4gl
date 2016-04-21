IMPORT JAVA java.io.FileOutputStream


IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFDocument
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFParagraph
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFRun
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFStyle

PUBLIC TYPE documentType XWPFDocument
PUBLIC TYPE paragraphType XWPFParagraph
PUBLIC TYPE runType XWPFRun
PUBLIC TYPE styleType XWPFStyle



FUNCTION document_create()
    RETURN XWPFDocument.create()
END FUNCTION



FUNCTION document_paragraph_create(d)
DEFINE d documentType
DEFINE p paragraphType

    LET p= d.createParagraph()
    RETURN p
END FUNCTION



FUNCTION paragraph_run_create(p)
DEFINE p paragraphType
DEFINE r runType

    LET r = p.createRun()
    RETURN r
END FUNCTION



FUNCTION run_text_set(r, t)
DEFINE r runType
DEFINE t STRING

    CALL r.setText(t)
END FUNCTION


FUNCTION run_break(r)
DEFINE r runType

    CALL r.addBreak()
END FUNCTION

FUNCTION run_bold_set(r, b)
DEFINE r runType
DEFINE b BOOLEAN
    IF b THEN
        CALL r.setBold(TRUE)
    ELSE
        CALL r.setBold(FALSE)
    END IF
END FUNCTION

FUNCTION run_italic_set(r, b)
DEFINE r runType
DEFINE b BOOLEAN
    IF b THEN
        CALL r.setItalic(TRUE)
    ELSE
        CALL r.setItalic(FALSE)
    END IF
END FUNCTION




FUNCTION document_write(d, filename)
DEFINE d documentType
DEFINE filename STRING
DEFINE fo FileOutputStream

    LET fo = FileOutputStream.create(filename)
    CALL d.write(fo)
    CALL fo.close()
END FUNCTION
