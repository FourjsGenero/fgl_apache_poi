IMPORT JAVA java.io.FileOutputStream

IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFDocument
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFParagraph
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFRun
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFStyle

PUBLIC TYPE documentType RECORD
    j_document XWPFDocument
END RECORD

PUBLIC TYPE paragraphType RECORD
    j_paragraph XWPFParagraph
END RECORD

PUBLIC TYPE runType RECORD
    j_run XWPFRun
END RECORD

PUBLIC TYPE styleType XWPFStyle



FUNCTION document_create() RETURNS documentType
DEFINE d documentType

    LET d.j_document = XWPFDocument.create()
    RETURN d.*
END FUNCTION



FUNCTION (this documentType) paragraph_create() RETURNS XWPFParagraph
DEFINE p XWPFParagraph

    LET p = this.j_document.createParagraph()
    RETURN p
END FUNCTION

FUNCTION (this paragraphType) run_create() RETURNS XWPFRun
DEFINE r XWPFRun

    LET r = this.j_paragraph.createRun()
    RETURN r
END FUNCTION



FUNCTION (this runType) text_set(t STRING) RETURNS ()
    CALL this.j_run.setText(t)
END FUNCTION




FUNCTION (this runType) break() RETURNS ()
    CALL this.j_run.addBreak()
END FUNCTION



FUNCTION (this runType) bold_set(b BOOLEAN) RETURNS ()
    CALL this.j_run.setBold(b)
END FUNCTION

FUNCTION (this runType) italic_set(i BOOLEAN) RETURNS ()
    CALL this.j_run.setItalic(i)
END FUNCTION



FUNCTION (this documentType) write(filename STRING) RETURNS ()
DEFINE fo FileOutputStream

    LET fo = FileOutputStream.create(filename)
    CALL this.j_document.write(fo)
    CALL fo.close()
END FUNCTION
