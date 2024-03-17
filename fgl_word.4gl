IMPORT JAVA java.io.FileOutputStream
IMPORT JAVA java.io.FileInputStream

IMPORT JAVA java.util.List

# for export to PDF, need to find appropriate jar.  Despite note not in ApachePOI
#import java org.apache.poi.xwpf.converter.pdf.PdfConverter
#import java org.apache.poi.xwpf.converter.pdf.PdfOptions

IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFDocument
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFParagraph
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFRun
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFStyles
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFTable
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFTableRow
IMPORT JAVA org.apache.poi.xwpf.usermodel.XWPFTableCell

PUBLIC TYPE documentType RECORD
    j_document XWPFDocument
END RECORD

PUBLIC TYPE tableType RECORD
    j_table XWPFTable
END RECORD

PUBLIC TYPE tableRowType RECORD
    j_table_row XWPFTableRow
END RECORD

PUBLIC TYPE tableCellType RECORD
    j_table_cell XWPFTableCell
END RECORD

PUBLIC TYPE paragraphType RECORD
    j_paragraph XWPFParagraph
END RECORD

PUBLIC TYPE runType RECORD
    j_run XWPFRun
END RECORD


# USeful links
# https://poi.apache.org/components/document/quick-guide-xwpf.html
# https://stackoverflow.com/questions/38105235/write-data-to-word-document-using-apache-poi (this is what got me first started
# https://poi.apache.org/apidocs/dev/org/apache/poi/xwpf/usermodel/XWPFDocument.html From here follow links to other classes

# Useful Tip
# if you right-click a docx/dotx etc and extract to zip, you will see that a docx is simply a compressed set of xml files
# you will find xml with names like styles.xml holds the styles information

# Code Note
# tried to be noun_verb e.g. document_create(), test_get()
# as the functions are 4gl centric, any get uses 1 as base, not 0 as in Java
# note the CAST technique when using get method to return an element from the list
# a pargraphis what it says it is, a run is an element of a paragraph that has the same format.  If fomrat changes then that is a new run
# e.g. <b>bold</b> and <i>italic</i> would be three runs
#
# parapgraph -> run
# table -> tablerow -> tablecell -> parapgraph -> run


# start a new document from scratch
FUNCTION document_create() RETURNS documentType
DEFINE d documentType

    LET d.j_document = XWPFDocument.create()
    RETURN d.*
END FUNCTION

# take a template (dotx as starting point)
FUNCTION document_create_from_template(template_filename STRING) RETURNS documentType
DEFINE d documentType
DEFINE fi FileInputStream


    LET fi = FileInputStream.create(template_filename)
    LET d.j_document = XWPFDocument.create(fi)
    CALL fi.close()
    
    -- https://stackoverflow.com/a/54377500/2088674
    CALL d.j_document.getPackage().replaceContentType( "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml","application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")

    RETURN d.*
END FUNCTION

# open existing word document
FUNCTION document_open(filename STRING) RETURNS documentType
DEFINE d documentType
DEFINE fi FileInputStream

    LET fi = FileInputStream.create(filename)
    LET d.j_document = XWPFDocument.create(fi)
    CALL fi.close()
    RETURN d.*
END FUNCTION

# documents
# save word document to file
FUNCTION (this documentType) write(filename STRING) RETURNS ()
DEFINE fo FileOutputStream

    LET fo = FileOutputStream.create(filename)
    CALL this.j_document.write(fo)
    CALL fo.close()
END FUNCTION

FUNCTION (this documentType) paragraph_count() RETURNS INTEGER
    RETURN this.j_document.getParagraphs().size()
END FUNCTION

FUNCTION (this documentType) table_count() RETURNS INTEGER
    RETURN this.j_document.getTables().size()
END FUNCTION

FUNCTION (this documentType) paragraph_get(idx INTEGER) RETURNS paragraphType
DEFINE p paragraphType
    LET p.j_paragraph = CAST(this.j_document.getParagraphs().get(idx-1) AS XWPFParagraph)
    RETURN p.*
END FUNCTION

FUNCTION (this documentType) paragraph_create() RETURNS paragraphType
DEFINE p paragraphType

    LET p.j_paragraph = this.j_document.createParagraph()
    RETURN p.*
END FUNCTION

FUNCTION (this documentType) table_get(idx INTEGER) RETURNS tableType
DEFINE p tableType
    LET p.j_table = CAST(this.j_document.getTables().get(idx-1) AS XWPFTable)
    RETURN p.*
END FUNCTION

FUNCTION (this documentType) table_create() RETURNS tableType
DEFINE p tableType

    LET p.j_table = this.j_document.createTable()
    RETURN p.*
END FUNCTION



# tables
FUNCTION (this tableType) table_row_count() RETURNS INTEGER
    RETURN this.j_table.getNumberOfRows()
END FUNCTION

FUNCTION (this tableType) table_row_get(idx INTEGER) RETURNS tableRowType
DEFINE p tableRowType
    LET p.j_table_row = CAST(this.j_table.getRow(idx-1) AS XWPFTableRow)
    RETURN p.*
END FUNCTION

FUNCTION (this tableType) table_row_create() RETURNS tableRowType
DEFINE p tableRowType

    LET p.j_table_row = this.j_table.createRow()
    RETURN p.*
END FUNCTION



# table rows
FUNCTION (this tableRowType) table_row_cell_count() RETURNS INTEGER
    RETURN this.j_table_row.getTableCells().size()
END FUNCTION

FUNCTION (this tableRowType) table_row_cell_get(idx INTEGER) RETURNS tableCellType
DEFINE p tableCellType
    LET p.j_table_cell = CAST(this.j_table_row.getCell(idx-1) AS XWPFTableCell)
    RETURN p.*
END FUNCTION

FUNCTION (this tableRowType) table_row_cell_create() RETURNS tableCellType
DEFINE p tableCellType

    LET p.j_table_cell = this.j_table_row.createCell()
    RETURN p.*
END FUNCTION



# table cells
FUNCTION (this tableCellType) text_get() RETURNS STRING
    RETURN this.j_table_cell.getText()
END FUNCTION

FUNCTION (this tableCellType) paragraph_count() RETURNS INTEGER
    RETURN this.j_table_cell.getParagraphs().size()
END FUNCTION

FUNCTION (this tableCellType) paragraph_get(idx INTEGER) RETURNS paragraphType
DEFINE p paragraphType
    LET p.j_paragraph = CAST(this.j_table_cell.getParagraphs().get(idx-1) AS XWPFParagraph)
    RETURN p.*
END FUNCTION

FUNCTION (this tableCellType) paragraph_create() RETURNS paragraphType
DEFINE p paragraphType

    LET p.j_paragraph = this.j_table_cell.addParagraph()
    RETURN p.*
END FUNCTION




# paragraphs
FUNCTION (this paragraphType) run_create() RETURNS runType
DEFINE r runType

    LET r.j_run = this.j_paragraph.createRun()
    RETURN r.*
END FUNCTION

FUNCTION (this paragraphType) run_count() RETURNS INTEGER
    RETURN this.j_paragraph.getRuns().size()
END FUNCTION

FUNCTION (this paragraphType) run_get(idx INTEGER) RETURNS runType
DEFINE r runType
    LET r.j_run = CAST(this.j_paragraph.getRuns().get(idx-1) AS XWPFRun)
    RETURN r.*
END FUNCTION

FUNCTION (this paragraphType) text_get() RETURNS STRING
    RETURN this.j_paragraph.getText()
END FUNCTION

FUNCTION (this paragraphType) style_set(s STRING) RETURNS ()
    CALL this.j_paragraph.setStyle(s)
END FUNCTION

FUNCTION (this paragraphType) style_get() RETURNS STRING
    RETURN this.j_paragraph.getStyle()
END FUNCTION



# runs
FUNCTION (this runType) text_set(t STRING) RETURNS ()
    CALL this.j_run.setText(t)
END FUNCTION

FUNCTION (this runType) text_get() RETURNS STRING
    RETURN this.j_run.getText(0)
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

FUNCTION (this runType) style_set(s STRING) RETURNS ()
    CALL this.j_run.setStyle(s)
END FUNCTION

FUNCTION (this runType) style_get() RETURNS STRING
    RETURN this.j_run.getStyle()
END FUNCTION





-- Some code that i had at one stage but I will keep for a whlte to keep techniques on hand
-- mauy need to use at some point
--    -- Copy Styles from template file to new document
--    LET j_styles = d.j_document.createStyles()
--    CALL j_styles.setStyles(j_template_document.getStyle())
-- 
--    -- Copy Paragraph from template file to new document
--    LET paragraph_count = j_template_document.getParagraphs().size()
--    FOR paragraph_idx = 1 TO paragraph_count
--        LET template_paragraph = CAST(j_template_document.getParagraphs().get(paragraph_idx-1) AS  XWPFParagraph)
--        LET new_paragraph =  d.j_document.createParagraph()
--        CALL new_paragraph.setStyle(template_paragraph.getStyle())
--
--        LET run_count = template_paragraph.getRuns().size()
--        FOR run_idx = 1 TO run_count
--            LET template_run = CAST(template_paragraph.getRuns().get(run_idx) AS XWPFRun)
--            LET new_run = new_paragraph.createRun()
--            CALL new_run.setText(template_run.getText(0))
--        END FOR
--    END FOR

     
    
