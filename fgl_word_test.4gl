IMPORT FGL fgl_word
IMPORT util

MAIN
DEFINE document     fgl_word.documentType
DEFINE paragraph    fgl_word.paragraphType
DEFINE run          fgl_word.runType

    -- create document
    CALL fgl_word.document_create() RETURNING document.*

    -- create paragraph
    #CALL fgl_word.document_paragraph_create(document) RETURNING paragraph
    CALL document.paragraph_create() RETURNING paragraph.*
    
    -- Create Run
    CALL paragraph.run_create() RETURNING run.*
    CALL run.text_set("The quick brown fox jumps over the lazy dog.")
   
    CALL document.paragraph_create() RETURNING paragraph.*
    CALL paragraph.run_create() RETURNING run.*
    CALL run.text_set("Some text in ")

    CALL paragraph.run_create() RETURNING run.*
    CALL run.text_set("bold")
    CALL run.bold_set(TRUE)

    CALL paragraph.run_create() RETURNING run.*
    CALL run.text_set(" and ")

    CALL paragraph.run_create() RETURNING run.*
    CALL run.text_set("italics")
    CALL run.italic_set(TRUE)

    CALL paragraph.run_create() RETURNING run.*
    CALL run.text_set(".")

    CALL document.paragraph_create() RETURNING paragraph.*
    CALL paragraph.run_create() RETURNING run.*
    CALL run.text_set(SFMT("This document created on %1 at %2 by fgl_word_test.4gl", TODAY, CURRENT HOUR TO SECOND))
    
    CALL document.write("fgl_word_test.docx")
END MAIN