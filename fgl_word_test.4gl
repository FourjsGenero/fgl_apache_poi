IMPORT FGL fgl_word
IMPORT util

MAIN
DEFINE document     fgl_word.documentType
DEFINE paragraph    fgl_word.paragraphType
DEFINE run          fgl_word.runType

    -- create document
    CALL fgl_word.document_create() RETURNING document

    -- create paragraph
    CALL fgl_word.document_paragraph_create(document) RETURNING paragraph

    -- Create Run
    CALL fgl_word.paragraph_run_create(paragraph) RETURNING run
    CALL fgl_word.run_text_set(run, "The quick brown fox jumps over the lazy dog.")
   
    CALL fgl_word.document_paragraph_create(document) RETURNING paragraph
    CALL fgl_word.paragraph_run_create(paragraph) RETURNING run
    CALL fgl_word.run_text_set(run, "Some text in ")

    CALL fgl_word.paragraph_run_create(paragraph) RETURNING run
    CALL fgl_word.run_text_set(run, "bold")
    CALL fgl_word.run_bold_set(run, TRUE)

    CALL fgl_word.paragraph_run_create(paragraph) RETURNING run
    CALL fgl_word.run_text_set(run, " and ")

    CALL fgl_word.paragraph_run_create(paragraph) RETURNING run
    CALL fgl_word.run_text_set(run, "italics")
    CALL fgl_word.run_italic_set(run, TRUE)

    CALL fgl_word.paragraph_run_create(paragraph) RETURNING run
    CALL fgl_word.run_text_set(run, ".")

    CALL fgl_word.document_paragraph_create(document) RETURNING paragraph
    CALL fgl_word.paragraph_run_create(paragraph) RETURNING run
    CALL fgl_word.run_text_set(run, SFMT("This document created on %1 at %2", TODAY, CURRENT HOUR TO SECOND))
    
    -- Write to File
    CALL fgl_word.document_write(document, "fgl_word_test.docx")
END MAIN