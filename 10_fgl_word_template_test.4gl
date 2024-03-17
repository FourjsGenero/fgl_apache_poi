
IMPORT FGL fgl_word

MAIN
DEFINE document     fgl_word.documentType
DEFINE paragraph    fgl_word.paragraphType
DEFINE run          fgl_word.runType

    -- create document
    CALL fgl_word.document_create_from_template("tf16402400_win32.dotx") RETURNING document.*
    #CALL fgl_word.document_create_from_template("quote.dotx") RETURNING document.*
    #CALL fgl_word.document_create_from_template("quote_byemail.dotx") RETURNING document.*

    -- Testing adding element using styke
    CALL document.paragraph_create() RETURNING paragraph.*
    CALL paragraph.style_set("Title")
    CALL paragraph.run_create() RETURNING run.*
    CALL run.text_set("Title")

    CALL document.paragraph_create() RETURNING paragraph.*
    CALL paragraph.run_create() RETURNING run.*
    CALL run.text_set("Normal")

    CALL document.write("fgl_word_template_test.docx")
END MAIN

