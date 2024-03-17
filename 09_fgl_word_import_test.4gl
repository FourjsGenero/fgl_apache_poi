IMPORT FGL fgl_word

MAIN

    DEFINE document fgl_word.documentType
    DEFINE paragraph fgl_word.paragraphType
    DEFINE run fgl_word.runType
    DEFINE table fgl_word.tableType
    DEFINE table_row fgl_word.tableRowType
    DEFINE table_cell fgl_word.tableCellType

    DEFINE paragraph_count, paragraph_idx INTEGER
    DEFINE run_count, run_idx INTEGER
    DEFINE table_count, table_idx INTEGER
    DEFINE table_row_count, table_row_idx INTEGER
    DEFINE table_cell_count, table_cell_idx INTEGER

    LET document = fgl_word.document_open("fgl_word_test.docx")
    
    LET paragraph_count = document.paragraph_count()
    DISPLAY SFMT("Paragraph Count = %1", paragraph_count USING "<&")
    FOR paragraph_idx = 1 TO paragraph_count
        LET paragraph = document.paragraph_get(paragraph_idx)
        DISPLAY SFMT("    Paragraph %1 = %2 Style = %3", paragraph_idx, paragraph.text_get(), paragraph.style_get())

        LET run_count = paragraph.run_count()
        DISPLAY SFMT("    Run Count = %1", run_count USING "<&")
        FOR run_idx = 1 TO run_count
            LET run = paragraph.run_get(run_idx)
            DISPLAY SFMT("        Run %1 = %2 Style = %3", run_idx, run.text_get(), run.style_get())
        END FOR
    END FOR

    LET table_count = document.table_count()
    DISPLAY SFMT("Table Count = %1", table_count USING "<&")
    FOR table_idx = 1 TO table_count
        DISPLAY SFMT("Table = %1", table_idx)
        LET table = document.table_get(table_idx)

        LET table_row_count = table.table_row_count()
        DISPLAY SFMT("    Table Row Count = %1", table_row_count)
        FOR table_row_idx = 1 TO table_row_count
            DISPLAY SFMT("    Table Row = %1", table_row_idx)
            LET table_row = table.table_row_get(table_row_idx)

            LET table_cell_count = table_row.table_row_cell_count()
            DISPLAY SFMT("        Table Cell Count = %1", table_cell_count)
            FOR table_cell_idx = 1 TO table_cell_count
                DISPLAY SFMT("        Table Cell = %1", table_cell_idx)
                LET table_cell = table_row.table_row_cell_get(table_cell_idx)
                DISPLAY SFMT("            Cell %1 = %2", table_cell_idx, table_cell.text_get())

                LET paragraph_count = table_cell.paragraph_count()
                FOR paragraph_idx = 1 TO paragraph_count
                    LET paragraph = table_cell.paragraph_get(paragraph_idx)
                    DISPLAY SFMT("                Paragraph %1 = %2 Style = %3",
                        paragraph_idx, paragraph.text_get(), paragraph.style_get())

                    LET run_count = paragraph.run_count()
                    DISPLAY SFMT("                    Run Count = %1", run_count USING "<&")
                    FOR run_idx = 1 TO run_count
                        LET run = paragraph.run_get(run_idx)
                        DISPLAY SFMT("                     Run %1 = %2 Style = %3", run_idx, run.text_get(), run.style_get())
                    END FOR
                END FOR
            END FOR
        END FOR
    END FOR
END MAIN
