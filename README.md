# FGL wrapper for the Apache POI Java Library

## Description

This example program wraps the Java Apache POI libraries in 4gl libraries
so that they can be called by a 4gl programmer.

The Apache POI libraries are used to interact with Word and Excel documents.  They also include some of the functions that are available inside Excel including financial functions.

## Usage

You will need to download the Java Apache POI libraries from:

https://poi.apache.org/

In the fgl_apache_poi node, set the POI_HOME variable to where you have
downloaded the Apache POI Libraries

As the versions increase you may need to alter the CLASSPATH variable set
in the same node.

Currently I use the value ...
``
$(POI_HOME)/poi-4.1.0.jar;$(POI_HOME)/poi-ooxml-4.1.0.jar;$(POI_HOME)/poi-ooxml-schemas-4.1.0.jar;$(POI_HOME)/ooxml-lib/curvesapi-1.06.jar;$(POI_HOME)/ooxml-lib/xmlbeans-3.1.0.jar;$(POI_HOME)/lib/commons-collections4-4.3.jar;$(POI_HOME)/lib/commons-compress-1.18.jar;$(POI_HOME)/lib/commons-math3-3.6.1.jar;$(CLASSPATH)
``
... and as you can see, there is some versioning in the filenames.


## Test Programs

### fgl_excel_test

Create an Excel Document (fgl_excel_test).  Note the last line has date/time indicating the file has just been created.  Also note the total line is a formula, not a value.

![Example Excel](https://user-images.githubusercontent.com/13615993/32205574-dded7afe-be54-11e7-9809-065ecc4f5b35.png)

### fgl_word_test

Create a Word Document (fgl_word_test).  Note the last line has date/time indicating the file has just been created

![Example Word](https://user-images.githubusercontent.com/13615993/32205573-ddb64584-be54-11e7-85be-20bc00c0da2a.png)


### fgl_excel_calculation

Illustrates how an Excel document can be created in memory and you can use the Excel built-in formulas such as IRR rather than creating 4gl equivalents for these functions

### fgl_excel_import

Illustrates how you can read an existing Excel spreadsheet


### fgl_excel_generic_test

Illustrates the use of a function that can take as input a SQL clause, and generate a spreadsheet containing an unload of that data.  Combines base.SqlHandle with the ApachePOI classes

### fgl_financial_test

Illustrates the use of the financial methods (NPV, IRR) etc that are available.  Note also inside fgl_financial the use of JAVA arrays to pass a list of payments

### fgl_excel_fit_to_page

Illustrates how you can fit to page an existing spreadsheet.  It takes two arguments, an existing filename and a new filename, and will save the existing file using the new filename having modified it to fit to 1 page.  With the example, fgl_excel_fit_to_page.xlsx is the original, fgl_excel_fit_to_page2.xlsx has been modified to print to 1x1 page, to verify go File->Print and note the appearance in the preview.

### fgl_excel_pivot

Illustrates how you can create a Pivot table.  It creates two sheets, one sheet with columnar data in 5 columns, and the second sheet that has a pivot table added with the 3rd column assigned to rows of the pivot table, the 4th column (Months) assigned to columns of the pivot table, and a grouped sum across these two fields of the data in the 5th column (Qty).

There is a quirk with this example, this line:

``
CALL pt.getCTPivotTableDefinition().getPivotFields().getPivotFieldArray(idx).setSortType(STFieldSortType.ASCENDING)
``

is commented out of the pivot_table_add_row and pivot_table_add_column functions as well as the line:

``
-- IMPORT JAVA org.openxmlformats.schemas.spreadsheetml.x2006.main.STFieldSortType
``

This is because the jar lib ooxml-schemas-1.4.jar is required and that is not part of the Apache POI download.  See comments such as https://stackoverflow.com/questions/38470735/how-do-you-sort-pivot-table-data-using-the-apache-poi-library-in-java

If you want to sort the columns and rows, uncomment these 3 lines and find and download the ooxml-schemas jar  (If there is a better way let me know)

![Pivot Example](https://user-images.githubusercontent.com/13615993/177439323-ae769063-d01a-4de1-a634-48961a7f0737.png)

### fgl_word_import

Import and read a Word document (paragraphs, runs, table, table rows, table cells at this stage)

### fgl_word_template

Open a Word template file (dotx) and start using that to create a Word document.

## Other Notes

Have a look at http://poi.apache.org/components/spreadsheet/quick-guide.html#DataFormats for examples of methods will need to add to do certain things in Excel
