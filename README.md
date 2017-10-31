# FGL wrapper for the Apache POI Java Library

## Description

This example program wraps the Java Apache POI libraries in 4gl libraries
so that they can be called by a 4gl programmer.

## Usage

You will need to download the Java Apache POI libraries from:

https://poi.apache.org/

In the fgl_apache_poi node, set the POI_HOME variable to where you have 
downloaded the Apache POI Libraries

As the versions increase you may need to alter the CLASSPATH variable set
in the same node.

Currently I use the value ...
``
CLASSPATH=$(POI_HOME)/poi-3.10.1-20140818.jar;$(POI_HOME)/poi-ooxml-3.10.1-20140818.jar;$(POI_HOME)/poi-ooxml-schemas-3.10.1-20140818.jar;$(POI_HOME)/ooxml-lib/dom4j-1.6.1.jar;$(POI_HOME)/ooxml-lib/xmlbeans-2.6.0.jar;$(CLASSPATH)
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