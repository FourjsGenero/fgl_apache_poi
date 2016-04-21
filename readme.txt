This example program wraps the Java Apache POI libraries in 4gl libraries
so that they can be called by a 4gl programmer.

You will need to download the Java Apache POI libraries from https://poi.apache.org/

In the fgl_apache_poi node, set the POI_HOME variable to where you have 
downloaded the Apache POI Libraries

As the versions increase you may need to alter the CLASSPATH variable set in
the same node.  Currently I use the value ...
CLASSPATH=$(POI_HOME)/poi-3.10.1-20140818.jar;$(POI_HOME)/poi-ooxml-3.10.1-20140818.jar;$(POI_HOME)/poi-ooxml-schemas-3.10.1-20140818.jar;$(POI_HOME)/ooxml-lib/dom4j-1.6.1.jar;$(POI_HOME)/ooxml-lib/xmlbeans-2.6.0.jar;$(CLASSPATH)
... and as you can hopefully see there is some versioning in the filenames.