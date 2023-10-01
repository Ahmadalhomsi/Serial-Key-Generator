Intro
=====

Synthetica  is a 'Look and Feel' for Swing and is based  on Synth which is part 
of version 1.5  of the Java2 Platform, Standard Edition. It provides components
with rounded borders,  shadowed popup menus, nice icons and a new, fresh  look. 
Moreover it enables you to  modify existing  LAF's, or to create  your  own LAF
only by editing a XML based configuration file. You don't have to write complex
Java-GUI-Code, but you can integrate own code by your needs.

Home Page
=========

General:        http://www.jyloo.com
Synthetica:     http://www.jyloo.com/synthetica
FAQ:            http://www.jyloo.com/synthetica/faq
Download:       http://www.jyloo.com/synthetica/download
License:        http://www.jyloo.com/synthetica/license

Contact Addresses
=================

General:        info@jyloo.com
Sales:          sales@jyloo.com	
Support:        support@jyloo.com

System Requirements
===================

Java SE 5 (JRE 1.5.0) or above, Java 6 or above recommended
Synthetica V2.20.0 or above

Integration
===========

1. Ensure that your classpath contains all required Synthetica libraries, means
   the core library (synthetica.jar) and the used theme library. 
   Note:  In Synthetica V2.X.X the  Standard theme is already  part of the core 
   library - it's not a separate library file. Each other theme also depends on 
   on the core library!  

2. Enable Synthetica Look and Feel on application startup:

    import de.javasoft.plaf.synthetica.SyntheticaSimple2DLookAndFeel;

    try 
    {
      UIManager.setLookAndFeel(new SyntheticaSimple2DLookAndFeel());
    } 
    catch (Exception e) 
    {
      e.printStackTrace();
    }

3. Synthetica V2.X.X  supports  Java 9  in legacy mode  since V2.30.0 - without 
   support of the Java module system (JPMS). When running on Java 9 you have to
   pass the arguments below  to the JVM for proper execution without any errors
   and warnings.

   --add-exports=java.desktop/sun.swing=ALL-UNNAMED
   --add-exports=java.desktop/sun.swing.table=ALL-UNNAMED
   --add-exports=java.desktop/sun.swing.plaf.synth=ALL-UNNAMED
   --add-opens=java.desktop/javax.swing.plaf.synth=ALL-UNNAMED
   --add-opens=java.desktop/javax.swing.plaf.basic=ALL-UNNAMED
   --add-opens=java.desktop/javax.swing=ALL-UNNAMED
   --add-opens=java.desktop/javax.swing.tree=ALL-UNNAMED
   --add-opens=java.desktop/java.awt.event=ALL-UNNAMED
   --add-exports=java.desktop/com.sun.java.swing.plaf.windows=ALL-UNNAMED
   --add-exports=java.desktop/sun.awt.shell=ALL-UNNAMED
   --add-exports=java.desktop/com.sun.awt=ALL-UNNAMED
   --add-exports=java.base/sun.security.action=ALL-UNNAMED
   
   Alternatively,  you can  put the arguments into the manifest.mf file of your 
   application by adding the attributes below  - this is the recommended method 
   for deployment of your application.
   
   Add-Exports: java.desktop/sun.swing java.desktop/sun.swing.table java.desktop/sun.swing.plaf.synth java.desktop/com.sun.java.swing.plaf.windows java.desktop/sun.awt.shell java.desktop/com.sun.awt java.base/sun.security.action   
   Add-Opens: java.desktop/javax.swing.plaf.synth java.desktop/javax.swing.plaf.basic java.desktop/javax.swing java.desktop/javax.swing.tree java.desktop/java.awt.event

   
   