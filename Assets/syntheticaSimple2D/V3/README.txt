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

Java 9 or higher


Integration
===========

1. Ensure that your classpath contains all required Synthetica libraries, means
   the core library (synthetica.jar) and the used theme library. 
   Note: In Synthetica V3 the core library contains Synthetica core files only, 
   no themes. Since V3 the standard theme is _not_ any longer part of  the core 
   library! So you have to put all required theme libraries  on  your classpath 
   or modulepath too.  


2. Enable Synthetica  Look and  Feel on application startup.  Theme class names 
   follows the syntax below: 
     de.javasoft.synthetica.[theme].Synthetica[theme]LookAndFeel
    
     Example - Simple2D theme:
    
     import de.javasoft.synthetica.simple2d.SyntheticaSimple2DLookAndFeel;

     public static void main(String[] args) throws Exception
     {
       UIManager.setLookAndFeel(new SyntheticaSimple2DLookAndFeel();
       EventQueue.invokeLater(() -> new MyDesktopApplication());
       ... 


3. Synthetica  V3  also  supports  the Java module system  (JPMS alias Jigsaw). 
   However, for proper execution  without any  errors and  warnings you have to
   pass the arguments below to the JVM.

   --add-exports=java.desktop/sun.swing=synthetica.base
   --add-exports=java.desktop/sun.swing.table=synthetica.base
   --add-exports=java.desktop/sun.swing.plaf.synth=synthetica.base
   --add-opens=java.desktop/javax.swing.plaf.synth=synthetica.base
   --add-opens=java.desktop/javax.swing.plaf.basic=synthetica.base
   --add-opens=java.desktop/javax.swing=synthetica.base
   --add-opens=java.desktop/javax.swing.tree=synthetica.base
   --add-opens=java.desktop/java.awt.event=synthetica.base
   --add-exports=java.desktop/sun.awt.shell=synthetica.base
   --add-exports=java.base/sun.security.action=synthetica.base
   (no longer required since Synthetica V3.2: --add-exports=java.desktop/com.sun.awt=synthetica.base)
   
   In case you're using SyntheticaAddons you also have to add the arguments below
   --add-exports=java.desktop/com.sun.java.swing.plaf.windows=synthetica.addons.base
   --add-opens=java.desktop/javax.swing.plaf.synth=synthetica.addons.base
   --add-exports=java.desktop/com.sun.java.swing.plaf.windows=synthetica.addons.table
   --add-exports=java.desktop/sun.swing=synthetica.addons.swingx   
   
   Legacy mode (classpath mode)
   
   If you still want to run your application in classpath mode use the JVM arguments
   listed below. 
   
   --add-exports=java.desktop/sun.swing=ALL-UNNAMED
   --add-exports=java.desktop/sun.swing.table=ALL-UNNAMED
   --add-exports=java.desktop/sun.swing.plaf.synth=ALL-UNNAMED
   --add-opens=java.desktop/javax.swing.plaf.synth=ALL-UNNAMED
   --add-opens=java.desktop/javax.swing.plaf.basic=ALL-UNNAMED
   --add-opens=java.desktop/javax.swing=ALL-UNNAMED
   --add-opens=java.desktop/javax.swing.tree=ALL-UNNAMED
   --add-opens=java.desktop/java.awt.event=ALL-UNNAMED
   --add-exports=java.desktop/sun.awt.shell=ALL-UNNAMED   
   --add-exports=java.base/sun.security.action=ALL-UNNAMED
   (no longer required since Synthetica V3.2: --add-exports=java.desktop/com.sun.awt=ALL-UNNAMED)
   
   In case you're using SyntheticaAddons you also have to add the argument below
   --add-exports=java.desktop/com.sun.java.swing.plaf.windows=ALL-UNNAMED
   
   Alternatively,  you can  put the arguments into the manifest.mf file of your 
   application by adding the attributes below - however, this will work only if 
   your application is executed as jar file in legacy mode (classpath mode). 
   Means as usual with specifying a class path instead of a module path.
   Note: Be aware that Web Start does not respect these entries - in this case 
   you have to use the JVM arguments from above and add it to the j2se entry of 
   your jnlp file.   
   
   Add-Exports: java.desktop/sun.swing java.desktop/sun.swing.table java.desktop/sun.swing.plaf.synth java.desktop/sun.awt.shell java.base/sun.security.action   
   Add-Opens: java.desktop/javax.swing.plaf.synth java.desktop/javax.swing.plaf.basic java.desktop/javax.swing java.desktop/javax.swing.tree java.desktop/java.awt.event

   When using SyntheticaAddons:
   Add-Exports: java.desktop/sun.swing java.desktop/sun.swing.table java.desktop/sun.swing.plaf.synth java.desktop/sun.awt.shell java.base/sun.security.action java.desktop/com.sun.java.swing.plaf.windows   
   Add-Opens: java.desktop/javax.swing.plaf.synth java.desktop/javax.swing.plaf.basic java.desktop/javax.swing java.desktop/javax.swing.tree java.desktop/java.awt.event


4. JPMS module names

   All modules of Synthetica V3 are listed below. 
   
   Synthetica core module          ->  synthetica.base
   Synthetica Dark theme           ->  synthetica.theme.dark
   Synthetica Plain theme          ->  synthetica.theme.plain
   Synthetica BlueLight theme      ->  synthetica.theme.bluelight
   Synthetica AluOxide theme       ->  synthetica.theme.aluoxide
   Synthetica Classy theme         ->  synthetica.theme.classy
   Synthetica BlackEye theme       ->  synthetica.theme.blackeye
   Synthetica Simple2D theme       ->  synthetica.theme.simple2d
   Synthetica WhiteVision theme    ->  synthetica.theme.whitevision
   Synthetica SkyMetallic theme    ->  synthetica.theme.skymetallic
   Synthetica MauveMetallic theme  ->  synthetica.theme.mauvemetallic
   Synthetica OrangeMetallic theme ->  synthetica.theme.orangemetallic
   Synthetica BlueSteel theme      ->  synthetica.theme.bluesteel
   Synthetica BlackMoon theme      ->  synthetica.theme.blackmoon
   Synthetica BlueMoon theme       ->  synthetica.theme.bluemoon
   Synthetica SilverMoon theme     ->  synthetica.theme.silvermoon
   Synthetica BlueIceMoon theme    ->  synthetica.theme.blueice
   Synthetica GreenDream theme     ->  synthetica.theme.greendream
   Synthetica BlackStar theme      ->  synthetica.theme.blackstar
   Synthetica Standard theme       ->  synthetica.theme.standard
   

5. JPMS Integration

   Generally you only have to add the modules of the required themes to your 
   application. By not using the 'static' clause all specified theme modules
   are mandatory to run your application. Which means at compile and run
   time. 
   
   With using the 'static' clause all specified theme modules are optional at
   run time. However, in this case you have to tell the JVM which of the theme 
   modules should be added to the module graph e.g. by using '--add-modules'.  

   Example for Plain theme usage:
   
    java --add-exports=java.desktop/sun.swing=synthetica.base --add-exports=java.desktop/sun.swing.table=synthetica.base --add-exports=java.desktop/sun.swing.plaf.synth=synthetica.base --add-opens=java.desktop/javax.swing.plaf.synth=synthetica.base --add-opens=java.desktop/javax.swing.plaf.basic=synthetica.base --add-opens=java.desktop/javax.swing=synthetica.base --add-opens=java.desktop/javax.swing.tree=synthetica.base --add-opens=java.desktop/java.awt.event=synthetica.base --add-exports=java.desktop/sun.awt.shell=synthetica.base --add-exports=java.base/sun.security.action=synthetica.base
    --module-path synthetica.jar;syntheticaPlain.jar;syntheticaDemo.jar --add-modules synthetica.theme.plain -m synthetica.demo/de.javasoft.synthetica.demo.Demo        

    Demo module-info for Standard/Plain/Dark theme support:
    
    module synthetica.demo
    {
      exports de.javasoft.synthetica.demo;
      requires static synthetica.theme.standard;
      requires static synthetica.theme.plain;
      requires static synthetica.theme.dark;
    }   
   