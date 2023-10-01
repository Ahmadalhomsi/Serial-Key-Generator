/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package serialcodegenerator3;

import com.formdev.flatlaf.FlatDarculaLaf;
import com.formdev.flatlaf.FlatIntelliJLaf;
import com.formdev.flatlaf.themes.FlatMacDarkLaf;
import de.javasoft.synthetica.simple2d.SyntheticaSimple2DLookAndFeel;
import java.text.ParseException;
import java.util.*;
import javax.swing.UIManager;

/**
 *
 * @author pooqw
 */
public class Main {

    public static void main(String[] args) throws ParseException {

    // Look and feel  
    
        try {
    UIManager.setLookAndFeel( new FlatDarculaLaf()); //FlatMacDarkLaf() FlatDarculaLaf()
    } catch( Exception ex ) {
    System.err.println( "Failed to initialize LaF" );
    }
        
    //-------

        GUI gui = new GUI();
        gui.setVisible(true);

        Excel excel = new Excel();
        Scanner S = new Scanner(System.in);

        int slct = 0;

        System.out.println("1 for Reading");
        System.out.println("2 for Writing");
        System.out.println("3 for Updating");
        slct = S.nextInt();

        try {

            if (slct == 1) {
                //excel.ExcelReader("List.xlsx");
            }

            if (slct == 2) {
                excel.ExcelWrite("List.xlsx");
            }

            if (slct == 3) {
                //excel.ExcelUpdate("List.xlsx");
            }

        } catch (Exception e) {
            System.out.println(e);
        }
    }

}
