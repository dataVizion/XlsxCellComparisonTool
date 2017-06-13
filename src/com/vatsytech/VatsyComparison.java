package com.vatsytech;

import com.vatsytech.ui.ComparisonFrame;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.UIManager;

public class VatsyComparison {
    
     public static void main(String[] args) {
        try {           
            //UIManager.setLookAndFeel("com.sun.java.swing.plaf.nimbus.NimbusLookAndFeel");
            UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
            ComparisonFrame a=new ComparisonFrame();

            a.setVisible(true);                                     
        } catch (Exception ex) {
            Logger.getLogger(VatsyComparison.class.getName()).log(Level.SEVERE, null, ex);
        }



    }


}
