
package com.vatsytech.office;

import java.awt.Component;
import javax.swing.JLabel;
import javax.swing.JTable;
import javax.swing.table.DefaultTableCellRenderer;

public class ExcelCellRenderer extends DefaultTableCellRenderer{


    
    @Override
    public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
         JLabel cell = (JLabel) super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);         
         ExcelCell a= (ExcelCell)table.getModel().getValueAt(row, column); 

        if(a!=null){
            cell.setText(a.getTxtValue());
            cell.setBackground(a.getBackGround());
            cell.setForeground(table.getForeground());
            cell.setToolTipText(null);
            if(a.getTxtComment()!=null ) {cell.setToolTipText(a.getTxtComment());}
            
        }
        else{
            cell.setText("");
            cell.setToolTipText(null);
            cell.setBackground(table.getBackground());
        }
        
        if(isSelected){
            cell.setBackground(table.getSelectionBackground());
            cell.setForeground(table.getSelectionForeground());
        } 
              
        
        
        return cell;
    }
    
}
