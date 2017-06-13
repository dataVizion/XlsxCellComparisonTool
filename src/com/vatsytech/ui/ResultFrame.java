package com.vatsytech.ui;

import com.vatsytech.office.ExcelCell;
import com.vatsytech.office.ExcelCellRenderer;
import com.vatsytech.office.ExcelClipboardAdapter;
import com.vatsytech.office.ExcelTableModel;
import com.vatsytech.office.OfficeUtils;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.table.TableModel;

public class ResultFrame extends javax.swing.JFrame {
    
    private TableModel model;
    private int totalErr;

    /**
     * Creates new form ResultFrame
     */
    public ResultFrame() {
        initComponents();
         ExcelClipboardAdapter c=new ExcelClipboardAdapter(jTableRes);
    }
    
    public ResultFrame(ExcelCell[][] m) {                
        this.setModel(m);
        initComponents();        
        jTableRes.setDefaultRenderer(ExcelCell.class, new ExcelCellRenderer());
        jTableRes.setModel(model);
        ExcelClipboardAdapter c=new ExcelClipboardAdapter(jTableRes);

        JTable rowTable = new RowNumberTable(jTableRes);
        jScrollPane1.setRowHeaderView(rowTable);
        jScrollPane1.setCorner(JScrollPane.UPPER_LEFT_CORNER,    rowTable.getTableHeader());
        
        //21/04/17
        //listens to cell selection and fills two text areas with the values compared
        jTableRes.getSelectionModel().addListSelectionListener(new RowListener());
        jTableRes.getColumnModel().getSelectionModel().addListSelectionListener(new RowListener());
        
        
    }
    
    //listens on cells selections
    private class RowListener implements ListSelectionListener {
        public void valueChanged(ListSelectionEvent event) {
            if (event.getValueIsAdjusting()) {
                return;
            }
             ExcelCell a= (ExcelCell)jTableRes.getModel().getValueAt(jTableRes.getSelectionModel().getLeadSelectionIndex(), jTableRes.getColumnModel().getSelectionModel().getLeadSelectionIndex());
             if(a!=null){
                 jTxt_ValueA.setText(a.getTxtValue());
                 if(a.getTxtComment()!=null) {
                    jTxt_ValueB.setText(a.getTxtComment().substring(a.getTxtComment().indexOf("</u><br>")+8, a.getTxtComment().length()).replace("</html>", ""));                                          
                 }
                 else jTxt_ValueB.setText(a.getTxtValue());
             }
             else{
                 jTxt_ValueA.setText("");jTxt_ValueB.setText("");
             }
        }
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jScrollPane1 = new javax.swing.JScrollPane();
        jTableRes = new javax.swing.JTable();
        jLabelNumberOfDiff = new javax.swing.JLabel();
        jTFNumberOfDiff = new javax.swing.JTextField();
        jLabelFileA = new javax.swing.JLabel();
        jScrollPane2 = new javax.swing.JScrollPane();
        jTxt_ValueA = new javax.swing.JTextArea();
        jLabelFileB = new javax.swing.JLabel();
        jScrollPane3 = new javax.swing.JScrollPane();
        jTxt_ValueB = new javax.swing.JTextArea();

        setDefaultCloseOperation(javax.swing.WindowConstants.DISPOSE_ON_CLOSE);

        jTableRes.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null},
                {null, null, null, null}
            },
            new String [] {
                "Title 1", "Title 2", "Title 3", "Title 4"
            }
        ));
        jTableRes.setAutoResizeMode(javax.swing.JTable.AUTO_RESIZE_OFF);
        jTableRes.setCellSelectionEnabled(true);
        jTableRes.setCursor(new java.awt.Cursor(java.awt.Cursor.DEFAULT_CURSOR));
        jScrollPane1.setViewportView(jTableRes);
        jTableRes.getColumnModel().getSelectionModel().setSelectionMode(javax.swing.ListSelectionModel.SINGLE_INTERVAL_SELECTION);

        jLabelNumberOfDiff.setText("Differences found :");

        jTFNumberOfDiff.setEditable(false);
        jTFNumberOfDiff.setText(String.valueOf(totalErr)
        );
        jTFNumberOfDiff.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTFNumberOfDiffActionPerformed(evt);
            }
        });

        jLabelFileA.setText("File A :");

        jTxt_ValueA.setEditable(false);
        jTxt_ValueA.setColumns(20);
        jTxt_ValueA.setRows(1);
        jScrollPane2.setViewportView(jTxt_ValueA);

        jLabelFileB.setText("File B :");

        jTxt_ValueB.setEditable(false);
        jTxt_ValueB.setColumns(20);
        jTxt_ValueB.setRows(1);
        jScrollPane3.setViewportView(jTxt_ValueB);

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addGap(6, 6, 6)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 0, Short.MAX_VALUE))
            .addGroup(layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabelNumberOfDiff)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING)
                        .addComponent(jLabelFileB)
                        .addComponent(jLabelFileA)))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jTFNumberOfDiff, javax.swing.GroupLayout.PREFERRED_SIZE, 166, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.TRAILING, false)
                        .addComponent(jScrollPane3, javax.swing.GroupLayout.Alignment.LEADING, javax.swing.GroupLayout.DEFAULT_SIZE, 310, Short.MAX_VALUE)
                        .addComponent(jScrollPane2, javax.swing.GroupLayout.Alignment.LEADING)))
                .addContainerGap(40, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jScrollPane1, javax.swing.GroupLayout.DEFAULT_SIZE, 138, Short.MAX_VALUE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.BASELINE)
                    .addComponent(jTFNumberOfDiff, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                    .addComponent(jLabelNumberOfDiff))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabelFileA)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 49, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addGroup(layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jLabelFileB)
                    .addComponent(jScrollPane3, javax.swing.GroupLayout.PREFERRED_SIZE, 49, javax.swing.GroupLayout.PREFERRED_SIZE))
                .addContainerGap())
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    public void setjLabelFileAText(String s) {
        this.jLabelFileA.setText(s);
    }


    public void setjLabelFileBText(String s) {
        this.jLabelFileB.setText(s);
    }

    private void jTFNumberOfDiffActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTFNumberOfDiffActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTFNumberOfDiffActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(ResultFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(ResultFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(ResultFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(ResultFrame.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new ResultFrame().setVisible(true);
            }
        });
    }
    
        /**
     * @return the model
     */
    public TableModel getModel() {
        return model;
    }
    public void setModel(ExcelCell[][] data) {      
        String[] title=new String[data[0].length];
        totalErr=0;
        for(int i=0;i<data[0].length;i++) {
            title[i]=OfficeUtils.getInstance().getCollName(i-1);
            if(i!=0) totalErr+=Integer.valueOf(data[0][i].getTxtValue());
        }
        title[0]="";
        //System.out.println("Number of rows = "+data.length+"Total err="+totalErr);
        ExcelTableModel  mod = new ExcelTableModel((ExcelCell[][])data, title); 
        this.model = mod;        
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JLabel jLabelFileA;
    private javax.swing.JLabel jLabelFileB;
    private javax.swing.JLabel jLabelNumberOfDiff;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JScrollPane jScrollPane3;
    private javax.swing.JTextField jTFNumberOfDiff;
    private javax.swing.JTable jTableRes;
    private javax.swing.JTextArea jTxt_ValueA;
    private javax.swing.JTextArea jTxt_ValueB;
    // End of variables declaration//GEN-END:variables
}