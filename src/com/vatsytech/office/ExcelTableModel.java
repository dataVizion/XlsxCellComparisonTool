
package com.vatsytech.office;

import javax.swing.table.AbstractTableModel;

public class ExcelTableModel extends AbstractTableModel{
    private ExcelCell[][] data;
    private String[] title;

    //Constructeur
    public ExcelTableModel(ExcelCell[][] data, String[] title){
      this.data = data;
      this.title = title;
    }

    //Retourne le nombre de colonnes
    public int getColumnCount() {
      return this.title.length;
    }

    //Retourne le nombre de lignes
    public int getRowCount() {
      return this.data.length;
    }
    
    /**
    * Retourne le titre de la colonne à l'indice spécifié
    */
    public String getColumnName(int col) {
      return this.title[col];
    }

    //Retourne la valeur à l'emplacement spécifié
    public Object getValueAt(int row, int col) {
      ExcelCell a= this.data[row][col];
      return a;
    }            
    
    public Class getColumnClass(int c) {
        //return getValueAt(0, c).getClass();
        return ExcelCell.class;
    }
  }

