package com.vatsytech.office;

import java.awt.Color;

public class ExcelCell {

	private int numRow;
	private int numCol;
	private String txtValue;
	private String txtComment;
        private Color backGround;
	
	
	public ExcelCell() {
		super();
	}
	
        public ExcelCell(int numRom, int numCol, String txtValue, String txtComment) {
		super();
		this.numRow = numRom;
		this.numCol = numCol;
		this.txtValue = txtValue;
		this.txtComment = txtComment;
	}
	
	public ExcelCell(String coordExcel, String txtValue, String txtComment) {
		super();
		this.numRow = OfficeUtils.getInstance().getNumRowFromCoord(coordExcel);
		this.numCol = OfficeUtils.getInstance().getNumCollFromCoord(coordExcel);
		this.txtValue = txtValue;
		this.txtComment = txtComment;
	}
	
	public int getNumRow() {
		return numRow;
	}
	public void setNumRow(int numRom) {
		this.numRow = numRom;
	}
	public int getNumCol() {
		return numCol;
	}
	public void setNumCol(int numCol) {
		this.numCol = numCol;
	}
	public String getTxtValue() {
		return txtValue;
	}
	public void setTxtValue(String txtValue) {
		this.txtValue = txtValue;
	}
	public String getTxtComment() {
		return txtComment;
	}
	public void setTxtComment(String txtComment) {
		this.txtComment = txtComment;
	}

    /**
     * @return the backGround
     */
    public Color getBackGround() {
        return backGround;
    }

    /**
     * @param backGround the backGround to set
     */
    public void setBackGround(Color backGround) {
        this.backGround = backGround;
    }
	
	
	
}
