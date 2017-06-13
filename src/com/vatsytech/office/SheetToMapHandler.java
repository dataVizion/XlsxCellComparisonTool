package com.vatsytech.office;

import java.util.HashMap;
import java.util.Map;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

public class SheetToMapHandler implements SheetContentsHandler{
	private boolean firstCellOfRow = false;
	Map<String,String> mres=new HashMap<String,String>();

	@Override
	public void cell(String arg0, String arg1, XSSFComment arg2) {
            if (firstCellOfRow) {
                firstCellOfRow = false;
            } else {    
                
        	mres.put(arg0, arg1);
            }		
	}

	@Override
	public void endRow(int arg0) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void headerFooter(String arg0, boolean arg1, String arg2) {
		// TODO Auto-generated method stub
		
	}

	@Override
	public void startRow(int arg0) {
		// TODO Auto-generated method stub		
	}
	

	
	public Map<String,String> getRes(){
		return mres;
	}
    
}
