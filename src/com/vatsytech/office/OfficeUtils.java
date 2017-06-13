
package com.vatsytech.office;

import java.awt.Color;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFReader.SheetIterator;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

public class OfficeUtils {
    private OfficeUtils () {
		
	}

    private static OfficeUtils INSTANCE = new OfficeUtils();
	
    public static OfficeUtils getInstance()
    {	return INSTANCE;
    }
    
    public Map<String,String> sheetToMap(File f,int sNum) throws InvalidFormatException, IOException, SAXException, OpenXML4JException, ParserConfigurationException{
        long chrono=java.lang.System.currentTimeMillis(),chrono2=0, temps = 0; 
        Map<String,String> resA;
        
        OPCPackage p = OPCPackage.open(f.getPath(), PackageAccess.READ);
        
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(p);
        XSSFReader xssfReader = new XSSFReader(p);
	       
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        int i=-1;InputStream stream=null;
        while(iter.hasNext() && i<sNum){
         stream = iter.next();
        i++;
        }
        InputSource sheetSource = new InputSource(stream);
        
        XMLReader sheetParser = SAXHelper.newXMLReader();
	StylesTable styles = xssfReader.getStylesTable();
	DataFormatter formatter = new DataFormatter();
        
        SheetToMapHandler sheetHandler=new SheetToMapHandler();
        ContentHandler handler = new XSSFSheetXMLHandler(styles, null, strings, sheetHandler, formatter, true);
        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
        resA=sheetHandler.getRes();
        
        p.close();
        System.gc();  
        
        chrono2 = java.lang.System.currentTimeMillis() ; 
	temps = chrono2 - chrono ; 
        System.out.println(f.getName()+" processed in :"+temps+" ms");
        return resA;
    }
    
    /*
    * Make use of method "sheetToMap"
    *@param f File from which extraction is needed
    */
    public Object[][] sheetToObjects(File f,int sNum)throws InvalidFormatException, IOException, SAXException, OpenXML4JException, ParserConfigurationException{     
        int maxRow=0,maxCol=0,temp=0,i=0,row=0,coll=0;   
        
        Map<String,String> res=sheetToMap(f,sNum);
        Set<Entry<String, String>> setHm = res.entrySet();
	Iterator<Entry<String, String>> it = setHm.iterator();        
        
        //Get size
        while(it.hasNext()){
            Entry<String, String> e = it.next();
            ExcelCell ac = new ExcelCell(e.getKey(),e.getValue(),"");
            temp=ac.getNumRow();
            maxRow=(temp>maxRow) ? temp:maxRow;
            temp=ac.getNumCol();
            maxCol=(temp>maxCol) ? temp:maxCol;
        }
        
        //Init collection
        Object[][] r=new Object[maxRow][maxCol];
        
        //Reste iterator and retrieve data
        it = setHm.iterator();   
        while(it.hasNext()){
            Entry<String, String> e = it.next();
            ExcelCell ac = new ExcelCell(e.getKey(),e.getValue(),"");         
            row=ac.getNumRow()-1;
            coll=ac.getNumCol()-1;

            r[row][coll]=ac.getTxtValue();
        }
        
       
        return r;
    }

    /*
    * Make use of method "sheetToMap"
    *@param f File from which extraction is needed
    */
    public ExcelCell[][] sheetToExcelCells(File f,int sNum)throws InvalidFormatException, IOException, SAXException, OpenXML4JException, ParserConfigurationException{
        int maxRow=0,maxCol=0,temp=0,i=0,row=0,coll=0;   
        
        Map<String,String> res=sheetToMap(f,sNum);
        Set<Entry<String, String>> setHm = res.entrySet();
	Iterator<Entry<String, String>> it = setHm.iterator();        
        
        //Get size
        while(it.hasNext()){
            Entry<String, String> e = it.next();
            ExcelCell ac = new ExcelCell(e.getKey(),e.getValue(),"");
            temp=ac.getNumRow();
            maxRow=(temp>maxRow) ? temp:maxRow;
            temp=ac.getNumCol();
            maxCol=(temp>maxCol) ? temp:maxCol;
        }
        
        //Init collection
        ExcelCell[][] r=new ExcelCell[maxRow][maxCol];
        
        //Reset iterator and retrieve data
        it = setHm.iterator();   
        while(it.hasNext()){
            Entry<String, String> e = it.next();
            ExcelCell ac = new ExcelCell(e.getKey(),e.getValue(),null);                  
            row=ac.getNumRow()-1;
            coll=ac.getNumCol()-1;
            ac.setBackGround(null);
            //For test purpose only
            //if(row%2==0){ac.setBackGround(Color.RED);ac.setTxtComment("testComment");}
            
            r[row][coll]=ac;
        }
        
        return r;
    }
    
     public ExcelCell[][] compareFiles(File fA,int sheetA,File fB,int sheetB) throws IOException, SAXException, OpenXML4JException, InvalidFormatException, ParserConfigurationException,Exception {
        long chrono=java.lang.System.currentTimeMillis(),chrono2=0, temps = 0; 
        ExcelCell[][] res;
        String a,b,coord;
        int maxRow=0,maxCol=0,temp=0,i=0,row=0,coll=0; 
        
        Map<String,String> resA=sheetToMap(fA,sheetA);
        Map<String,String> resB=sheetToMap(fB,sheetB);
              
        //Map<String, String> resA2 = new TreeMap<String, String>(resA);
        Set<Entry<String, String>> setHm = resA.entrySet();
	Iterator<Entry<String, String>> it = setHm.iterator();  
        
        //Get size
        while(it.hasNext()){
            Entry<String, String> e = it.next();
            ExcelCell ac = new ExcelCell(e.getKey(),e.getValue(),"");
            temp=ac.getNumRow();
            maxRow=(temp>maxRow) ? temp:maxRow;
            temp=ac.getNumCol();
            maxCol=(temp>maxCol) ? temp:maxCol;
        }
        
        //Init collection
        res=new ExcelCell[maxRow+1][maxCol+1];
               
        //Copy and compare
        it = setHm.iterator();   
        while(it.hasNext()){
            Entry<String, String> e = it.next();
            ExcelCell ac = new ExcelCell(e.getKey(),e.getValue(),null);       
            row=ac.getNumRow();
            coll=ac.getNumCol();            
            a=e.getValue();b=resB.get(e.getKey());
            if(b==null || a.compareTo(b)!=0){
                ac.setTxtComment("<html><u><b>"+fB.getName()+"</b></u>"+"<br>"+String.valueOf(b)+"</html>");
                ac.setBackGround(Color.RED);
            }
            res[row][coll]=ac;
            //res[row][0] <--holds nb.diff per rows
            //res[0][row] <--holds nb.diff per columns
        }
        
        //this part resolves issues with NULL cells in fileA, as those are not retrieved by sheetToMap method
        //we check NULL cells in res[][], and set correct values inside
        for(i=1;i<maxRow+1;i++){
            for(int j=1;j<maxCol+1;j++){
                if(res[i][j]==null) {                    
                    coord=this.getCollName(j-1)+i;
                    ExcelCell ac = new ExcelCell(coord,null,null);  
                    b=resB.get(coord);
                    if(b!=null){                         
                        ac.setTxtComment("<html><u><b>"+fB.getName()+"</b></u>"+"<br>"+String.valueOf(b)+"</html>");
                        ac.setBackGround(Color.RED);
                    }
                    res[i][j]=ac;
                }                
            }
        }
        
        
        //this final part counts differences per rows and columns
        int errRow,errCol;
        
        chrono2 = java.lang.System.currentTimeMillis() ; 
	temps = chrono2 - chrono ; 
        System.out.println("Before Counting differences in :"+temps+" ms");
        
        for (i=0;i<maxRow+1;i++){
            errRow=0;
            for (int j=0;j<maxCol+1;j++){        
                //if(i==3 && j==25) System.out.println("2-25"+res[i][j].getTxtValue());
                if(res[i][j]!=null && res[i][j].getTxtComment()!=null) {
                    errRow++;
                };
            }
            ExcelCell ac= new ExcelCell("",String.valueOf(errRow),null); 
            ac.setBackGround(Color.LIGHT_GRAY);
            //if(errRow>0) ac.setBackGround(Color.magenta);
            res[i][0]=ac;
        }
        
        for (int j=0;j<maxCol+1;j++){
            errCol=0;
            for (i=0;i<maxRow+1;i++){                
                if(res[i][j]!=null && res[i][j].getTxtComment()!=null) {                    
                    errCol++;
                };
            }
            ExcelCell ac= new ExcelCell("",String.valueOf(errCol),null);  
            ac.setBackGround(Color.LIGHT_GRAY);
            //if(errCol>0) ac.setBackGround(Color.magenta);
            if(j==0) ac.setTxtValue("Delta");
            res[0][j]=ac;
        }
        
        chrono2 = java.lang.System.currentTimeMillis() ; 
	temps = chrono2 - chrono ; 
        System.out.println(fA.getName()+" compared to "+fB.getName()+" in :"+temps+" ms");
            
        return res;
    }

     

    

    
    
    /**
     *
     * @param i
     * @return
     */
    public String getCollName(int i) {
        if( i<0 ) {
        return "-"+getCollName(-i-1);
        }

        int quot = i/26;
        int rem = i%26;
        char letter = (char)((int)'A' + rem);
        if( quot == 0 ) {
            return ""+letter;
        } else {
            return getCollName(quot-1) + letter;
        }
    }
            
    
    /*
    *@param coordExcel Row and column number (e.g "A45")
    */
    public int getNumRowFromCoord(String coordExcel){
		Pattern pat = Pattern.compile("(\\d+)");
		Matcher m = pat.matcher(coordExcel);

		if (m.find()) return Integer.parseInt(m.group(1));
		else return 0;
	}

    /*
    *@param coordExcel Row and column number (e.g "A45")
    */
    public int getNumCollFromCoord(String coordExcel){
            Pattern pat = Pattern.compile("([A-Z]*)");
            Matcher m = pat.matcher(coordExcel);
            
            if (m.find()) {
                return getNumCollFromStringColl(m.group(1));
            }
            else return 0;
    }

    /*
    *@param str A string representing a column (e.g "45")
    */
    public int getNumCollFromStringColl(String str) {
            char[] ch  = str.toCharArray();
            int round=0;
            int r=0;
            int temp;
            for(char c : ch)
            {
                    temp = (int)c;
                    r=temp-64+round;
                    round+=26;
            }
            
            //System.out.println(str+" = "+r);
            return r;
    }
    
    
    public String[] getSheetNamesFromXlsx(File f) throws InvalidFormatException, IOException, SAXException, OpenXML4JException, ParserConfigurationException{
    OPCPackage p = OPCPackage.open(f.getPath(), PackageAccess.READ);
    XSSFReader xssfReader = new XSSFReader(p);
    
    SheetIterator iter=(SheetIterator) xssfReader.getSheetsData();   
    int i = 0;
    while(iter.hasNext()) {
        i++;
        iter.next();        
    }
    String[] r=new String[i];
    
    iter=(SheetIterator) xssfReader.getSheetsData();
    i=0;
    while(iter.hasNext()) {
       InputStream a=(InputStream) iter.next();
        InputSource b = new InputSource(a);
        r[i]=iter.getSheetName();
        i++;
    }

    p.close();
    return r;
    }
}
