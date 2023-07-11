package excel_POI_P;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class ReadExcelFileToList {


	 
    public static List<Country> readExcelData(String fileName) {
        List<Country> countriesList = new ArrayList<Country>();
        Workbook workb = new XSSFWorkbook();
        try {
            //Datei  ansprechen
            FileInputStream fis = new FileInputStream(fileName);
            
            
           // Workbook workb = null;
            if(fileName.toLowerCase().endsWith("xlsx")){
                workb = new XSSFWorkbook(fis);
            }else if(fileName.toLowerCase().endsWith("xls")){
                workb = new HSSFWorkbook(fis);
            }
            
            //Tabellenblatt auswählen
            int numberOfSheets = workb.getNumberOfSheets();
            
            //Schleife für Tabellenblätter
            for(int i=0; i < numberOfSheets; i++){
                
                //Auswahl Tabellenblatt
                Sheet sheet = (Sheet)workb.getSheetAt(i);
            
                
                //Zeile auswählen
                Iterator<Row> rowIterator = sheet.iterator();
                while (rowIterator.hasNext()) 
                {
                    String name = "";
                    String shortCode = "";
                    
                    //Zeilenobjekt 
                    Row row = (Row)rowIterator.next();
                    
                    //Jede Zeile hat Spalten, Spalten auswählen
                    Iterator<Cell> cellIterator = row.cellIterator();
                     
                    while (cellIterator.hasNext()) 
                    {
                        //Spalte 
                        Cell cell = cellIterator.next();
                        
                        //Spakten checken
                        switch(cell.getCellType()){
                        case STRING:
                          
                      // Fall Spalte Typ String 
                        
                            if(shortCode.equalsIgnoreCase("")){
                                shortCode = cell.getStringCellValue().trim();
                            }else if(name.equalsIgnoreCase("")){
                                //2nd column
                                name = cell.getStringCellValue().trim();
                            }else{
                                //random data, leave it
                                System.out.println("Random data::"+cell.getStringCellValue());
                            }
                            break;
                       
                            // Spalte nummerisch
                        case NUMERIC:
                            System.out.println("Random data::"+cell.getNumericCellValue());
                        case BLANK:
                            break;
                        case BOOLEAN:
                            break;
                        case ERROR:
                            break;
                        case FORMULA:
                            break;
                        case _NONE:
                            break;
                        default:
                            break;
                        }
                    } //Ende Spalten Iterator
                    Country c = new Country(name, shortCode);
                    countriesList.add(c);
                } //Ende Zeilen Iterator
                
                
            } //  Ende Tabellenblattschleife            
            //Schließen  input stream
            fis.close();
           
            
            
        } catch (IOException e) {
            e.printStackTrace();
        }
        
        return countriesList;
    }

    public static void main(String args[]){
        List<Country> list = readExcelData("C:\\test\\land.xlsx");
        System.out.println("Länder Liste\n"+list);
    }

}