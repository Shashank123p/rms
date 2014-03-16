import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.Statement;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Nishanth
 */
public class apachidynamic {
    public static void main(String[] args) 
    {
        int i,j=0;
        try
        {
            
            String connectionURL = "jdbc:mysql://localhost:3306/apachi";
            FileInputStream file = new FileInputStream(new File("C:\\Users\\Nishanth\\Desktop\\Book1.xlsx"));
 
            //Create Workbook instance holding reference to .xlsx file
            XSSFWorkbook workbook = new XSSFWorkbook(file);
            //System.out.println(workbook.getNumCellStyles());
 
            //Get first/desired sheet from the workbook
            XSSFSheet sheet = workbook.getSheetAt(0);
            int MAX_COLUMN=0;
       int rowStart = sheet.getFirstRowNum();
    int rowEnd = sheet.getLastRowNum();

    for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
       Row r = sheet.getRow(rowNum);

       int lastColumn = r.getLastCellNum();
       if (MAX_COLUMN < lastColumn)
       {
         MAX_COLUMN = lastColumn;  
       }
      
    }
          
            
            
            switch(MAX_COLUMN)
            {
                case 4:
                    Iterator<Row> rowIterator = sheet.iterator();                
                    String sql="insert into fourfields values (?, ?, ?, ?) ";
                    while (rowIterator.hasNext()) 
                    {
                        i=0;
                        Class.forName("com.mysql.jdbc.Driver");
                        Connection connection = DriverManager.getConnection(connectionURL, "root", "root");
                        Statement s= connection.createStatement();

                        PreparedStatement pst = connection.prepareStatement(sql);
                        Row row = rowIterator.next();
                        
                        Iterator<Cell> cellIterator = row.cellIterator();

                        for(int k=0;k<MAX_COLUMN;k++)
                        {
                            Cell cell=row.getCell(k);
                            if(cell==null)
                            {   
                                if (k=0)
                                {
                                System.out.println( "noSecondaryKeyWord");
                                pst.setString(i, "noSecondaryKeyWord"); 
                                }
                                i=i+1;
                                System.out.println( "found");
                                pst.setString(i, null);
                            }
                            else
                            {
                                switch (cell.getCellType()) 
                                {
                                    case Cell.CELL_TYPE_NUMERIC:
                                         i=i+1;
                                            int in= (int)cell.getNumericCellValue();
                                            String st=" "+in;
                                            pst.setString(i, st);
                                            System.out.print(cell.getNumericCellValue() + "\t");
                                            System.out.println("here"+i);
                                            break;
                                    case Cell.CELL_TYPE_STRING:
                                             i=i+1;
                                            pst.setString(i, cell.getStringCellValue());
                                            System.out.print(cell.getStringCellValue() + "\t");
                                            System.out.println("here"+i);
                                            break;
                                }
                            }
                        }
                       
                           System.out.println(" ");
                        file.close();
                        int nnum=pst.executeUpdate();
                        pst.close();
                        }
                        
                        
                    }
        }
        
    catch (Exception e) 
        {
            e.printStackTrace();
        }
    }
}
