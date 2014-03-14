
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
           Iterator<Row> ir = sheet.iterator();
           Row r=ir.next();
           System.out.println(r.getLastCellNum());
            
            
            switch(r.getLastCellNum())
            {
                case 4:
                    Iterator<Row> rowIterator = sheet.iterator();                
                    String sql="insert into fourfields values (?";
                    for(int p=0;p<r.getLastCellNum()-1;p++)
                    {
                      sql+=",?";
                    }
                    sql+=")";
                    while (rowIterator.hasNext()) 
                    {
                        i=0;
                        Class.forName("com.mysql.jdbc.Driver");
                        Connection connection = DriverManager.getConnection(connectionURL, "root", "root");
                        Statement s= connection.createStatement();

                        PreparedStatement pst = connection.prepareStatement(sql);
                        Row row = rowIterator.next();
                        
                        Iterator<Cell> cellIterator = row.cellIterator();
                        int min=row.getFirstCellNum();
                        int max=row.getLastCellNum();
                         System.out.println("firstcell"+min+"lastcell"+max);
                        for(int k=min;k<max;k++)
                        {
                            Cell cell=row.getCell(k);
                            if(cell==null)
                            {
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
