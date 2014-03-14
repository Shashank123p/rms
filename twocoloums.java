
import java.io.File;
import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Statement;
import java.util.Iterator;
import java.util.Scanner;
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
public class twocoloums {
     public static void main(String[] args) 
    {
        int i,j=0;
        while(true)
        {
        System.out.println("enter your choice");
        Scanner sc= new Scanner(System.in);
        System.out.println("1.uploadin excelsheet\n2.retreivng info");
        int ch=sc.nextInt();
        
        switch(ch)
        {
            case 1:
                
                    try
                    {
            
                        String connectionURL = "jdbc:mysql://localhost:3306/apachi";
                        FileInputStream file = new FileInputStream(new File("C:\\Users\\Nishanth\\Desktop\\Book2.xlsx"));
 
                        
                        XSSFWorkbook workbook = new XSSFWorkbook(file);
                        
                        
                        XSSFSheet sheet = workbook.getSheetAt(0);
                        Iterator<Row> ir = sheet.iterator();
                       Row r=ir.next();
                       System.out.println(r.getLastCellNum());
            
            
                    switch(r.getLastCellNum())
                    {
                        case 2:
                            Iterator<Row> rowIterator = sheet.iterator();                
                            String sql="insert into twofields values (?";
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
                                
                                for(int k=min;k<max;k++)
                                {
                                    Cell cell=row.getCell(k);
                                    if(cell==null)
                                    {
                                        i=i+1;
                                       
                                        pst.setString(i, null);       // if excel sheet has blank shell then we will commit as null value into datanbase
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
                                                    
                                                    break;
                                            case Cell.CELL_TYPE_STRING:
                                                     i=i+1;
                                                    pst.setString(i, cell.getStringCellValue());
                                                    System.out.print(cell.getStringCellValue() + "\t");
                                                   
                                                    break;
                                        }
                                    }
                                }
                                System.out.println(" ");
                                file.close();
                                int nnum=pst.executeUpdate();
                                pst.close();
                                }
                               break;

                            }
                        }
                                    catch (Exception e) 
                            {
                                e.printStackTrace();
                            }
                  break;
            case 2:
                System.out.println("enter the info you want");
                String s=sc.next();
                try {
                        Class.forName("com.mysql.jdbc.Driver");

                        Connection con = DriverManager.getConnection("jdbc:mysql://localhost:3306/apachi", "root", "root");
                        Statement st = con.createStatement();
                        String sql = "select * from twofields where column1='" + s + "'";
                        ResultSet rs=st.executeQuery(sql);
                        
                        while(rs.next())
                        {
                            
                            String col=rs.getString("column2");
                            System.out.println(col);
                        }
                        
                }
                catch(Exception e)
                {
                    e.printStackTrace();
                }
                break;
        }
        }
    }
}
