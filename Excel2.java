import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;




public class Excel2 {

    //private static final String FILE_NAME = "/tmp/MyFirstExcel.xlsx";

    public static void main(String[] args) throws IOException {

    	Workbook wb = new HSSFWorkbook();
        //Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");

        // Create a row and put some cells in it. Rows are 0 based.
//        Row row = sheet.createRow((short)0);
//        // Create a cell and put a value in it.
//        Cell cell = row.createCell(0);
//        cell.setCellValue(1);
//
//        // Or do it on one line.
//        row.createCell(1).setCellValue(1.2);
//        row.createCell(2).setCellValue(
//             ("This is a string"));
//        row.createCell(3).setCellValue(true);

        // Write the output to a file
        
        
        Object[][] datatypes = {
                {"Datatype", "Type", "Size(in bytes)"},
                {"int", "Primitive", 4},
                {"float", "Primitive", 4},
                {"double", "Primitive", 8},
                {"char", "Primitive", 1},
                {"String", "Non-Primitive", "No fixed size"}
        };

        int rowNum = 0;
        System.out.println("Creating excel");

        for (Object[] datatype : datatypes) {
            Row row = sheet.createRow(rowNum++);
            int colNum = 0;
            for (Object field : datatype) {
                Cell cell = row.createCell(colNum++);
                if (field instanceof String) {
                    cell.setCellValue((String) field);
                } else if (field instanceof Integer) {
                    cell.setCellValue((Integer) field);
                }
            }
        }

        
      
        
        FileOutputStream fileOut = new FileOutputStream("C:\\Users\\ankitmalik\\new2.xlsx");
        wb.write(fileOut);
        fileOut.close();
}
    }