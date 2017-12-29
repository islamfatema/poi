package poiDemo.poi;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.WorkbookUtil;
 
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
 
import java.io.FileOutputStream;


public class ReadExcel {

	public static void main(String[] args)  {
		 Workbook workbook = new HSSFWorkbook();
         
         //creates an empty sheet with the name "Sheet0".
         //for every new sheet the number at the end increases ("Sheet1", "Sheet2",...)
         Sheet sheet = workbook.createSheet("Rose");
         Row row=sheet.createRow(0);
         Cell cell=row.createCell(3);
        cell.setCellValue("Hi there");
         try {
                 FileOutputStream output = new FileOutputStream("Test1.xls");
              workbook.write(output);
                 output.close();
         } catch(Exception e) {
                 e.printStackTrace();
         }
	}

}
