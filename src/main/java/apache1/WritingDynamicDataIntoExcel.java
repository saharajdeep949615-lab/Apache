package apache1;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

public class WritingDynamicDataIntoExcel {
    public static void main(String[] args) throws IOException {

        FileOutputStream file=new FileOutputStream(System.getProperty("user.dir")+"\\testdata\\myfiledaynamic.xlsx");
        XSSFWorkbook workbook=new XSSFWorkbook();

        XSSFSheet sheet=workbook.createSheet();

        Scanner sc=new Scanner(System.in);

        System.out.println("Enter how many rows");
        int noOfRows=sc.nextInt();
        System.out.println("Enter how many columns");
        int noOfColumns=sc.nextInt();

        for(int i=0;i<=noOfRows;i++){
            XSSFRow row=sheet.createRow(i);
            for(int j=0;j<noOfColumns;j++){
                XSSFCell cell=row.createCell(j);
                cell.setCellValue(sc.next());
            }
        }

        workbook.write(file);
        workbook.close();
        file.close();

        System.out.println("File is created.....");
    }
}
