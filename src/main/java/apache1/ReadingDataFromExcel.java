package apache1;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

//Excel file---->workbook---->sheets---->Rows---->cells

public class ReadingDataFromExcel {

    public static void main(String[] args) throws IOException {

        FileInputStream file=new FileInputStream(System.getProperty("user.dir")+"\\testdata\\country_capital_population_sample.xlsx");

        XSSFWorkbook workbook=new XSSFWorkbook(file);

        XSSFSheet sheet=workbook.getSheet("sheet1");

        int totalrows=sheet.getLastRowNum(); //rows are counting from 0
        int totalCells=sheet.getRow(1).getLastCellNum(); //cells are counting from 1

        System.out.println("number of rows:"+totalrows);
        System.out.println("number of cells:"+totalCells);

        for(int i=0;i<=totalrows;i++){
            XSSFRow rows=sheet.getRow(i);
            for(int j=0;j<totalCells;j++){
                XSSFCell cell=rows.getCell(j);
                System.out.print(cell.toString());
            }
            System.out.println();
        }
        workbook.close();
        file.close();
    }

}
