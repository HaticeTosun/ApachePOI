import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

public class TestCases {
    private static final String FILE_NAME = "src/test/resources/Books.xlsx";

    // this method ignores blanks
    @Test
    public void test1() throws IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = datatypeSheet.iterator();

        while (iterator.hasNext()) {
            Row currentRow = iterator.next();
            Iterator<Cell> cellIterator = currentRow.iterator();

            while (cellIterator.hasNext()) {
                Cell currentCell = cellIterator.next();
                if (currentCell.getCellType() == CellType.STRING) {
                    System.out.print(currentCell.getStringCellValue() + " -- ");
                } else if (currentCell.getCellType() == CellType.NUMERIC) {
                    System.out.print(currentCell.getNumericCellValue() + " -- ");
                }
            }
            System.out.println();
        }
    }

    //this method reads all values
    @Test
    public void test2() throws IOException {
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        int lastRowNum = datatypeSheet.getLastRowNum();
        int firstRowNum = datatypeSheet.getFirstRowNum();
        int rowCount = lastRowNum - firstRowNum;
        int cellCount = datatypeSheet.getRow(1).getLastCellNum() - datatypeSheet.getRow(1).getFirstCellNum();

        for (int i = 0; i < rowCount; i++) {
            Row row = datatypeSheet.getRow(i);
            if(row != null) {
                for (int j = 0; j < cellCount; j++) {
                    if(row.getCell(j) != null)
                        System.out.print(row.getCell(j).toString() + " --- ");
                }
            }
            System.out.println();
        }
    }

}
