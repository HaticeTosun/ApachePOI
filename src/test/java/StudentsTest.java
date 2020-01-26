import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.Assert;
import org.testng.annotations.DataProvider;
import org.testng.annotations.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class StudentsTest {
    @Test
    public void test() throws IOException {
        FileInputStream excelFile = new FileInputStream( new File("src/test/resources/students.xlsx") );
        Workbook wb = new XSSFWorkbook( excelFile );

        Sheet sh = wb.getSheet( "data" );
        Assert.assertNotEquals( sh, null );

        int rowCount = sh.getLastRowNum() - sh.getFirstRowNum();
        System.out.println(rowCount);

        Row firstRow = sh.getRow( 0 );
        Assert.assertNotEquals( firstRow, null );

        int columnCount = firstRow.getLastCellNum() - firstRow.getFirstCellNum();
        System.out.println(columnCount);

        for(int i = 0; i < rowCount; i++) {
            Row currentRow = sh.getRow( i );
            for(int j = 0; j < columnCount; j++) {
                Cell cell = currentRow.getCell( j );
                System.out.println(cell.toString());
            }
        }

    }


    @DataProvider(name = "studentsData")
    public Object[][] studentsData(){
        return null;
    }
}
