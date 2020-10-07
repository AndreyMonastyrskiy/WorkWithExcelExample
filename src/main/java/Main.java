import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class Main {
    public static void main(String[] args) {
        final String patch = "D:\\Projects\\Java\\PFR\\ZakonnyePredstavitely\\";
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(patch + "Законные представители.xlsx"));
            XSSFSheet zeroSheet = workbook.getSheetAt(0);
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
