import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        final String patch = "D:\\Projects\\Java\\PFR\\ZakonnyePredstavitely\\";
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(patch + "Законные представители.xlsx"));
            HashMap<String, Predstavitel> zakonPredstavitel = new HashMap<>();
            XSSFSheet zeroSheet = workbook.getSheetAt(0);
            //Iterate through each rows one by one
            Iterator<Row> rowIterator = zeroSheet.iterator();
            int counter = 1;
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                if (counter < 9) {
                    counter++;
                    continue;
                }
                zakonPredstavitel.put(row.getCell(13).getStringCellValue(),
                        new Predstavitel(row.getCell(1).getStringCellValue(), row.getCell(3).getStringCellValue()));

                counter++;
            }
            System.out.println("Total rows read: " + counter);
            if (zakonPredstavitel.containsKey("164-457-766 00")) {
                System.out.println(zakonPredstavitel.get("164-457-766 00").fio + " " +
                        zakonPredstavitel.get("164-457-766 00").snils);
            } else {
                System.out.println("Snils not found!");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
