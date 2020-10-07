import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
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
            workbook.close();
            System.out.println("Total rows read: " + counter);
           /* if (zakonPredstavitel.containsKey("164-457-766 00")) {
                System.out.println(zakonPredstavitel.get("164-457-766 00").fio + " " +
                        zakonPredstavitel.get("164-457-766 00").snils);
            } else {
                System.out.println("Snils not found!");
            }*/
            System.out.println("Work with мамки_1_ребенок.xls");
            HSSFWorkbook smol1 = new HSSFWorkbook(new FileInputStream(patch + "Смоленск\\мамки_1_ребенок.xls"));
            HSSFSheet sml1Sheet = smol1.getSheetAt(0);
            counter = 0;
            rowIterator = sml1Sheet.iterator();
            while (rowIterator.hasNext())
            {
                Row row = rowIterator.next();
                if (counter < 1) {
                    Cell fioCell = row.createCell(11, CellType.STRING);
                    Cell snilsCell = row.createCell(12, CellType.STRING);
                    fioCell.setCellValue("ФИО законного представителя");
                    snilsCell.setCellValue("СНИЛС законного представителя");
                    counter++;
                    continue;
                }
                if (zakonPredstavitel.containsKey(row.getCell(4).getStringCellValue())) {
                    Cell fioCell = row.createCell(11, CellType.STRING);
                    Cell snilsCell = row.createCell(12, CellType.STRING);
                    Predstavitel predstavitel = zakonPredstavitel.get(row.getCell(4).getStringCellValue());
                    fioCell.setCellValue(predstavitel.fio);
                    snilsCell.setCellValue(predstavitel.snils);
                }
                counter++;
            }
            System.out.println("Total rows read: " + counter);
            smol1.write(new FileOutputStream(patch + "out\\Smolensk\\мамки_1_ребенок.xls"));
            System.out.println("Completed");

        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
