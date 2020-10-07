import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

public class Main {
    public static void main(String[] args) {
        final String patch = "D:\\Projects\\Java\\PFR\\ZakonnyePredstavitely\\";
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(new FileInputStream(patch + "Законные представители.xlsx"));
            HashMap<String, Predstavitel> zakonPredstavitel = new HashMap<>();
            XSSFSheet zeroSheet = workbook.getSheetAt(0);
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

            System.out.println("Work with мамки_многодетки.xls");
            HSSFWorkbook smol2 = new HSSFWorkbook(new FileInputStream(patch + "Смоленск\\мамки_многодетки.xls"));
            HSSFSheet sml2Sheet = smol2.getSheetAt(0);
            counter = 0;
            rowIterator = sml2Sheet.iterator();
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
            smol2.write(new FileOutputStream(patch + "out\\Smolensk\\мамки_многодетки.xls"));

            System.out.println("Work with files in 1 folder...");
            for (final File fileEntry : new File(patch + "1\\").listFiles()) {
                System.out.println("Work with file: " + fileEntry.getName());

                HSSFWorkbook hssfWorkbook = new HSSFWorkbook(new FileInputStream(fileEntry.getCanonicalPath()));
                HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(0);
                counter = 0;
                rowIterator = hssfSheet.iterator();
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
                hssfWorkbook.write(new FileOutputStream(patch + "out\\1\\" + fileEntry.getName()));

            }


            System.out.println("Completed");

        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
