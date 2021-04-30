package common;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;

public class Main {
    public static void main(String[] args) throws Exception {

        try {
            //Workbook
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet spreadsheet = workbook.createSheet("Hello");
            String[] columnHeads = {"namn", "age"};

            Row headerRow = spreadsheet.createRow(0);

            for (int i = 0; i < columnHeads.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columnHeads[i]);

            }


            Row row = spreadsheet.createRow(6);
            Cell cell = row.createCell(10);
            cell.setCellValue("Hejsan");


            ArrayList<Competitor> a = createData();
            int rownum = 1;
            for (Competitor i : a) {
                Row row2 = spreadsheet.createRow(rownum++);
                row2.createCell(0).setCellValue(i.name);
                row2.createCell(1).setCellValue(i.age);
            }
            for (int i = 0; i < columnHeads.length; i++) {
                spreadsheet.autoSizeColumn(i);
            }
            FileOutputStream fileOut = new FileOutputStream(new File("C:\\Users\\46793\\Documents\\Testea\\hej.xlsx"));
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("Excel file created");


        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private static ArrayList<Competitor> createData() {
        ArrayList<Competitor> a = new ArrayList();
        a.add(new Competitor("Hanna", 25));
        a.add(new Competitor("Mahdi", 35));


        return a;
    }
}
