package common;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Scanner;

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

            String[] storingNames = new String[2];
            String[] name = new String[2];
            String [] age = new String[2];
            String[] gender = new String[2];
            int rownum = 1;
            Scanner scan = new Scanner(System.in);
            for(int i = 0; i < storingNames.length; i++){
                name[i] = scan.nextLine();
                age[i] = scan.nextLine();
                Row row2 = spreadsheet.createRow(rownum++);
                row2.createCell(0).setCellValue(name[i]);
                row2.createCell(1).setCellValue(age[i]);



            }





           /* Row row = spreadsheet.createRow(6);
            Cell cell = row.createCell(10);
            cell.setCellValue(name);
*/


            /*int rownum = 1;
            for (Competitor i : a) {
                Row row2 = spreadsheet.createRow(rownum++);
                row2.createCell(0).setCellValue(i.name);
                row2.createCell(1).setCellValue(i.age);
            }*/



            for (int i = 0; i < columnHeads.length; i++) {
                spreadsheet.autoSizeColumn(i);
            }
            FileOutputStream fileOut = new FileOutputStream(new File("C:\\Users\\46793\\Documents\\TestExcel\\hej.xlsx"));
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("Excel file created");


        } catch (Exception e) {
            e.printStackTrace();
        }


    }

   /* private static ArrayList<Competitor> createData() {

        ArrayList<Competitor> a = new ArrayList();
        a.add(new Competitor("Hanna", 25));
        a.add(new Competitor("Mahdi", 35));


        return a;

    }
    */
}
