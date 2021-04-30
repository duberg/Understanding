package common;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.instrument.Instrumentation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.Collections;
import java.util.Iterator;
import java.util.Scanner;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

public class Main {
    public static void main(String[] args) throws Exception {
        disableAccessWarnings();

        String name = "Ozzy";

        //hitta filen
        String excelFilePath = "C:\\Users\\46793\\Documents\\Testea\\hej.xlsx";
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet firstSheet = workbook.getSheetAt(0);
        int rowCount = firstSheet.getLastRowNum();
        Row row = firstSheet.createRow(++rowCount);
        int columnCount = 0;

        Cell cell = row.createCell(columnCount);
        cell.setCellValue(name);
        inputStream.close();

        FileOutputStream outputStream = new FileOutputStream("C:\\Users\\46793\\Documents\\Testea\\hej.xlsx");
        workbook.write(outputStream);
        outputStream.close();
        System.out.println("hell yeah");


    }


    @SuppressWarnings("unchecked")
    public static void disableAccessWarnings() {
        try {
            Class unsafeClass = Class.forName("sun.misc.Unsafe");
            Field field = unsafeClass.getDeclaredField("theUnsafe");
            field.setAccessible(true);
            Object unsafe = field.get(null);

            Method putObjectVolatile = unsafeClass.getDeclaredMethod("putObjectVolatile", Object.class, long.class, Object.class);
            Method staticFieldOffset = unsafeClass.getDeclaredMethod("staticFieldOffset", Field.class);

            Class loggerClass = Class.forName("jdk.internal.module.IllegalAccessLogger");
            Field loggerField = loggerClass.getDeclaredField("logger");
            Long offset = (Long) staticFieldOffset.invoke(unsafe, loggerField);
            putObjectVolatile.invoke(unsafe, loggerClass, offset, null);
        } catch (Exception ignored) {
        }
    }






    }


























        /*try {
            //Workbook
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet spreadsheet = workbook.createSheet("Hello");
            String[] columnHeads = {"namn", "age"};

            Row headerRow = spreadsheet.createRow(0);

            for (int i = 0; i < columnHeads.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(columnHeads[i]);

            }

            String[] storingNames = new String[1];
            String[] name = new String[1];
            String [] age = new String[1];
            String[] gender = new String[1];
            int rownum = 1;
            Scanner scan = new Scanner(System.in);
            for(int i = 0; i < storingNames.length; i++){
                name[i] = scan.nextLine();
                age[i] = scan.nextLine();
                Row row2 = spreadsheet.createRow(rownum++);
                row2.createCell(0).setCellValue(name[i]);
                row2.createCell(1).setCellValue(age[i]);



            }


*/


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

/*

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


    }*/

   /* private static ArrayList<Competitor> createData() {

        ArrayList<Competitor> a = new ArrayList();
        a.add(new Competitor("Hanna", 25));
        a.add(new Competitor("Mahdi", 35));


        return a;

    }
    */


