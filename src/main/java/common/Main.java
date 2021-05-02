package common;


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.lang.instrument.Instrumentation;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Collections;
import java.util.Iterator;
import java.util.Scanner;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

public class Main {
    public static void main(String[] args) throws Exception {
        disableWarnings();
        Scanner scan = new Scanner(System.in);
        Scanner intScan = new Scanner(System.in);

        File tempFile = new File("C:\\Users\\46793\\Documents\\Testea\\hej.xlsx");
        boolean exists = tempFile.exists();
        System.out.println(exists);


        if (exists == true) {



            //hitta filen
            String excelFilePath = "C:\\Users\\46793\\Documents\\Testea\\hej.xlsx";
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));

            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            XSSFSheet firstSheet = workbook.getSheetAt(0);

            DataFormatter formatter = new DataFormatter();
            System.out.println("Skriv 1 om du vill ändra något");
            int choice = intScan.nextInt();

            if (choice == 1) {

                //Byta ut en cell till en annan
                System.out.println("Vilket namn vill du byta ut?");
                String searchName = scan.nextLine();
                System.out.println("Vad vill du ändra det till?");
                String changeName = scan.nextLine();
                for (XSSFSheet sheet : workbook) {
                    for (Row row : sheet) {
                        for (Cell cell : row) {
                            if (formatter.formatCellValue(cell).contains(searchName)) {
                                cell.setCellValue(changeName);
                            }
                        }
                    }
                    FileOutputStream outputStream = new FileOutputStream("C:\\Users\\46793\\Documents\\Testea\\hej.xlsx");
                    workbook.write(outputStream);
                    outputStream.close();
                    System.out.println("ändrat");
                }
            } else {

                int rowCount = firstSheet.getLastRowNum();


                int k = 0;

                while (k != 3) {


                    Row row2 = firstSheet.createRow(++rowCount);
                    if(rowCount == 5){
                        System.out.println("I'm sorry");
                        k = 3;
                    }else {
                        int columnCount = 0;
                        Cell cell2 = row2.createCell(columnCount);
                        System.out.println("Enter your name:");
                        String name = scan.nextLine();
                        // System.out.println("Enter your age");
                        //String age2 = scan.nextLine();
                        cell2.setCellValue(name);
                        //cell2 = row2.createCell(columnCount + 1);
                        //cell2.setCellValue(age2);
                        inputStream.close();
                        System.out.println("type 3 to exit, or anything else to add another");
                        k = intScan.nextInt();
                    }

                }

                FileOutputStream outputStream = new FileOutputStream("C:\\Users\\46793\\Documents\\Testea\\hej.xlsx");
                workbook.write(outputStream);
                outputStream.close();
                System.out.println("hell yeah");


            }
        }
        if (exists == false) {

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
            String[] name2 = new String[1];
            String[] age = new String[1];
            String[] gender = new String[1];
            int rownum = 1;
            for (int i = 0; i < storingNames.length; i++) {
                System.out.println("Please enter your name:");
                name2[i] = scan.nextLine();
                System.out.println("Please enter your age");
                age[i] = scan.nextLine();
                Row row2 = spreadsheet.createRow(rownum++);
                row2.createCell(0).setCellValue(name2[i]);
                row2.createCell(1).setCellValue(age[i]);


            }
            FileOutputStream outputStream = new FileOutputStream("C:\\Users\\46793\\Documents\\Testea\\hej.xlsx");
            workbook.write(outputStream);
            outputStream.close();
            System.out.println("andra");
        }





    }

    @SuppressWarnings("unchecked")
    private static void disableWarnings() {

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


