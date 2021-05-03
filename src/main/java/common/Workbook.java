package common;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
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

public class Workbook {


    Scanner scan = new Scanner(System.in);
    Scanner intScan = new Scanner(System.in);

    File tempFile = new File("C:\\Users\\46793\\Documents\\Testea\\hej.xlsx");
    boolean exists = tempFile.exists();


    public void enterWorkbook() {
        try {
            if (exists == true) {


                //hitta filen
                String excelFilePath = "C:\\Users\\46793\\Documents\\Testea\\hej.xlsx";
                FileInputStream inputStream = new FileInputStream(new File(excelFilePath));


                XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

                XSSFSheet firstSheet = workbook.getSheetAt(0);

                firstSheet = workbook.getSheetAt(0);     //creating a Sheet object to retrieve object
                Iterator<Row> itr = firstSheet.iterator();    //iterating over excel file
                while (itr.hasNext()) {
                    Row row = itr.next();
                    Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
                    while (cellIterator.hasNext()) {
                        Cell cell = cellIterator.next();
                        switch (cell.getCellType()) {
                            case Cell.CELL_TYPE_STRING:    //field that represents string cell type
                                System.out.print(cell.getStringCellValue() + "\t\t\t");
                                break;
                            case Cell.CELL_TYPE_NUMERIC:    //field that represents number cell type
                                System.out.print(cell.getNumericCellValue() + "\t\t\t");
                                break;
                            default:
                        }
                    }
                    System.out.println("");
                }

                DataFormatter formatter = new DataFormatter();
                System.out.println("Type 1 if you'd like to change something then 'enter'");
                System.out.println("type any other digit to add another contestant, then 'enter'");
                int choice = intScan.nextInt();

                if (choice == 1) {

                    //Byta ut en cell till en annan
                    System.out.println("Which name/age would you like to change?");
                    String searchName = scan.nextLine();
                    System.out.println("What do you wanna change it to?");
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
                        System.out.println("changed");
                    }
                } else {

                    int rowCount = firstSheet.getLastRowNum();


                    int k = 0;

                    while (k != 3) {


                        Row row2 = firstSheet.createRow(++rowCount);
                        if (rowCount == 5) {
                            System.out.println("I'm sorry, the list is full");
                            k = 3;
                        } else {
                            int columnCount = 0;
                            Cell cell2 = row2.createCell(columnCount);
                            System.out.println("Enter your name:");
                            String name = scan.nextLine();
                            System.out.println("Enter your age");
                            String age2 = scan.nextLine();
                            cell2.setCellValue(name);
                            cell2 = row2.createCell(columnCount + 1);
                            cell2.setCellValue(age2);
                            try {
                                inputStream.close();
                            } catch (IOException e) {
                                e.printStackTrace();
                            }
                            System.out.println("type 3 to exit, or type any other digit then 'enter' to add another contestant");
                            k = intScan.nextInt();
                        }

                    }

                    FileOutputStream outputStream = new FileOutputStream("C:\\Users\\46793\\Documents\\Testea\\hej.xlsx");

                    workbook.write(outputStream);

                    outputStream.close();

                    System.out.println("Contestant/s added");


                }
            }
            if (!exists) {

                //Workbook
                XSSFWorkbook workbook = new XSSFWorkbook();
                XSSFSheet spreadsheet = workbook.createSheet("Hello");
                String[] columnHeads = {"Name:", "Age:"};

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

                System.out.println("Created new excel");
            }


        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
