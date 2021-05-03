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
        Scanner intChoice = new Scanner(System.in);
        Workbook min = new Workbook();
        int k = 0;
        while (k != 2){
            min.enterWorkbook();
            System.out.println("if you'd like to exit type 2 then 'enter'");
            System.out.println("if you'd like to run the program again, type any digits other than 2 then 'enter'");
            k = intChoice.nextInt();
            //loopen måste fixas för om man skapar dokument och sen inte går ur programmet och in igen, så förstår inte programmet att excelen redan skapats, är nog lättfixat om allting får en metod
        }
        System.out.println("see you next time");
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


