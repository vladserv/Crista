package Excel;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.util.ArrayList;

public class ExcelPrint {
    public static void printConsole(){
        ArrayList<ArrayList<String>> rList;

        rList = ExcelGrouping.getResultTable();

        // Проходим по всей результирующей таблице и выводим ее в консоль
        for(ArrayList<String> foreachList : rList){
            for(String foreachString : foreachList){
                System.out.printf("%-15s", foreachString);
            }
            System.out.println();
        }
    }

    public static void printExcel(String pathFileOut){
        ArrayList<ArrayList<String>> rList;

        rList = ExcelGrouping.getResultTable();

        // Создаем новый Excel файл и рабочую книгу
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Данные");

        int indexRow = 0;
        int indexCell = 0;
        Row row = sheet.createRow(indexRow);
        for(ArrayList<String> foreachList : rList) {
            for (String foreachString : foreachList) {
                row.createCell(indexCell).setCellValue(foreachString);
                indexCell++;
            }
            indexCell = 0;
            row = sheet.createRow(++indexRow);
        }
        try {
            FileOutputStream fileOut = new FileOutputStream(pathFileOut);
            workbook.write(fileOut);
            fileOut.close();
            System.out.println("Запись прошла успешно");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
