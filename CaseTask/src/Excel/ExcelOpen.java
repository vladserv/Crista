package Excel;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.Scanner;

public class ExcelOpen {
    public static Sheet open(){
        Scanner scan = new Scanner(System.in);
        try {
            System.out.print("Введите путь к файлу excel: ");
            String filePath = scan.nextLine();

            // Открываем файл Excel
            FileInputStream fis = new FileInputStream(filePath);
            Workbook workbook = WorkbookFactory.create(fis);

            System.out.print("Введите номер листа, который хотите обработать (нумерация начинается с 1): ");
            int numSheet = scan.nextInt() - 1; // Для удобства пользователя нумерация начинается с 1

            return workbook.getSheetAt(numSheet);
        }catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }
}
