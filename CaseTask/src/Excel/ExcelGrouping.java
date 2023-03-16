package Excel;

import org.apache.poi.ss.usermodel.*;

import java.util.*;

public class ExcelGrouping {

    // Создаем результирующую таблицу
    private static ArrayList<ArrayList<String>> resultTable = new ArrayList<>();

    public static ArrayList<ArrayList<String>> getResultTable() {
        return resultTable;
    }

    public static void setResultTable(ArrayList<ArrayList<String>> resultTable) {
        ExcelGrouping.resultTable = resultTable;
    }

    public static void grouping(){
        Scanner scan = new Scanner(System.in);

        // Открываем лист в файле excel
        Sheet sheet = ExcelOpen.open();

        // Создаем регуляр для фильтрации нежелательных символов
        String regex = "[^\\p{IsCyrillic}\\p{IsLatin}\\p{IsDigit}]";

        // Формируем результирующую таблицу
        assert sheet != null;
        ExcelFormationResultTable.formationResultTable(sheet, regex);

        System.out.print("Куда вывести сгруппированную таблицу: \n 1 - в консоль \n 2 - в excel файл" +
                "\n Ваш выбор: ");
        int answerUser = scan.nextInt();

        // Вывод результирующей таблицы
        switch(answerUser){
            case 1:
                ExcelPrint.printConsole();
                break;
            case 2:
                System.out.print("Введите путь к файлу excel в который записать данные: ");
                scan.nextLine();
                String pathOutFile = scan.nextLine();
                ExcelPrint.printExcel(pathOutFile);
                break;
            default:
                System.out.println("Ошибка. Такого способа вывода нет!");
                break;
        }
    }
}
