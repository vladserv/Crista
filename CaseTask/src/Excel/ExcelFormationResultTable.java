package Excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.*;

public class ExcelFormationResultTable {
    public static void formationResultTable(Sheet sheet, String regex){

        // Создаем map для хранения строк с одинаковым критерием
        Map<String, List<Row>> rowsByCrit = new LinkedHashMap<>();

        // Определяем индексы столбцов
        List<Integer> colCrit = new ArrayList<>(); // Индексы столбцов "CRIT"
        List<Integer> colSum = new ArrayList<>(); // Индексы столбцов "SUM"
        List<Integer> colMax = new ArrayList<>(); // Индексы столбцов "MAX"
        List<Integer> colMin = new ArrayList<>(); // Индексы столбцов "MIN"
        List<Integer> colConc = new ArrayList<>(); // Индексы столбцов "CONC"
        Set<Integer> colValues = new HashSet<>(); // Индексы столбцов значений "-"

        // Заполняем описанные выше листы соответствующими индексами
        assert sheet != null;
        Row firstRow = sheet.getRow(0);
        for (int i = 0; i < firstRow.getLastCellNum(); i++) {
            Cell cell = firstRow.getCell(i);
            if (cell != null) {
                switch (cell.getStringCellValue().toLowerCase().replaceAll(regex + "-", "")){
                    case "crit":
                        colCrit.add(cell.getColumnIndex());
                        break;
                    case "sum":
                        colSum.add(cell.getColumnIndex());
                        break;
                    case "max":
                        colMax.add(cell.getColumnIndex());
                        break;
                    case "min":
                        colMin.add(cell.getColumnIndex());
                        break;
                    case "conc":
                        colConc.add(cell.getColumnIndex());
                        break;
                    case "-":
                        colValues.add(cell.getColumnIndex());
                        break;
                }
            }else {
                break;
            }
        }

        // Проходим по всем заполненным строкам
        int numRows = sheet.getLastRowNum();
        for (int i = 1; i <= numRows; i++) {
            Row row = sheet.getRow(i); // Получаем i-ую строку в excel файлу
            if (row != null) {
                // Получаем значения критериев и записываем в лист
                List<String> critValues = new ArrayList<>();
                for (int colIndex : colCrit) {
                    Cell critCell = row.getCell(colIndex);
                    String critValue = "";
                    if(critCell.getCellType() == CellType.STRING){
                        critValue = critCell.getStringCellValue().replaceAll(regex, "");
                    }else {
                        critValue = String.valueOf(critCell.getNumericCellValue());
                    }
                    critValues.add(critValue);
                }

                // Формируем ключ для map
                String critKey = String.join(",", critValues);

                // Если строка с такими критериями уже есть в map, то соединяем
                if (rowsByCrit.containsKey(critKey)) {
                    List<Row> rows = rowsByCrit.get(critKey);
                    Row lastRow = rows.get(rows.size() - 1);

                    // Складываем значения столбца "SUM"
                    for (int colIndex : colSum) {
                        Cell lastCell = lastRow.getCell(colIndex);
                        Cell currCell = row.getCell(colIndex);
                        double lastValue = lastCell.getNumericCellValue();
                        double currValue = currCell.getNumericCellValue();
                        lastCell.setCellValue(lastValue + currValue);
                    }

                    // Находим максимальное значение в столбце "Max"
                    for (int colIndex : colMax) {
                        Cell lastCell = lastRow.getCell(colIndex);
                        Cell currCell = row.getCell(colIndex);
                        double lastValue = lastCell.getNumericCellValue();
                        double currValue = currCell.getNumericCellValue();
                        double maxValue = Math.max(lastValue, currValue);
                        lastCell.setCellValue(maxValue);
                    }

                    // Находим минимальное значение в столбце "Min"
                    for (int colIndex : colMin) {
                        Cell lastCell = lastRow.getCell(colIndex);
                        Cell currCell = row.getCell(colIndex);
                        double lastValue = lastCell.getNumericCellValue();
                        double currValue = currCell.getNumericCellValue();
                        double minValue = Math.min(lastValue, currValue);
                        lastCell.setCellValue(minValue);
                    }

                    // Склеиваем строки в столбце "Conc"
                    for (int colIndex : colConc) {
                        Cell lastCell = lastRow.getCell(colIndex);
                        Cell currCell = row.getCell(colIndex);
                        String lastValue = lastCell.getStringCellValue().replaceAll(regex, "");
                        String currValue = currCell.getStringCellValue().replaceAll(regex, "");
                        lastCell.setCellValue(lastValue + currValue);
                    }

                } else {
                    // Если строка с такими критериями еще не была, то добавляем ее в map
                    List<Row> rows = new ArrayList<>();
                    rows.add(row);
                    rowsByCrit.put(critKey, rows);
                }
            }
        }

        // Инициализация результирующей таблицы
        ArrayList<ArrayList<String>> resultTable = new ArrayList<>();

        // Формирование результирующей таблицы
        for (Map.Entry<String, List<Row>> entry : rowsByCrit.entrySet()) {
            List<Row> rows = entry.getValue();
            for (int i = 0; i < rows.size(); i++) {
                Row row = rows.get(i);
                ArrayList<String> tempArray = new ArrayList<>();
                for(int j = 0; j < row.getLastCellNum(); j++) {
                    if(!colValues.contains(j)) {
                        Cell cell = row.getCell(j);
                        if (cell.getCellType() == CellType.STRING) {
                            tempArray.add(cell.getStringCellValue().replaceAll(regex, ""));
                        } else {
                            tempArray.add(String.valueOf(cell.getNumericCellValue()));
                        }
                    }
                }
                resultTable.add(tempArray);
            }
        }

        ExcelGrouping.setResultTable(resultTable);
    }
}
