package org.example.report;

import java.io.FileNotFoundException;


public class Report_503 {

    public void createReport_503() throws FileNotFoundException {

        FilterController filterController = new FilterController();
        //Выбор файла, создание документа
//        String fileChoose = String.valueOf(file);
//        Workbook wb = new Workbook();
//        wb.loadFromFile(fileChoose, ",", 1, 1);
//        Worksheet sheet = wb.getWorksheets().get(0);
//        int lastRow = sheet.getLastRow();
//        sheet.getCellRange("A1:V" + lastRow).setIgnoreErrorOptions(EnumSet.of(IgnoreErrorType.NumberAsText));
//        CellRange range = sheet.getCellRange("M1:M" + lastRow);
//        range.setNumberFormat("dd.mm.yyyy");
//        //Перенос текста по столбцам и применение автофильтра
//        AutoFiltersCollection filters = sheet.getAutoFilters();
//        filters.setRange(sheet.getCellRange(1, 1, lastRow, 22));
//        //Фильтр колонки "Статус"
//        filters.addFilter(1, "Сформирован");
//        //Фильтр колонки "Завершили формирование"
//        if (currentDate == null) {
//            currentDate = LocalDate.now();
//        }
//        filters.customFilter(2, FilterOperatorType.NotEqual, formatDate.format(currentDate), true, FilterOperatorType.NotEqual, "");
//        filters.filter();
//
//        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\503.xlsx");
    }
}

