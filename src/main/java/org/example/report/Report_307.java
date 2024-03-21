package org.example.report;

import com.spire.xls.*;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.DateTimeGroupingType;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;

import java.io.FileNotFoundException;
import java.time.LocalDate;

public class Report_307 {

    public void createReport_307() throws FileNotFoundException {
        FilterController filterController = new FilterController();

        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(filterController.choosenFile());
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastRow, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 307 ОШИБКИ:
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        //Фильтр колонки "Дата прихода"
        LocalDate currenDateMinus = filterController.dateCurrent().minusDays(1);
        filters.addDateFilter(12, DateTimeGroupingType.Day, currenDateMinus.getYear(), currenDateMinus.getMonthValue(), currenDateMinus.getDayOfMonth(), 0, 0, 0);
        //Фильтр колонки "Зона"
        filters.addFilter(10, "Зона возвратов");
        //Фильтр колонки "Транзит"
        filters.addFilter(17, "Транзитный пункт");
        //Фильтр колонки "Поток"
        filters.addFilter(16, "Прямой");
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Отправление");
        filters.filter();

        wb.saveToFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\307.xlsx");
    }
}

