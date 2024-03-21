package org.example.report;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;

import java.io.FileNotFoundException;

public class Report_Transit {

    public void createReport_Transit() throws FileNotFoundException {
        FilterController filterController = new FilterController();

        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(filterController.choosenFile());
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastRow, 34));
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Транзитная коробка");
        //Фильтр колонки "Наименование"
        filters.customFilter(1, FilterOperatorType.Equal, "sr*");
        filters.filter();

        wb.saveToFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\Транзитные коробки.xlsx");

    }
}


