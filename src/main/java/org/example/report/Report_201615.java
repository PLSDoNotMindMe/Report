package org.example.report;

import com.spire.xls.*;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;

import java.io.FileNotFoundException;
import java.time.LocalDate;

public class Report_201615 {

    public void createReport_201615() throws FileNotFoundException {
        FilterController filterController = new FilterController();

        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(filterController.choosenFile());
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastRow, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 201/615 ОШИБКИ:
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        filters.addFilter(2, "Прибыл в место назначения");
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
        //Фильтр колонки "Текущее место"
        filters.customFilter(11, FilterOperatorType.Equal, "Зона контроля-Зона контроля-Found-04MU/Зона контроля-Found-04KU", false, FilterOperatorType.Equal, "Зона контроля-Found");
        //Фильтр колонки "Цена"
        filters.customFilter(9, FilterOperatorType.NotEqual, " ");
        //Фильтр колонки "Дата прихода"
        CellRange range = sheet.getCellRange("M1:M" + lastRow);
        range.setNumberFormat("dd.MM.yyyy");

        LocalDate currenDateMinus = filterController.dateCurrent().minusDays(1);
        filters.customFilter(12, FilterOperatorType.NotEqual, filterController.dateCurrent(), true, FilterOperatorType.NotEqual, currenDateMinus);
        filters.filter();

        if (filterController.checkPivot()) {
            filterController.fileCheck();
            //Копирование видимых ячеек
            Worksheet sheet1 = wb.getWorksheets().add("201,615");

            int index = 0;
            for (int i = 1; i <= sheet.getRows().length; i++) {
                if (sheet.getRowIsHide(i)) {
                    continue;
                } else {
                    sheet1.insertRow(index + 1);
                    sheet.copy(sheet.getRows()[i - 1], sheet1.getRows()[index], true, true, true);
                    index++;
                }
                System.out.println(i);
            }
            //Копирование листа в другой файл
            Workbook wb2 = new Workbook();
            wb2.loadFromFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + filterController.formatDate.format(FilterController.currentDate) + ".xlsx");
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Found");
            Worksheet sheetwork = wb2.getWorksheets().add("201,615");
            sheetwork.copyFrom(sheet1);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH" + lastRow);
            PivotCache cache = wb2.getPivotCaches().add(dataRange);
            PivotTable pt = sheetOfWorkbook1.getPivotTables().add("Количество по полю ID предмета", sheetOfWorkbook1.getCellRange("A3"), cache);
            PivotField pf = null;
            if (pt.getPivotFields().get("Текущее место") instanceof PivotField) {
                pf = (PivotField) pt.getPivotFields().get("Текущее место");
            }
            pf.setAxis(AxisTypes.Row);
            PivotField pf1 = null;
            if (pt.getPivotFields().get("Тип") instanceof PivotField) {
                pf1 = (PivotField) pt.getPivotFields().get("Тип");
            }
            pf1.setAxis(AxisTypes.Row);
            pt.getDataFields().add(pt.getPivotFields().get("ID предмета"), "Количество по полю ID предмета", SubtotalTypes.Sum);
            PivotField pf2 = null;
            if (pt.getPivotFields().get("Дата прихода на СЦ") instanceof PivotField) {
                pf2 = (PivotField) pt.getPivotFields().get("Дата прихода на СЦ");
            }
            pf2.setAxis(AxisTypes.Column);
            pt.getOptions().setColumnHeaderCaption("Дата прихода на СЦ");
            PivotField pf3 = null;
            if (pt.getPivotFields().get("Поток") instanceof PivotField) {
                pf3 = (PivotField) pt.getPivotFields().get("Поток");
            }
            pf3.setAxis(AxisTypes.Page);
            pt.setBuiltInStyle(PivotBuiltInStyles.PivotStyleMedium10);

            wb2.save();
        } else {
            wb.saveToFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\201,615.xlsx");
        }

    }
}

