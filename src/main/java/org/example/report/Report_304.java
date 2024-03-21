package org.example.report;

import com.spire.xls.*;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.DateTimeGroupingType;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;

import java.io.FileNotFoundException;
import java.time.LocalDate;

public class Report_304 {

    public void createReport_304() throws FileNotFoundException {
        FilterController filterController = new FilterController();

        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(filterController.choosenFile());
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastRow, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 304 ОШИБКИ:
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Отправление");
        filters.addFilter(3, "Тарный ящик");
        //Фильтр колонки "Цена"
        filters.customFilter(9, FilterOperatorType.NotEqual, " ");
        //Фильтр колонки "Дата прихода"
        CellRange range = sheet.getCellRange("M1:M" + lastRow);
        range.setNumberFormat("dd.mm.yyyy");
        LocalDate currenDateMinus = filterController.dateCurrent().minusDays(1);
        filters.addDateFilter(12, DateTimeGroupingType.Day, currenDateMinus.getYear(), currenDateMinus.getMonthValue(), currenDateMinus.getDayOfMonth(), 0, 0, 0);
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        //Фильтр колонки "Поток"
        filters.addFilter(16, "Прямой");
        //Фильтр колонки "Транзит"
        filters.addFilter(17, "Транзитный пункт");
        //Фильтр колонки "Зона"
        filters.customFilter(10, FilterOperatorType.NotEqual, "Зона контроля", true, FilterOperatorType.NotEqual, "Зона возвратов");
        filters.filter();

        if (filterController.checkPivot()) {
            filterController.fileCheck();

            //Копирование видимых ячеек
            Worksheet sheet1 = wb.getWorksheets().add("304");

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
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Нарушение SLA обработки на ТСЦ");
            Worksheet sheetwork = wb2.getWorksheets().add("304");
            sheetwork.copyFrom(sheet1);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH" + lastRow);
            PivotCache cache = wb2.getPivotCaches().add(dataRange);
            PivotTable pt = sheetOfWorkbook1.getPivotTables().add("Количество по полю ID предмета", sheetOfWorkbook1.getCellRange("A3"), cache);
            PivotField pf = null;
            if (pt.getPivotFields().get("Тип") instanceof PivotField) {
                pf = (PivotField) pt.getPivotFields().get("Тип");
            }
            pf.setAxis(AxisTypes.Row);
            pt.getDataFields().add(pt.getPivotFields().get("ID предмета"), "Количество по полю ID предмета", SubtotalTypes.Sum);
            PivotField pf2 = null;
            if (pt.getPivotFields().get("Дата прихода на СЦ") instanceof PivotField) {
                pf2 = (PivotField) pt.getPivotFields().get("Дата прихода на СЦ");
            }
            pf2.setAxis(AxisTypes.Column);
            pt.getOptions().setColumnHeaderCaption("Дата прихода на СЦ");
            pt.setBuiltInStyle(PivotBuiltInStyles.PivotStyleMedium10);

            wb2.save();
        } else {
            wb.saveToFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\304.xlsx");
        }
        System.out.println(filterController.dateCurrent());
    }
}

