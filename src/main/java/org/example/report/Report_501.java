package org.example.report;

import com.spire.xls.*;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.DateTimeGroupingType;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;

import java.io.FileNotFoundException;
import java.time.LocalDate;

public class Report_501 {

    public void createReport_501() throws FileNotFoundException {
        FilterController filterController = new FilterController();

        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(filterController.choosenFile());
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastRow, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 501 ОШИБКИ:
        //Добавить столбец для ВПР
        sheet.insertColumn(34);
        sheet.get(1, 34).setValue("ВПР");
        sheet.get(1, 33).get("AH1").setStyle(sheet.get(1, 34).get("AG1").getStyle());
        sheet.get(1, 34).autoFitColumns();
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Отправление");
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
        //Фильтр колонки "Поток"
        filters.addFilter(16, "Прямой");
        //Фильтр колонки "Транзит"
        filters.addFilter(17, "Транзитный пункт");
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        //Фильтр колонки "Сортировочный центр"
        filters.customFilter(23, FilterOperatorType.NotEqual, "СПБ_ТСЦ_Шушары");
        //Фильтр колонки "Дата прихода"
        LocalDate currenDateMinus = filterController.dateCurrent().minusDays(1);
        filters.addDateFilter(12, DateTimeGroupingType.Day, currenDateMinus.getYear(), currenDateMinus.getMonthValue(), currenDateMinus.getDayOfMonth(), 0, 0, 0);
        filters.filter();

        if (filterController.checkPivot()) {
            filterController.fileCheck();
            //Копирование видимых ячеек
            Worksheet sheet4 = wb.getWorksheets().add("501");

            int index = 0;
            for (int i = 1; i <= sheet.getRows().length; i++) {
                if (sheet.getRowIsHide(i)) {
                    continue;
                } else {
                    sheet4.insertRow(index + 1);
                    sheet.copy(sheet.getRows()[i - 1], sheet4.getRows()[index], true, true, true);
                    index++;
                }
                System.out.println(i);
            }
            //Копирование листа в другой файл
            Workbook wb2 = new Workbook();
            wb2.loadFromFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + filterController.formatDate.format(FilterController.currentDate) + ".xlsx");
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Не отправлен из магистрали");
            Worksheet sheetwork = wb2.getWorksheets().add("501");
            Worksheet sheet2 = wb2.getWorksheets().add("Задержка отправки груза");
            Worksheet sheet3 = wb2.getWorksheets().add("Задержка отправки Xdoc");
            sheetwork.copyFrom(sheet4);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH" + lastRow);
            PivotCache cache = wb2.getPivotCaches().add(dataRange);
            PivotTable pt = sheetOfWorkbook1.getPivotTables().add("Количество по полю ID предмета", sheetOfWorkbook1.getCellRange("A3"), cache);
            PivotField pf = null;
            if (pt.getPivotFields().get("Текущее место") instanceof PivotField) {
                pf = (PivotField) pt.getPivotFields().get("Текущее место");
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
            wb.saveToFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\502.xlsx");
        } else {
            wb.saveToFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\501.xlsx");

        }
    }
}

