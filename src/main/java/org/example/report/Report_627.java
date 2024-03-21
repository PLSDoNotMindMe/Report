package org.example.report;

import com.spire.xls.*;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;

import java.io.FileNotFoundException;

public class Report_627 {

    public void createReport_627() throws FileNotFoundException {
        FilterController filterController = new FilterController();

        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(filterController.choosenFile());
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastRow, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 627 ОШИБКИ:
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Коробка");
        filters.addFilter(3, "Мешок");
        filters.addFilter(3, "Сейф пакет");
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        filters.addFilter(2, "Прибыл в место назначения");
        //Фильтр колонки "Цена"
        filters.customFilter(9, FilterOperatorType.NotEqual, " ");
        //Фильтр колонки "Дата прихода"
        CellRange range = sheet.getCellRange("M1:M" + lastRow);
        range.setNumberFormat("dd.MM.yyyy");
        filters.customFilter(12, FilterOperatorType.NotEqual, filterController.dateCurrent());
        filters.filter();

        if (filterController.checkPivot()) {
            filterController.fileCheck();

            //Копирование видимых ячеек
            Worksheet sheet1 = wb.getWorksheets().add("627");

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
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Необработанные ТМ");
            Worksheet sheetwork = wb2.getWorksheets().add("627");
            Worksheet sheetwork1 = wb2.getWorksheets().add("Проверенные груза");

            sheetwork.copyFrom(sheet1);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH" + lastRow);
            PivotCache cache = wb2.getPivotCaches().add(dataRange);
            PivotTable pt = sheetOfWorkbook1.getPivotTables().add("Количество по полю ID предмета", sheetOfWorkbook1.getCellRange("A3"), cache);
            PivotField pf = null;
            if (pt.getPivotFields().get("Тип") instanceof PivotField) {
                pf = (PivotField) pt.getPivotFields().get("Тип");
            }
            PivotField pf1 = null;
            if (pt.getPivotFields().get("Текущее место") instanceof PivotField) {
                pf1 = (PivotField) pt.getPivotFields().get("Текущее место");
            }
            pf.setAxis(AxisTypes.Row);
            pf1.setAxis(AxisTypes.Row);
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
            wb.saveToFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\627.xlsx");
        }

    }
}

