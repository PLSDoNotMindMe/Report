package org.example.report;

import com.spire.xls.*;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.DateTimeGroupingType;

import java.io.FileNotFoundException;
import java.time.LocalDate;

public class Report_308 {

    public void createReport_308() throws FileNotFoundException {
        FilterController filterController = new FilterController();

        Workbook wb = new Workbook();
        wb.loadFromFile(filterController.fileChoose);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastRow, 34));
        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 308 ОШИБКИ:
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Груз");
        filters.addFilter(3, "RollCage");
        filters.addFilter(3, "Мешок");
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
        //Фильтр колонки "Зона"
        filters.addFilter(10, "Зона контроля");
        filters.addFilter(10, "Зона приемки");
        filters.addFilter(10, "Шут");
        //Фильтр колонки "Дата прихода"
        CellRange range = sheet.getCellRange("M1:M" + lastRow);
        range.setNumberFormat("dd.mm.yyyy");
        if (filterController.currentDate == null) {
            filterController.currentDate = LocalDate.now();
        }
        LocalDate currenDateMinus = filterController.currentDate.minusDays(1);
        filters.addDateFilter(12, DateTimeGroupingType.Day, currenDateMinus.getYear(), currenDateMinus.getMonthValue(), currenDateMinus.getDayOfMonth(), 0, 0, 0);
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        filters.filter();

        if (filterController.checkChek == true) {
            filterController.fileCheck();
            //Копирование видимых ячеек
            Worksheet sheet3 = wb.getWorksheets().add("308");
            int index = 0;
            for (int i = 1; i <= sheet.getRows().length; i++) {
                if (sheet.getRowIsHide(i)) {
                    continue;
                } else {
                    sheet3.insertRow(index + 1);
                    sheet.copy(sheet.getRows()[i - 1], sheet3.getRows()[index], true, true, true);
                    index++;
                }
                System.out.println(i);
            }

            //Копирование листа в другой файл
            Workbook wb2 = new Workbook();
            wb2.loadFromFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + filterController.formatDate.format(filterController.currentDate) + ".xlsx");
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Некорректное размещение груза");
            Worksheet sheetwork = wb2.getWorksheets().add("308");
            sheetwork.copyFrom(sheet3);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH" + lastRow);
            ;
            PivotCache cache = wb2.getPivotCaches().add(dataRange);
            PivotTable pt = sheetOfWorkbook1.getPivotTables().add("Количество по полю ID предмета", sheetOfWorkbook1.getCellRange("A3"), cache);
            PivotField pf = null;
            if (pt.getPivotFields().get("Зона") instanceof PivotField) {
                pf = (PivotField) pt.getPivotFields().get("Зона");
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
            wb.saveToFile("C:\\Users\\" + filterController.user + "\\Desktop\\Ошибки\\308.xlsx");
        }
    }
}

