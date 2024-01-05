package org.example.report;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.DateTimeGroupingType;
import com.spire.xls.*;

import java.io.File;

public class Main {
    public static void main(String[] args) {

        File f = new File("C:\\Users\\SerPivas\\Desktop\\Ошибки");
        try{
            if(f.mkdir()) {
                System.out.println("Directory Created");
            } else {
                System.out.println("Directory is not created");
            }
        } catch(Exception e){
            e.printStackTrace();
        }


        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile("C:\\Users\\SerPivas\\Downloads\\stock_23641575982000.xlsx");
        Worksheet sheet = wb.getWorksheets().get(0);
        sheet.setName("Сток");
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1,1,20762,34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 308 ОШИБКИ:
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Груз");
        filters.addFilter(3, "RollCage");
        filters.addFilter(3, "Мешок");
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "empty");
        //Фильтр колонки "Зона"
        filters.addFilter(10, "Зона контроля");
        filters.addFilter(10, "Зона приемки");
        filters.addFilter(10, "Шут");
        //Фильтр колонки "Дата прихода"
        java.time.LocalDate current_date = java.time.LocalDate.now().minusDays(1);
        filters.addDateFilter(12, DateTimeGroupingType.Day, current_date.getYear(), current_date.getMonthValue(), current_date.getDayOfMonth(), 0,0,0);
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        filters.filter();
        wb.saveToFile("C:\\Users\\SerPivas\\Desktop\\Ошибки\\308.xlsx");

    }
}