package org.example.report;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.DateTimeGroupingType;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;
import javafx.fxml.FXML;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.input.MouseEvent;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.Scanner;

public class FilterController {

    String nameFile;
    String user;
    FileChooser fileChooser = new FileChooser();

    @FXML
    private TextField NameIn;

    @FXML
    void name(MouseEvent event) {
        nameFile = NameIn.getText();
        System.out.println(nameFile);
    }

     @FXML
    void createFolder(MouseEvent event) {
        user = System.getProperty("user.name");

        File f = new File("C:\\Users\\" + user + "\\Desktop\\Ошибки");
        try {
            if (f.mkdir()) {
                System.out.println("Directory Created");
            } else {
                System.out.println("Directory is not created");
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @FXML
    void Error308(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile("C:\\Users\\SerPivas\\Downloads\\" + nameFile + ".xlsx");
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 20762, 34));

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
        filters.addDateFilter(12, DateTimeGroupingType.Day, current_date.getYear(), current_date.getMonthValue(), current_date.getDayOfMonth(), 0, 0, 0);
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        filters.filter();
        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\308.xlsx");
    }

    @FXML
    void Error501(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile("C:\\Users\\SerPivas\\Downloads\\" + nameFile + ".xlsx");
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 20762, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 501 ОШИБКИ:
        //Добавить столбец для ВПР
        sheet.insertColumn(34);
        sheet.get(1,34).setValue("ВПР");
        sheet.get(1,33).get(String.format("AH1")).setStyle(sheet.get(1,34).get(String.format("AG1")).getStyle());
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Отправление");
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "empty");
        //Фильтр колонки "Поток"
        filters.addFilter(16, "Прямой");
        //Фильтр колонки "Транзит"
        filters.addFilter(17, "Транзитный пункт");
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        //Фильтр колонки "В перевозке"
        filters.customFilter(23, FilterOperatorType.NotEqual,"СПБ_ТСЦ_Шушары");
        //Фильтр колонки "Дата прихода"
        java.time.LocalDate current_date = java.time.LocalDate.now().minusDays(1);
        filters.addDateFilter(12, DateTimeGroupingType.Day, current_date.getYear(), current_date.getMonthValue(), current_date.getDayOfMonth(), 0, 0, 0);
        filters.filter();
        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\501.xlsx");
    }

    @FXML
    void Error304(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile("C:\\Users\\SerPivas\\Downloads\\" + nameFile + ".xlsx");
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 20762, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 304 ОШИБКИ:
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "empty");
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Отправление");
        filters.addFilter(3, "Тарный ящик");
        //Фильтр колонки "Цена"
        filters.customFilter(9, FilterOperatorType.NotEqual," ");
        //Фильтр колонки "Дата прихода"
        java.time.LocalDate current_date = java.time.LocalDate.now().minusDays(1);
        filters.addDateFilter(12, DateTimeGroupingType.Day, current_date.getYear(), current_date.getMonthValue(), current_date.getDayOfMonth(), 0, 0, 0);
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        //Фильтр колонки "Поток"
        filters.addFilter(16, "Прямой");
        //Фильтр колонки "Транзит"
        filters.addFilter(17, "Транзитный пункт");
        //Фильтр колонки "Зона"
        filters.customFilter(10, FilterOperatorType.NotEqual,"Зона контроля", true,FilterOperatorType.NotEqual,"Зона возвратов");
        filters.filter();
        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\304.xlsx");
    }

    @FXML
    void Error201615(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile("C:\\Users\\SerPivas\\Downloads\\" + nameFile + ".xlsx");
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 20762, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 201/615 ОШИБКИ:
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        filters.addFilter(2, "Прибыл в место назначения");
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "empty");
        //Фильтр колонки "Текущее место"
        filters.customFilter(11,FilterOperatorType.Equal,"Зона контроля-Зона контроля-Found-04MU/Зона контроля-Found-04KU",false, FilterOperatorType.Equal, "Зона контроля-Found");
        //Фильтр колонки "Цена"
        filters.customFilter(9, FilterOperatorType.NotEqual," ");
        //Фильтр колонки "Дата прихода"
        java.time.LocalDate current_date = java.time.LocalDate.now();
        java.time.LocalDate current_date1 = java.time.LocalDate.now().minusDays(1);
        filters.customFilter(12,FilterOperatorType.NotEqual,current_date, true, FilterOperatorType.NotEqual,current_date1);
        filters.filter();
        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\201,615.xlsx");
    }

    @FXML
    void Error106(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile("C:\\Users\\SerPivas\\Downloads\\" + nameFile + ".xlsx");
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 20762, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 106 ОШИБКИ:
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "empty");
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        //Фильтр колонки "Зона"
        filters.addFilter(10,"Зона контроля");
        //Фильтр колонки "Текущее место"
        filters.addFilter(11,"Зона контроля/Зона контроля-Expired SLA");
        filters.filter();
        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\106.xlsx");
    }

    @FXML
    void Error307(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile("C:\\Users\\SerPivas\\Downloads\\" + nameFile + ".xlsx");
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 20762, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 307 ОШИБКИ:
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "empty");
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        //Фильтр колонки "Дата прихода"
        java.time.LocalDate current_date = java.time.LocalDate.now().minusDays(1);
        filters.addDateFilter(12, DateTimeGroupingType.Day, current_date.getYear(), current_date.getMonthValue(), current_date.getDayOfMonth(), 0, 0, 0);
        //Фильтр колонки "Зона"
        filters.addFilter(10,"Зона возвратов");
        //Фильтр колонки "Транзит"
        filters.addFilter(17, "Транзитный пункт");
        //Фильтр колонки "Поток"
        filters.addFilter(16, "Прямой");
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Отправление");
        filters.filter();
        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\307.xlsx");
    }

    @FXML
    void Error601(MouseEvent event) {

    }

}

