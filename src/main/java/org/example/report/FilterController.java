package org.example.report;

import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.DateTimeGroupingType;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.File;
import java.io.FileNotFoundException;
import java.net.URL;
import java.time.LocalDate;
import java.util.ResourceBundle;
import java.util.Scanner;





public class FilterController implements Initializable {

    FileChooser fileChooser = new FileChooser();
    String user;
    String fileCh;
    String filename;
    String dateCh;

    @FXML
    void chooseFile(MouseEvent event) {

        File file = fileChooser.showOpenDialog(new Stage());
        file.getAbsoluteFile();
        fileCh = String.valueOf(file);
        filename = file.getName();
        nameout.setText(filename);
    }

    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {
        user = System.getProperty("user.name");
        fileChooser.setInitialDirectory(new File("C:\\Users\\" + user + "\\Desktop"));
    }

    @FXML
    private Label nameout;


    @FXML
    void name(MouseEvent event) {

    }

    @FXML
    private Button newfile;


    @FXML
    void createFile(MouseEvent event) {

    }

    @FXML
    private Label ErrorChoose;

    @FXML
    private Separator seperator1;


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
        wb.loadFromFile(fileCh);
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
        wb.loadFromFile(fileCh);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 20762, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 501 ОШИБКИ:
        //Добавить столбец для ВПР
        sheet.insertColumn(34);
        sheet.get(1,34).setValue("ВПР");
        sheet.get(1,33).get(String.format("AH1")).setStyle(sheet.get(1,34).get(String.format("AG1")).getStyle());
        sheet.get(1,34).autoFitColumns();
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
        //Фильтр колонки "Сортировочный центр"
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
        wb.loadFromFile(fileCh);
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
        wb.loadFromFile(fileCh);
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
        wb.loadFromFile(fileCh);
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
        wb.loadFromFile(fileCh);
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
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileCh);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 20762, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 601 ОШИБКИ:
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        filters.addFilter(2, "Прибыл в место назначения");
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "empty");
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Отправление");
        filters.addFilter(3, "Экземпляр товара");
        //Фильтр колонки "Текущее место"
        filters.removeFilter(11,"Компенсированные");
        filters.removeFilter(11,"Протечка");
        filters.removeFilter(11,"Просроченные");
        //Фильтр колонки "Поток"
        filters.addFilter(16, "Возвратный");
        //Добавить столбец для ВПР
        sheet.insertColumn(2);
        sheet.get(1,2).setValue("Детализация");
        sheet.get(1,33).get(String.format("B1")).setStyle(sheet.get(1,34).get(String.format("A1")).getStyle());
        sheet.get(1,2).autoFitColumns();
        filters.filter();
        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\601.xlsx");
    }

    @FXML
    void Error627(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileCh);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 20762, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 627 ОШИБКИ:
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Коробка");
        filters.addFilter(3, "Мешок");
        filters.addFilter(3, "Сейф пакет");
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        filters.addFilter(2, "Прибыл в место назначения");
        //Фильтр колонки "Цена"
        filters.customFilter(9, FilterOperatorType.NotEqual," ");
        //Фильтр колонки "Дата прихода"
        java.time.LocalDate current_date = java.time.LocalDate.now();
        filters.customFilter(12,FilterOperatorType.NotEqual,current_date);
        filters.filter();
        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\627.xlsx");
    }

    @FXML
    void Transit(MouseEvent event) {

    }


}



