package org.example.report;

import com.spire.xls.*;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.core.spreadsheet.autofilter.DateTimeGroupingType;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.*;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.ResourceBundle;


public class FilterController implements Initializable {

    FileChooser fileChooser = new FileChooser();
    static String fileChoose;
    static LocalDate currentDate;
    String user = System.getProperty("user.name");
    static boolean isCheck;
    DateTimeFormatter formatDate = DateTimeFormatter.ofPattern("dd.MM.yyyy");

    public LocalDate dateCurrent() {
      return currentDate;
    }
    public String choosenFile() {
       return fileChoose;
    }
    public boolean checkPivot() {
        return isCheck;
    }
    public void fileCheck() {
        Path path = Path.of("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatDate.format(currentDate) + ".xlsx");
        if (Files.notExists(path)) {
            Workbook wb = new Workbook();
            wb.getWorksheets().clear();
            wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatDate.format(currentDate) + ".xlsx");
        }
    }
    public void checkPivotTable() {
        if (Check.isSelected()) {
            isCheck = true;
        } else {
            isCheck = false;
        }
    }
    @FXML
    public void getDate(ActionEvent event) {
        currentDate = myDatePicker.getValue();}
    @FXML
    public void chooseFile(MouseEvent event) {
        File file = fileChooser.showOpenDialog(new Stage());
        file.getAbsoluteFile();
        fileChoose = String.valueOf(file);
        String fileName = file.getName();
        nameOut.setText(fileName);
    }
    @Override
    public void initialize(URL url, ResourceBundle resourceBundle) {
        fileChooser.setInitialDirectory(new File("C:\\Users\\" + user + "\\Downloads"));
    }
    @FXML
    void CheckPt(ActionEvent event) {

    }

    @FXML
    public CheckBox Check;

    @FXML
    private Label nameOut;

    @FXML
    public DatePicker myDatePicker;

    @FXML
    void Error503(MouseEvent event) {
//        //Выбор файла, создание документа
//        File file = fileChooser.showOpenDialog(new Stage());
//        file.getAbsoluteFile();
//        String fileChoose = String.valueOf(file);
//        Workbook wb = new Workbook();
//        wb.loadFromFile(fileChoose, ",", 1, 1);
//        Worksheet sheet = wb.getWorksheets().get(0);
//        int lastRow = sheet.getLastRow();
//        sheet.getCellRange("A1:V" + lastRow).setIgnoreErrorOptions(EnumSet.of(IgnoreErrorType.NumberAsText));
//        CellRange range = sheet.getCellRange("M1:M" + lastRow);
//        range.setNumberFormat("dd.mm.yyyy");
//        //Перенос текста по столбцам и применение автофильтра
//        AutoFiltersCollection filters = sheet.getAutoFilters();
//        filters.setRange(sheet.getCellRange(1, 1, lastRow, 22));
//        //Фильтр колонки "Статус"
//        filters.addFilter(1, "Сформирован");
//        //Фильтр колонки "Завершили формирование"
//        if (currentDate == null) {
//            currentDate = LocalDate.now();
//        }
//        filters.customFilter(2, FilterOperatorType.NotEqual, formatDate.format(currentDate), true, FilterOperatorType.NotEqual, "");
//        filters.filter();
//
//        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\503.xlsx");

    }

    @FXML
    void Error308(MouseEvent event) throws FileNotFoundException {
        checkPivotTable();
        Report_308 object = new Report_308();
        object.createReport_308();
    }

    @FXML
    void Error501(MouseEvent event) throws FileNotFoundException{
        checkPivotTable();
        Report_501 object = new Report_501();
        object.createReport_501();
    }

    @FXML
    void Error304(MouseEvent event) throws FileNotFoundException {
        checkPivotTable();
        Report_304 object = new Report_304();
        object.createReport_304();
    }

    @FXML
    void Error201615(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileChoose);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        System.out.println(lastRow);
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
        if (currentDate == null) {
            currentDate = LocalDate.now();
        }
        LocalDate currenDateMinus = currentDate.minusDays(1);
        filters.customFilter(12, FilterOperatorType.NotEqual, currentDate, true, FilterOperatorType.NotEqual, currenDateMinus);
        filters.filter();

        if (Check.isSelected()) {
            fileCheck();
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
            wb2.loadFromFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatDate.format(currentDate) + ".xlsx");
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
            wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\201,615.xlsx");
        }

    }

    @FXML
    void Error106(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileChoose);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastRow, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 106 ОШИБКИ:
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        //Фильтр колонки "Зона"
        filters.addFilter(10, "Зона контроля");
        //Фильтр колонки "Текущее место"
        filters.addFilter(11, "Зона контроля/Зона контроля-Expired SLA");
        filters.filter();

        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\106.xlsx");
    }

    @FXML
    void Error307(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileChoose);
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
        if (currentDate == null) {
            currentDate = LocalDate.now();
        }
        LocalDate currenDateMinus = currentDate.minusDays(1);
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

        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\307.xlsx");
    }

    @FXML
    void Error627(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileChoose);
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
        if (currentDate == null) {
            currentDate = LocalDate.now();
        }
        filters.customFilter(12, FilterOperatorType.NotEqual, currentDate);
        filters.filter();

        if (Check.isSelected()) {
            fileCheck();

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
            wb2.loadFromFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatDate.format(currentDate) + ".xlsx");
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Необработанные ТМ");
            Worksheet sheetwork = wb2.getWorksheets().add("627");
            Worksheet sheetwork1 = wb2.getWorksheets().add("Проверенные груза");

            sheetwork.copyFrom(sheet1);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH" + lastRow);
            ;
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
            wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\627.xlsx");
        }

    }

    @FXML
    void Transit(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileChoose);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastRow = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastRow, 34));
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Транзитная коробка");
        //Фильтр колонки "Наименование"
        filters.customFilter(1,FilterOperatorType.Equal,"sr*");
        filters.filter();

        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Транзитные коробки.xlsx");

    }


}



