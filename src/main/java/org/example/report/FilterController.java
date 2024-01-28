package org.example.report;

import com.almasb.fxgl.audio.Audio;
import com.almasb.fxgl.audio.AudioPlayer;
import com.spire.pdf.htmlconverter.qt.Clip;
import com.spire.xls.*;
import com.spire.xls.collections.AutoFiltersCollection;
import com.spire.xls.collections.PivotTablesCollection;
import com.spire.xls.core.spreadsheet.autofilter.DateTimeGroupingType;
import com.spire.xls.core.spreadsheet.autofilter.FilterOperatorType;
import com.spire.xls.core.spreadsheet.pivottables.XlsPivotField;
import javafx.event.ActionEvent;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.input.MouseEvent;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.jetbrains.annotations.NotNull;

import javax.sound.sampled.AudioInputStream;
import javax.sound.sampled.AudioSystem;
import javax.sound.sampled.UnsupportedAudioFileException;
import java.io.*;
import java.net.URL;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Date;
import java.util.EnumSet;
import java.util.List;
import java.util.ResourceBundle;
import java.util.stream.Collectors;
import java.util.stream.Stream;





public class FilterController implements Initializable {


    FileChooser fileChooser = new FileChooser();
    String user;
    String fileCh;
    String fileerror;
    String filename;
    java.time.LocalDate current_date = java.time.LocalDate.now().minusDays(1);
    java.time.LocalDate newdate = LocalDate.now();
    DateTimeFormatter formatdate = DateTimeFormatter.ofPattern("dd.MM.yyyy");
    String audeioPath = "C:\\Windows\\Media\\Windows Message Nudge.wav";


    @FXML
    private CheckBox Check;


    @FXML
    private Label nameout;

    public FilterController() throws FileNotFoundException {
    }


    @FXML
    void name(MouseEvent event) {

    }

    @FXML
    private Button newfile;

    @FXML
    private Label ErrorChoose;

    @FXML
    private Separator seperator1;

    @FXML
    void CheckPt(ActionEvent event) {


    }

    // Выбор файла
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

    //Создание ексель файла
    @FXML
    void createFile(MouseEvent event) {
        Workbook wb = new Workbook();
        wb.getWorksheets().clear();
        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatdate.format(newdate) + ".xlsx");
    }

    //Создание папки
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
    void Error503(MouseEvent event) {
        //Выбор файла, создание документа
        File file = fileChooser.showOpenDialog(new Stage());
        file.getAbsoluteFile();
        fileerror = String.valueOf(file);
        Workbook wb = new Workbook();
        wb.loadFromFile(fileerror, ",");
        Worksheet sheet = wb.getWorksheets().get(0);
        CellRange usedRange = sheet.getAllocatedRange();
        usedRange.setIgnoreErrorOptions(EnumSet.of(IgnoreErrorType.NumberAsText));

        CellRange range = sheet.getCellRange("M1:M40000");
        range.setNumberFormat("dd.mm.yyyy");
        //Перенос текста по столбцам и применение автофильтра
        AutoFiltersCollection filters = sheet.getAutoFilters();
        filters.setRange(sheet.getCellRange(1, 1, 40000, 22));
        //Фильтр колонки "Статус"
        filters.addFilter(6,"Сформирован");
        //Фильтр колонки "Завершили формирование"
        java.time.LocalDate current_date1 = java.time.LocalDate.now();
        filters.customFilter(12,FilterOperatorType.NotEqual,current_date1,true,FilterOperatorType.NotEqual,"");
        filters.filter();

        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\503.xlsx");

    }


    @FXML
    void Error308(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileCh);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastr = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastr, 34));


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
        CellRange range = sheet.getCellRange("M1:M"+lastr);;
        range.setNumberFormat("dd.mm.yyyy");
        filters.addDateFilter(12, DateTimeGroupingType.Day, current_date.getYear(), current_date.getMonthValue(), current_date.getDayOfMonth(), 0, 0, 0);
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        filters.filter();

        if (Check.isSelected()) {

            //Копирование видимых ячеек
            Worksheet sheet3 = wb.getWorksheets().add("308");
            Worksheet sheet4 = wb.getWorksheets().add("Ненормативные возвраты");

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
            wb2.loadFromFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatdate.format(newdate) + ".xlsx");
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Некорректное размещение груза");
            Worksheet sheetwork = wb2.getWorksheets().add("308");
            sheetwork.copyFrom(sheet3);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH"+lastr);;
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
            wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\308.xlsx");
        }
    }


    @FXML
    void Error501(MouseEvent event) throws UnsupportedAudioFileException, IOException {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileCh);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastr = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastr, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 501 ОШИБКИ:
        //Добавить столбец для ВПР
        sheet.insertColumn(34);
        sheet.get(1,34).setValue("ВПР");
        sheet.get(1,33).get(String.format("AH1")).setStyle(sheet.get(1,34).get(String.format("AG1")).getStyle());
        sheet.get(1,34).autoFitColumns();
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
        filters.customFilter(23, FilterOperatorType.NotEqual,"СПБ_ТСЦ_Шушары");
        //Фильтр колонки "Дата прихода"
        filters.addDateFilter(12, DateTimeGroupingType.Day, current_date.getYear(), current_date.getMonthValue(), current_date.getDayOfMonth(), 0, 0, 0);
        filters.filter();

        if (Check.isSelected()) {

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
            wb2.loadFromFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatdate.format(newdate) + ".xlsx");
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Не отправлен из магистрали");
            Worksheet sheetwork = wb2.getWorksheets().add("501");
            Worksheet sheet2 = wb2.getWorksheets().add("Задержка отправки груза");
            Worksheet sheet3 = wb2.getWorksheets().add("Задержка отправки Xdoc");
            sheetwork.copyFrom(sheet4);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH"+lastr);;
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
            wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\502.xlsx");
        } else {
            wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\501.xlsx");

        }

    }

    @FXML
    void Error304(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileCh);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastr = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastr, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 304 ОШИБКИ:
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        //Фильтр колонки "Тип"
        filters.addFilter(3, "Отправление");
        filters.addFilter(3, "Тарный ящик");
        //Фильтр колонки "Цена"
        filters.customFilter(9, FilterOperatorType.NotEqual," ");
        //Фильтр колонки "Дата прихода"
        CellRange range = sheet.getCellRange("M1:M"+lastr);;
        range.setNumberFormat("dd.mm.yyyy");
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

        if (Check.isSelected()) {

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
            wb2.loadFromFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatdate.format(newdate) + ".xlsx");
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Не отправлен из магистрали");
            Worksheet sheetwork = wb2.getWorksheets().add("304");
            sheetwork.copyFrom(sheet1);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH"+lastr);;
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
            wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\304.xlsx");
        }

        wb.saveToFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\304.xlsx");
    }



    @FXML
    void Error201615(MouseEvent event) {
        //Создание документа, установка автофильтра
        Workbook wb = new Workbook();
        wb.loadFromFile(fileCh);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastr = sheet.getLastRow();
        System.out.println(lastr);
        filters.setRange(sheet.getCellRange(1, 1, lastr, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 201/615 ОШИБКИ:
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        filters.addFilter(2, "Прибыл в место назначения");
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
        //Фильтр колонки "Текущее место"
        filters.customFilter(11,FilterOperatorType.Equal,"Зона контроля-Зона контроля-Found-04MU/Зона контроля-Found-04KU",false, FilterOperatorType.Equal, "Зона контроля-Found");
        //Фильтр колонки "Цена"
        filters.customFilter(9, FilterOperatorType.NotEqual," ");
        //Фильтр колонки "Дата прихода"
        CellRange range = sheet.getCellRange("M1:M"+lastr);
        range.setNumberFormat("dd.MM.yyyy");
        java.time.LocalDate current_date1 = java.time.LocalDate.now();
        filters.customFilter(12,FilterOperatorType.NotEqual,current_date1,true,FilterOperatorType.NotEqual,current_date);
        filters.filter();

        if (Check.isSelected()) {

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
            wb2.loadFromFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatdate.format(newdate) + ".xlsx");
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Found");
            Worksheet sheetwork = wb2.getWorksheets().add("201,615");
            sheetwork.copyFrom(sheet1);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH"+lastr);
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
        wb.loadFromFile(fileCh);
        Worksheet sheet = wb.getWorksheets().get(0);
        AutoFiltersCollection filters = sheet.getAutoFilters();
        int lastr = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastr, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 106 ОШИБКИ:
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
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
        int lastr = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastr, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 307 ОШИБКИ:
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
        //Фильтр колонки "В перевозке"
        filters.addFilter(27, "Нет");
        //Фильтр колонки "Дата прихода"
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
        int lastr = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastr, 34));

        //ПРИМЕНЕНИЕ ФИЛЬТРОВ 601 ОШИБКИ:
        //Фильтр колонки "Статус"
        filters.addFilter(2, "Сформирован");
        filters.addFilter(2, "Прибыл в место назначения");
        //Фильтр колонки "Контейнер (груз)"
        filters.addFilter(4, "");
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
        int lastr = sheet.getLastRow();
        filters.setRange(sheet.getCellRange(1, 1, lastr, 34));

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
        CellRange range = sheet.getCellRange("M1:M"+lastr);
        range.setNumberFormat("dd.MM.yyyy");
        java.time.LocalDate current_date1 = java.time.LocalDate.now();
        filters.customFilter(12,FilterOperatorType.NotEqual,current_date1);
        filters.filter();

        if (Check.isSelected()) {

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
            wb2.loadFromFile("C:\\Users\\" + user + "\\Desktop\\Ошибки\\Ежедневный отчёт по ошибкам СПБ_ТСЦ_Шушары " + formatdate.format(newdate) + ".xlsx");
            Worksheet sheetOfWorkbook1 = wb2.getWorksheets().add("Необработанные ТМ");
            Worksheet sheetwork = wb2.getWorksheets().add("627");
            Worksheet sheetwork1 = wb2.getWorksheets().add("Проверенные груза");

            sheetwork.copyFrom(sheet1);

            //Сводная таблица
            CellRange dataRange = sheetwork.getCellRange("A1:AH"+lastr);;
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

    }


}



