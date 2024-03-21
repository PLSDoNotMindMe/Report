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
        currentDate = myDatePicker.getValue();
    }

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
    void Error503(MouseEvent event) throws FileNotFoundException {
        Report_503 object = new Report_503();
        object.createReport_503();
    }

    @FXML
    void Error308(MouseEvent event) throws FileNotFoundException {
        checkPivotTable();
        Report_308 object = new Report_308();
        object.createReport_308();
    }

    @FXML
    void Error501(MouseEvent event) throws FileNotFoundException {
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
    void Error201615(MouseEvent event) throws FileNotFoundException {
        checkPivotTable();
        Report_201615 object = new Report_201615();
        object.createReport_201615();
    }

    @FXML
    void Error106(MouseEvent event) throws FileNotFoundException {
        Report_106 object = new Report_106();
        object.createReport_106();
    }

    @FXML
    void Error307(MouseEvent event) throws FileNotFoundException {
        Report_307 object = new Report_307();
        object.createReport_307();
    }

    @FXML
    void Error627(MouseEvent event) throws FileNotFoundException {
        checkPivotTable();
        Report_627 object = new Report_627();
        object.createReport_627();
    }

    @FXML
    void Transit(MouseEvent event) throws FileNotFoundException {
        Report_Transit object = new Report_Transit();
        object.createReport_Transit();
    }

}



