package com.hpcnt.deviceRateAutomation.controller;

import com.hpcnt.deviceRateAutomation.model.Device;
import javafx.scene.control.Alert;
import javafx.scene.control.ButtonType;
import javafx.scene.control.TextInputDialog;
import javafx.scene.layout.GridPane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * Created by owen151128 on 2018. 9. 5.
 */
public class MainController {
    private String iosPrefix1 = "Apple";
    private String iosPrefix2 = "iPhone";
    private String buffer1;
    private final String AND = "AOS_";
    private final String IOS = "IOS_";
    private String[] title;
    private String fileName;
    private boolean isIos;
    public GridPane mainPane;
    private Alert alert;

    public MainController() {
        setAlert();
        isIos = false;
    }

    public void onIosPrefixButtonClicked() {

        TextInputDialog dialog = new TextInputDialog();
        dialog.setContentText("Ios Prefix 1 : ");
        dialog.getEditor().setText(iosPrefix1);
        dialog.getDialogPane().lookupButton(ButtonType.CANCEL).setVisible(false);
        Optional<String> result = dialog.showAndWait();
        if (!result.isPresent())
            return;
        result.ifPresent(s -> buffer1 = s);
        dialog = new TextInputDialog();
        dialog.getDialogPane().lookupButton(ButtonType.CANCEL).setVisible(false);
        dialog.setContentText("Ios Prefix 2 : ");
        dialog.getEditor().setText(iosPrefix2);
        result = dialog.showAndWait();
        if (!result.isPresent())
            return;
        result.ifPresent(s -> {
            iosPrefix1 = buffer1;
            iosPrefix2 = s;
        });
        Alert alert = new Alert(Alert.AlertType.INFORMATION);
        alert.setContentText("Ios Prefix 1 : " + iosPrefix1 + "\n" + "Ios Prefix 2 : " + iosPrefix2);
        alert.showAndWait();
    }

    public void onButtonClicked() {
        try {
            Stage mainStage = (Stage) mainPane.getScene().getWindow();
            SimpleDateFormat sdf = new SimpleDateFormat("yy_MM_dd_HH::mm::ss");
            fileName = sdf.format(new Date());
            ArrayList<Device> deviceList = new ArrayList<>();
            String regex = ".+\\.(xlsx)$";
            Pattern pattern = Pattern.compile(regex);
            FileChooser fileChooser = new FileChooser();
            File targetFile = fileChooser.showOpenDialog(mainStage);
            Matcher matcher = pattern.matcher(targetFile.getName());
            XSSFWorkbook workbook = new XSSFWorkbook(targetFile);
            Cell cell = null;
            int nameIndex = 0, versionIndex = 0, sessionIndex = 0;

            if (!matcher.matches())
                alert.showAndWait();
            else {
                Row row = workbook.getSheet("데이터세트1").getRow(1);
                String os = row.getCell(0) + "";
                isIos = (os.contains(iosPrefix1) || os.contains(iosPrefix2));

                for (Row r : workbook.getSheet("데이터세트1")) {

                    if (r.getCell(0) != null) {
                        if (r.getCell(0).getStringCellValue().equals("휴대기기 정보")) {
                            for (int i = 0; i < r.getLastCellNum(); i++) {
                                if (r.getCell(i) != null) {
                                    String result = r.getCell(i).getStringCellValue();
                                    switch (result) {
                                        case "휴대기기 정보":
                                            nameIndex = i;
                                            break;
                                        case "운영체제 버전":
                                            versionIndex = i;
                                            break;
                                        case "세션":
                                            sessionIndex = i;
                                    }
                                }
                            }
                            title = new String[3];
                            title[0] = r.getCell(nameIndex) + "";
                            title[1] = r.getCell(versionIndex) + "";
                            title[2] = r.getCell(sessionIndex) + "";
                            continue;   // 휴대기기 정보 및 세션 같은 데이터 값이 아닌 경우를 의미함
                        }
                    } else {
                        continue;   // null 값 있을때 즉, 총 세션수가 있을때를 의미함
                    }

                    Device device = new Device(r.getCell(nameIndex) + "", r.getCell(versionIndex) + "", (int) Double.parseDouble(r.getCell(sessionIndex) + ""));
                    deviceList.add(device);
                }

                DescendingSession descending = new DescendingSession();
                Collections.sort(deviceList, descending);


                if (isIos)
                    fileName = IOS + fileName + ".xlsx";
                else
                    fileName = AND + fileName + ".xlsx";
                File resultFile = new File(System.getProperty("user.home") + File.separator + "Desktop" + File.separator, fileName);
                FileOutputStream fos = new FileOutputStream(resultFile);

                XSSFWorkbook resultBook = new XSSFWorkbook();
                XSSFSheet sheet = null;
                SimpleDateFormat data = new SimpleDateFormat("yyMMdd");

                if (isIos)
                    sheet = resultBook.createSheet(IOS + data.format(new Date()));
                else
                    sheet = resultBook.createSheet(AND + data.format(new Date()));

                XSSFRow resultRow = null;
                XSSFCell resultCell = null;

                XSSFCellStyle style = resultBook.createCellStyle();
                style.setBorderTop(BorderStyle.THIN);
                style.setBorderBottom(BorderStyle.THIN);
                style.setBorderLeft(BorderStyle.THIN);
                style.setBorderRight(BorderStyle.THIN);
                style.setAlignment(HorizontalAlignment.CENTER);

                int rowNo = 0, cellNo = 0;

                resultRow = sheet.createRow(rowNo++);
                if (!isIos) {
                    resultCell = resultRow.createCell(cellNo);
                    resultCell.setCellStyle(style);
                    resultCell.setCellValue(title[0]);
                    sheet.autoSizeColumn(cellNo++);
                }
                resultCell = resultRow.createCell(cellNo);
                resultCell.setCellStyle(style);
                resultCell.setCellValue(title[1]);
                sheet.autoSizeColumn(cellNo++);
                resultCell = resultRow.createCell(cellNo);
                resultCell.setCellStyle(style);
                resultCell.setCellValue(title[2]);
                sheet.autoSizeColumn(cellNo);

                for (Device d : deviceList) {
                    cellNo = 0;
                    resultRow = sheet.createRow(rowNo++);
                    if (!isIos) {
                        resultCell = resultRow.createCell(cellNo);
                        resultCell.setCellStyle(style);
                        resultCell.setCellValue(d.getName());
                        sheet.autoSizeColumn(cellNo++);
                    }
                    resultCell = resultRow.createCell(cellNo);
                    resultCell.setCellStyle(style);
                    resultCell.setCellValue(d.getVersion());
                    sheet.autoSizeColumn(cellNo++);
                    resultCell = resultRow.createCell(cellNo);
                    resultCell.setCellStyle(style);
                    resultCell.setCellValue(d.getSession());
                    sheet.autoSizeColumn(cellNo);
                }
                resultBook.write(fos);
                resultBook.close();
                fos.close();
            }
            Alert alert = new Alert(Alert.AlertType.INFORMATION, "작업이 완료되었습니다.");
            alert.showAndWait();
        } catch (IOException | InvalidFormatException e) {
            e.printStackTrace();
        }
    }

    private void setAlert() {
        alert = new Alert(Alert.AlertType.ERROR, "file isn't xlsx file");
    }

    class DescendingSession implements Comparator<Device> {

        @Override
        public int compare(Device o1, Device o2) {
            return o2.getSession().compareTo(o1.getSession());
        }
    }
}
