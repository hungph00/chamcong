package com.example.demo.service.impl;

import com.example.demo.model.ExcelColumn;
import com.example.demo.model.TongHopLoiChamCong;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.List;

@Service
public class InsertExcelService {

    private CellStyle style;

    private CellStyle styleError;

    public byte[] generateExcelAsByteArray(List<TongHopLoiChamCong> list) throws Exception {
        Workbook workbook = createWorkbookWithData(list);
        return writeWorkbookToByteArray(workbook);
    }

    private Workbook createWorkbookWithData(List<TongHopLoiChamCong> list) throws Exception {
        InputStream i = this.getClass().getClassLoader().getResourceAsStream("static/tem_mau.xlsx");
        if (i == null) {
            throw new Exception("Lỗi tệp");
        }
        Workbook workbook = new XSSFWorkbook(i);
        Sheet sheet = workbook.getSheetAt(0);
        populateSheetWithData(sheet, workbook, list);
        i.close();
        return workbook;
    }

    private void populateSheetWithData(Sheet sheet, Workbook workbook, List<TongHopLoiChamCong> list) throws IllegalAccessException, NoSuchFieldException {
        int rowIndex = 1;
        int numericalOrder = 1;
        styleToCell(workbook);
        styleToCellError(workbook);
        for (TongHopLoiChamCong item : list) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) {
                row = sheet.createRow(rowIndex);
            }
            Cell sttCell = row.createCell(0);
            sttCell.setCellValue(numericalOrder++);

            populateRowWithObjectData(row, item, workbook);
//            handleSetStyleToCell(row, workbook);
            rowIndex++;
        }
    }

    private void populateRowWithObjectData(Row row, TongHopLoiChamCong obj, Workbook workbook) throws IllegalAccessException, NoSuchFieldException {
        Field[] fields = obj.getClass().getDeclaredFields();
        for (Field field : fields) {
            field.setAccessible(true);
//            if (field.get(obj) instanceof String || field.get(obj) instanceof Number) {
            ExcelColumn column = field.getAnnotation(ExcelColumn.class);
            if (column != null) {
                int colIndex = column.value() - 1; // ExcelColumn assumes 1-based index
                Cell cell = row.createCell(colIndex);
                cell.setCellValue(String.valueOf(field.get(obj)));
                if (colIndex > 0 && colIndex < 32) {
                    Field errorField = obj.getClass().getDeclaredField("error" + colIndex);
                    errorField.setAccessible(true);
                    boolean di = (boolean) errorField.get(obj);
                    if (di) {
                        cell.setCellStyle(styleError);
                    } else {
                        cell.setCellStyle(style);
                    }
                } else {
                    cell.setCellStyle(style);
                }
            }
//            }
//            else {
//                Object nestedObject = field.get(obj);
//                if (nestedObject != null) {
//                    populateRowWithObjectData(row, nestedObject, workbook);
//                }
//            }
        }
    }

    private void handleSetStyleToCell(Row row, Workbook workbook) {
        for (int cellIndex = 0; cellIndex < row.getLastCellNum(); cellIndex++) {
            Cell cell = row.getCell(cellIndex);
            if (cell != null) {
                cell.setCellStyle(style);
            }
        }
    }

    private void styleToCell(Workbook workbook) {
        style = workbook.createCellStyle();
        //Setting borders
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());

        // Setting text alignment to center
        style.setWrapText(true);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
    }

    private void styleToCellError(Workbook workbook) {
        styleError = workbook.createCellStyle();
        //Setting borders
        styleError.setBorderBottom(BorderStyle.THIN);
        styleError.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        styleError.setBorderLeft(BorderStyle.THIN);
        styleError.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        styleError.setBorderRight(BorderStyle.THIN);
        styleError.setRightBorderColor(IndexedColors.BLACK.getIndex());
        styleError.setBorderTop(BorderStyle.THIN);
        styleError.setTopBorderColor(IndexedColors.BLACK.getIndex());

        // Setting text alignment to center
        styleError.setWrapText(true);
        styleError.setAlignment(HorizontalAlignment.CENTER);
        styleError.setVerticalAlignment(VerticalAlignment.CENTER);

        styleError.setFillForegroundColor(IndexedColors.YELLOW.index); // Set the foreground color to yellow
        styleError.setFillPattern(FillPatternType.SOLID_FOREGROUND);     }

    private byte[] writeWorkbookToByteArray(Workbook workbook) throws IOException {
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
        return outputStream.toByteArray();
    }
}
