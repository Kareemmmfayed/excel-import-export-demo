package com.excel.import_export_demo.excel;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

@Service
public class ExcelExportSerivce {

    public <T> ResponseEntity<byte[]> exportToExcel(List<T> data, Class<T> clazz, String fileName) throws IllegalAccessException, IOException {

        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet(fileName);

        CellStyle cellStyle = createHeaderStyle(workbook);

        // this only works for classes not recored
        Field[] fields = clazz.getDeclaredFields();

        // Field[] fields = clazz.getRecordComponents() != null
        //         ? clazz.getRecordComponents()[0].getDeclaringRecord().getDeclaredFields()
        //         : clazz.getDeclaredFields();

        Row header = sheet.createRow(0);

        int col = 0;

        for(Field field : fields) {
            field.setAccessible(true);
            Cell cell = header.createCell(col++);
            cell.setCellValue(formatHeader(field.getName()));
            cell.setCellStyle(cellStyle);
        }

        int rowNum = 1;
        for(T item : data) {
            Row row = sheet.createRow(rowNum++);

            int colNum = 0;

            for(Field field : fields) {
                Cell cell = row.createCell(colNum++);
                Object value = field.get(item);
                cell.setCellValue(value == null ? "" : value.toString());
                cell.setCellStyle(cellStyle);
            }
        }

        for (int i = 0; i < fields.length; i++) {
            sheet.autoSizeColumn(i);
        }

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        workbook.write(out);
        workbook.close();

        byte[] bytes = out.toByteArray();

        return ResponseEntity.ok()
                .header("Content-Disposition", "attachment; filename=" + fileName)
                .header("Content-Type", "application/octet-stream")
                .body(bytes);

    } 

    private CellStyle createHeaderStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 12);

        style.setFont(font);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setBorderBottom(BorderStyle.THIN);

        return style;
    }

    private String formatHeader(String fieldName) {
        return fieldName.replaceAll("([a-z])([A-Z])", "$1 $2")
                .toUpperCase();
    }
}
