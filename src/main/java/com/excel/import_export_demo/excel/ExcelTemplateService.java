package com.excel.import_export_demo.excel;

import java.io.IOException;

import org.apache.commons.io.output.ByteArrayOutputStream;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

@Service
public class ExcelTemplateService {

    public ResponseEntity<byte[]> generateTemplate(String[] headers, String fileName) throws IOException {
        Workbook wb = new XSSFWorkbook();
		Sheet sheet = wb.createSheet(fileName);

		Row header = sheet.createRow(0);

        for (int i = 0; i < headers.length; i++) {
			header.createCell(i).setCellValue(headers[i]);
		}

		ByteArrayOutputStream out = new ByteArrayOutputStream();
		wb.write(out);
		wb.close();

		return ResponseEntity.ok().header("Content-Disposition", "attachment; filename=" + fileName + ".xlsx")
				.body(out.toByteArray());
    }
}
