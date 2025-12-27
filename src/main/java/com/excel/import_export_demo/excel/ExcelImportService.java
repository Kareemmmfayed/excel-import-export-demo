package com.excel.import_export_demo.excel;

import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

@Service
public class ExcelImportService {

    public <T> List<T> importFronExcel(MultipartFile file, ExcelRowMapper<T> rowMapper) {
        List<T> results = new ArrayList<>();

        try (
            InputStream is = file.getInputStream();
            Workbook workbook = new XSSFWorkbook(is);
        ) {

            Sheet sheet = workbook.getSheetAt(0);
            Iterator<Row> row = sheet.iterator();

            if(row.hasNext()) row.next();

            while(row.hasNext()) {
                Row currentRow = row.next();

                if(isRowEmpty(currentRow)) continue;

                results.add(rowMapper.mapRow(currentRow));
            }
            
        } catch (Exception e) {
            System.out.println(e);
        }

        return results;
    }

    private Boolean isRowEmpty(Row row) {
        if(row == null) return true;

        for(int i = row.getFirstCellNum(); i <= row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);

            if(cell != null && cell.getCellType() != CellType.BLANK) return false;
        }

        return true;
    }

    public static String getCellValueAsString(Cell cell) {
		if (cell == null) return "";
		
		return switch (cell.getCellType()) {
		case STRING -> cell.getStringCellValue().trim();
		case NUMERIC -> {
			if (DateUtil.isCellDateFormatted(cell)) {
				yield cell.getDateCellValue().toString();
			} else {
				BigDecimal bd = BigDecimal.valueOf(cell.getNumericCellValue());
				yield bd.stripTrailingZeros().toPlainString();
			}
		}
		case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
		default -> "";
		};
	}
}
