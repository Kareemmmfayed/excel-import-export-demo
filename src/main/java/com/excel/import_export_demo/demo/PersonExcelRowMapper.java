package com.excel.import_export_demo.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Component;

import com.excel.import_export_demo.excel.ExcelImportService;
import com.excel.import_export_demo.excel.ExcelRowMapper;

import lombok.RequiredArgsConstructor;

@Component
@RequiredArgsConstructor
public class PersonExcelRowMapper implements ExcelRowMapper<PersonRequest> {

    @Override
    public PersonRequest mapRow(Row row) {
        String name = getString(row, 0);
        int age = Integer.parseInt(getString(row, 1));

        return new PersonRequest(name, age);
    }

	private String getString(Row row, int col) {
		Cell cell = row.getCell(col);
		return ExcelImportService.getCellValueAsString(cell);
	}

}
