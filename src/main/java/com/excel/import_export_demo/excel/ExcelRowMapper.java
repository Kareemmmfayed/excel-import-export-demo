package com.excel.import_export_demo.excel;

import org.apache.poi.ss.usermodel.Row;

@FunctionalInterface
public interface ExcelRowMapper<T> {
    T mapRow(Row row);
}
