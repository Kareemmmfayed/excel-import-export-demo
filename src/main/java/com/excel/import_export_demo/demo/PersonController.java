package com.excel.import_export_demo.demo;

import java.io.IOException;
import java.util.List;

import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.excel.import_export_demo.excel.ExcelExportSerivce;
import com.excel.import_export_demo.excel.ExcelImportService;
import com.excel.import_export_demo.excel.ExcelTemplateService;

import lombok.RequiredArgsConstructor;

@RestController("/api/v1/excel/person")
@RequiredArgsConstructor
public class PersonController {

    private final PersonSerivce personSerivce;
    private final ExcelExportSerivce excelExportSerivce;
    private final ExcelImportService excelImportService;
    private final ExcelTemplateService excelTemplateService;
    private final PersonExcelRowMapper personExcelRowMapper;

    @GetMapping("/export")
    public ResponseEntity<byte[]> exportBuyOrders() throws Exception {

                List<PersonResponse> persons = personSerivce.getPersons();

                return excelExportSerivce.exportToExcel(
                persons,
                PersonResponse.class,
                "persons-" + System.currentTimeMillis() + ".xlsx"
                );
    }

    @PostMapping(value = "/import", consumes = "multipart/form-data")
	public ResponseEntity<List<Long>> importBuyOrders(@RequestParam("file") MultipartFile file) {
		try {
			List<PersonRequest> requests = excelImportService.importFronExcel(file, personExcelRowMapper);

			List<Long> createdIds = requests.stream().map(personSerivce::addPerson).toList();

			return ResponseEntity.ok(createdIds);

		} catch (Exception e) {
			System.out.println(e);
			return ResponseEntity.badRequest().body(List.of());
		}
	}

	@GetMapping("/import/template")
	public ResponseEntity<byte[]> getImportTemplate() throws IOException {
        String[] headers = {"NAME", "AGE"};
		
        return excelTemplateService.generateTemplate(headers, "persons");
	}

}
