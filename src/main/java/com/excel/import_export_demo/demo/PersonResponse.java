package com.excel.import_export_demo.demo;

import lombok.Builder;
import lombok.Getter;
import lombok.Setter;

@Getter
@Setter
@Builder
public class PersonResponse {
    private String name;

    private int age;
}
