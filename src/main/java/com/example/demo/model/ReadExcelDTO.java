package com.example.demo.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.Date;

@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
public class ReadExcelDTO {
    @ExcelColumn(1)
    private String maNhanVien;
    @ExcelColumn(2)
    private String thoiGianChamCong;

    private String ngay;
}
