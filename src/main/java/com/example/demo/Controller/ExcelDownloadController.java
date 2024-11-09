package com.example.demo.Controller;

import com.example.demo.service.thongTinChiTietLoi.ReadExcelService;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class ExcelDownloadController {
    private final ReadExcelService readExcelService;

    public ExcelDownloadController(ReadExcelService readExcelService) {
        this.readExcelService = readExcelService;
    }

//    private final ExcelModifier excelModifier;
//
//    public ExcelDownloadController(ExcelModifier excelModifier) {
//        this.excelModifier = excelModifier;
//    }
//
//    @GetMapping("/downloadExcel")
//    public ResponseEntity<byte[]> downloadExcel() {
//        try {
//            byte[] excelFile = excelModifier.generateExcelAsByteArray();
//            HttpHeaders headers = new HttpHeaders();
//            headers.setContentType(new MediaType("application", "force-download")); // This will force the download in most browsers
//            headers.set(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=example.xlsx"); // Customize the file name as needed
//
//            return new ResponseEntity<>(excelFile, headers, HttpStatus.OK);
//        } catch (Exception e) {
//            return new ResponseEntity<>(null, HttpStatus.INTERNAL_SERVER_ERROR);
//        }
//    }

    @GetMapping("/chi-tiet-loi")
    public ResponseEntity<?> chiTietLoi(MultipartFile file) {
        try {
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            String header = "attachment; filename=chi-tiet-loi.xlsx";
            headers.set(HttpHeaders.CONTENT_DISPOSITION, header);
            return new ResponseEntity<>(readExcelService.readExcelData(file, 1), headers, HttpStatus.OK);
        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }

    @GetMapping("/tong-hop-loi")
    public ResponseEntity<?> tongHopLoi(MultipartFile file) {
        try {
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            String header = "attachment; filename=tong-hop-loi.xlsx";
            headers.set(HttpHeaders.CONTENT_DISPOSITION, header);
            return new ResponseEntity<>(readExcelService.readExcelData(file, 2), headers, HttpStatus.OK);
        } catch (Exception e) {
            e.printStackTrace();
            return new ResponseEntity<>(e.getMessage(), HttpStatus.INTERNAL_SERVER_ERROR);
        }
    }
}

