package com.example.demo.service.thongTinChiTietLoi;

import com.example.demo.model.ExcelColumn;
import com.example.demo.model.ReadExcelDTO;
import com.example.demo.model.TongHopLoiChamCong;
import com.example.demo.service.impl.InsertExcelService;
import com.example.demo.service.soLoiChamCong.TinhLoiService;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.*;

import static java.util.stream.Collectors.groupingBy;

@Service
public class ReadExcelService {
    private final ObjectMapper mapper;
    private final TinhLoiChamCong tinhLoiChamCong;
    private final TinhLoiService tinhLoiService;
    private final InsertExcelService insertExcelService;
    private static final DataFormatter formatter = new DataFormatter();

    public ReadExcelService(ObjectMapper mapper, TinhLoiChamCong tinhLoiChamCong, TinhLoiService tinhLoiService, InsertExcelService insertExcelService) {
        this.mapper = mapper;
        this.tinhLoiChamCong = tinhLoiChamCong;
        this.tinhLoiService = tinhLoiService;
        this.insertExcelService = insertExcelService;
    }

    public byte[] readExcelData(MultipartFile file, Integer number) throws Exception {
        if (!isXlsx(file)) {
            throw new Exception("Lỗi định dạng tệp");
        }
        List<Object> records = new LinkedList<>();
        try (InputStream inputStream = file.getInputStream()) {
            Workbook workbook = new XSSFWorkbook(inputStream);
            boolean isStop = true;
            int totalNumberOfSheet = workbook.getNumberOfSheets();
            int numberOfSheet = 0;
            Sheet sheet = null;
            while (isStop) {
                if (numberOfSheet == totalNumberOfSheet) {
                    throw new Exception("Sheet excel lỗi");
                }
                sheet = workbook.getSheetAt(numberOfSheet);
                if (sheet != null && sheet.getSheetName() != null && sheet.getSheetName().equals("DanhSachChamCong")) {
                    isStop = false;
                }
                numberOfSheet++;
            }
            for (Row row : sheet) {
                if (row.getRowNum() < 1) continue; // Skip headers
                ReadExcelDTO record = new ReadExcelDTO();
                for (Field field : ReadExcelDTO.class.getDeclaredFields()) {
                    if (field.isAnnotationPresent(ExcelColumn.class)) {
                        int column = field.getAnnotation(ExcelColumn.class).value();
                        field.setAccessible(true);
                        Cell cell = row.getCell(column - 1);
                        String cellValue = getCellValueAsString(cell);
                        if (!cellValue.trim().isEmpty()) {
                            if (column != 2) {
                                field.set(record, cellValue);
                            } else {
                                field.set(record, formatDateTime(cellValue, "MM/dd/yyyy HH:mm", "MM/dd/yyyy HH:mm"));
                                record.setNgay(formatDateTime(cellValue, "MM/dd/yyyy HH:mm", "dd/MM/yyyy"));
                            }
                        }
                    }
                }
                records.add(record);
            }
        } catch (IOException | IllegalAccessException e) {
            e.printStackTrace();
        }
        Map<String, Map<String, List<ReadExcelDTO>>> groupMcc = groupMCC(removeEmptyObjects(records));
        List<TongHopLoiChamCong> lists = new ArrayList<>();
        if (number == 1) {
            lists = tinhLoiChamCong.processChamCong(groupMcc);
        }
        if (number == 2) {
            lists = tinhLoiService.processChamCong(groupMcc);
        }
        return insertExcelService.generateExcelAsByteArray(lists);
    }

//    public static Date stringToDate(String date, String format) throws Exception {
//        try {
//            SimpleDateFormat simpleDateFormat = new SimpleDateFormat(format);
//            simpleDateFormat.setLenient(false);
//            return simpleDateFormat.parse(date);
//        } catch (Exception e) {
//            throw new Exception("Lỗi định dạng thời gian");
//        }
//    }

    public static String formatDateTime(String dateStr, String formatInput, String formatOutput) throws Exception {
        SimpleDateFormat originalFormat = new SimpleDateFormat(formatInput);
        SimpleDateFormat targetFormat = new SimpleDateFormat(formatOutput);
        try {
            Date date = originalFormat.parse(dateStr);
            return targetFormat.format(date);
        } catch (Exception e) {
            throw new Exception("Lỗi định dạng thời gian");
        }
    }

    public Map<String, Map<String, List<ReadExcelDTO>>> groupMCC(List<Object> list) {
        if (list.size() > 0) {
            List<ReadExcelDTO> readExcelDTOList = mapper.convertValue(list, new TypeReference<List<ReadExcelDTO>>() {
            });
            return readExcelDTOList.stream().collect(groupingBy(ReadExcelDTO::getMaNhanVien, groupingBy(ReadExcelDTO::getNgay)));
        }
        return null;
    }


    private static String getCellValueAsString(Cell cell) {
        if (cell == null) return "";
        return formatter.formatCellValue(cell);
    }

    private boolean isXlsx(MultipartFile file) throws Exception {
        if (file != null && file.getSize() > 0) {
            String extension = buildExtensionFile(file.getOriginalFilename());
            return "xlsx".equalsIgnoreCase(extension);
        }
        throw new Exception("Tệp không tồn tại hoặc rỗng");
    }

    public static String buildExtensionFile(String originalFileName) throws Exception {
        if (originalFileName != null) {
            String[] arrayStr = originalFileName.split("\\.");
            if (arrayStr.length == 0) {
                return originalFileName;
            }
            return arrayStr[arrayStr.length - 1];
        }
        throw new Exception("OriginalFileName is null");
    }

    public static <T> List<T> removeEmptyObjects(List<T> list) {
        list.removeIf(ReadExcelService::isAllFieldsNullOrEmpty);
        return list;
    }

    private static <T> boolean isAllFieldsNullOrEmpty(T obj) {
        if (obj == null) {
            return true;
        }
        try {
            for (Field field : obj.getClass().getDeclaredFields()) {
                field.setAccessible(true);
                Object value = field.get(obj);
                if (value != null && !isEmpty(value)) {
                    return false;
                }
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        }
        return true;
    }

    private static boolean isEmpty(Object value) {
        if (value instanceof String) {
            return ((String) value).trim().isEmpty();
        }
        return false;
    }
}
// Chia thành 2 case sáng và chiều, nếu như sáng tính như sau :
// Với buổi sáng lấy thời gian trong ngày từ 6:00 đến 13:30 và lấy thời gian sớm nhất
// Nếu như thoiGianChamCong trong khoảng 06:00 đến 08:00 thì trả ra 0
// Nếu từ 08:00 đến 10:00 là đi muộn thì trả ra 1
// Nếu từ 10:01 đến 13:30 là nghỉ buổi sáng => "nghỉ sáng"
// và sau 13:30 không tính chấm công
// Chiều tính như sau:
// Với buổi chiều lấy từ 13:30 - 23:59 và lấy thời gian muộn nhất
// Nếu 17:30 đến 23:59 thi về đúng giờ trả ra  0
// từ 15:30 đến 17:29 thì trả ra 1 lỗi
// từ 13:30 - 15:30 không tính công buổi chiều

// Nếu không co thông tin ngày đó => nghỉ cả ngay
// Lỗi sẽ được cộng dồn trong ngày

// kết quả tra ra có dạng như sau:
// nghỉ sáng + đi làm chiều chấm công sau 13h30 => không tính chấm công // chưa viet
// Nghỉ sáng + 1 lỗi
// 1 lỗi + Nghỉ chiều
// nghỉ sáng + nghỉ chiều
// làm đủ: >=0 lỗi
// nghỉ cả ngày


