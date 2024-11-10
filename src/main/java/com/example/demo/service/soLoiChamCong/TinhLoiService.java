package com.example.demo.service.soLoiChamCong;

import com.example.demo.model.ErrorDTO;
import com.example.demo.model.ReadExcelDTO;
import com.example.demo.model.TongHopLoiChamCong;
import com.example.demo.service.common.CommonService;
import org.springframework.stereotype.Service;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

@Service
public class TinhLoiService {
    private CommonService commonService;
    private DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy HH:mm");

    public ErrorDTO processChamCongForDay(List<ReadExcelDTO> thoiGianChamCongList) throws Exception {
        LocalDateTime gioSangSomNhat = null;
        LocalDateTime gioChieuMuonNhat = null;

        for (ReadExcelDTO record : thoiGianChamCongList) {
            try {
                LocalDateTime chamCong = LocalDateTime.parse(record.getThoiGianChamCong(), formatter);
                int phut = chamCong.getHour() * 60 + chamCong.getMinute();
                // Xử lý cho buổi sáng từ 6:00 đến 13:30, lấy thời gian sớm nhất
                if (phut >= 6 * 60 && phut <= (13 * 60) + 30) {
                    if (gioSangSomNhat == null || chamCong.isBefore(gioSangSomNhat)) {
                        gioSangSomNhat = chamCong;
                    }
                }
                // Xử lý cho buổi chiều từ 13:30 đến 23:59, lấy thời gian muộn nhất
                if ((phut >= (13 * 60) + 30 && phut <= (23 * 60) + 59)) {
                    if (gioChieuMuonNhat == null || chamCong.isAfter(gioChieuMuonNhat)) {
                        gioChieuMuonNhat = chamCong;
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
                throw new Exception("Lỗi định dạng thời gian " + record.getThoiGianChamCong());
            }
        }
        // Gọi hàm xử lý cho buổi sáng và chiều
        ErrorDTO ketQuaSang = processMorning(gioSangSomNhat);
        ErrorDTO ketQuaChieu = processAfternoon(gioChieuMuonNhat, gioSangSomNhat);

        return new ErrorDTO(
                ketQuaSang.getDay() + ketQuaChieu.getDay(),
                ketQuaSang.isError() || ketQuaChieu.isError(),
                ketQuaSang.getTongSoPhutDiMuon() + ketQuaChieu.getTongSoPhutDiMuon()
        );
    }

    private ErrorDTO processMorning(LocalDateTime gioSang) {
        if (gioSang == null) {
//            return "Nghỉ sáng";
            return new ErrorDTO(0.0, false, 0.0);
        }
        int phut = (gioSang.getHour() * 60) + gioSang.getMinute();
        if (phut >= 6 * 60 && phut <= (8 * 60) + 5) {
//            return "Sáng đi làm: 0 lỗi";
            return new ErrorDTO(0.5, false, 0.0);
        } else if (phut <= 10 * 60) {
//            return "Sáng đi làm: 1 lỗi";
            return new ErrorDTO(0.5, true, phut - ((8 * 60) + 5));
        } else if (phut <= (13 * 60) + 30) {
            return new ErrorDTO(0.0, false, 0.0);
        }
        return new ErrorDTO(0.0, true, 0.0);
//        return "Không tính chấm công";
    }

    private ErrorDTO processAfternoon(LocalDateTime gioChieu, LocalDateTime gioSang) {
        if (gioChieu == null) {
            return new ErrorDTO(0.0, false, 0.0);
//            return "Nghỉ chiều";
        }
        int phut = (gioChieu.getHour() * 60) + gioChieu.getMinute();

        if (gioSang == null && phut > (13 * 60) + 30) {
            return new ErrorDTO(0.0, false, 0.0);
//            return "Không tính chấm công";
        }
        if (phut >= (17 * 60) + 30 && phut <= (23 * 60) + 59) {
            return new ErrorDTO(0.5, false, 0.0);
//            return "Chiều: 0 lỗi";
        } else if (phut >= (15 * 60) + 30 && phut <= (17 * 60) + 29) {
            return new ErrorDTO(0.5, true, ((17 * 60) + 30) - phut);
//            return "Chiều: 1 lỗi";
        }
        return new ErrorDTO(0.0, false, 0.0);
//        return "Không tính chấm công chiều";
    }

    public List<TongHopLoiChamCong> processChamCong(Map<String, Map<String, List<ReadExcelDTO>>> chamCongData) throws Exception {
        List<TongHopLoiChamCong> tongHopList = new ArrayList<>();

        for (String mcc : chamCongData.keySet()) {
            double tongHopPhutDiMuon = 0.0;
            TongHopLoiChamCong tongHop = new TongHopLoiChamCong();
            tongHop.setMCC(mcc);

            Map<String, List<ReadExcelDTO>> chamCongNgay = chamCongData.get(mcc);

            for (String ngay : chamCongNgay.keySet()) {
                if (!CommonService.isWeekend(ngay)) {
                    List<ReadExcelDTO> thoiGianChamCongList = chamCongNgay.get(ngay);
                    ErrorDTO ketQua = processChamCongForDay(thoiGianChamCongList);
                    tongHopPhutDiMuon += ketQua.getTongSoPhutDiMuon();
                    updateTongHop(tongHop, ngay, ketQua);
                }
            }
            tongHop.setTongHopSoPhutDiMuon(Double.toString(tongHopPhutDiMuon));
            tongHopList.add(tongHop);
        }

        return tongHopList;
    }

    private void updateTongHop(TongHopLoiChamCong tongHop, String ngay, ErrorDTO ketQua) throws Exception {
        String[] parts = ngay.split("/");
        int day = Integer.parseInt(parts[0]);
        String result = ketQua.getDay().toString();
        boolean isError = ketQua.isError();
        switch (day) {
            case 1:
                tongHop.setDay1(result);
                tongHop.setError1(isError);
                break;
            case 2:
                tongHop.setDay2(result);
                tongHop.setError2(isError);
                break;
            case 3:
                tongHop.setDay3(result);
                tongHop.setError3(isError);
                break;
            case 4:
                tongHop.setDay4(result);
                tongHop.setError4(isError);
                break;
            case 5:
                tongHop.setDay5(result);
                tongHop.setError5(isError);
                break;
            case 6:
                tongHop.setDay6(result);
                tongHop.setError6(isError);
                break;
            case 7:
                tongHop.setDay7(result);
                tongHop.setError7(isError);
                break;
            case 8:
                tongHop.setDay8(result);
                tongHop.setError8(isError);
                break;
            case 9:
                tongHop.setDay9(result);
                tongHop.setError9(isError);
                break;
            case 10:
                tongHop.setDay10(result);
                tongHop.setError10(isError);
                break;
            case 11:
                tongHop.setDay11(result);
                tongHop.setError11(isError);
                break;
            case 12:
                tongHop.setDay12(result);
                tongHop.setError12(isError);
                break;
            case 13:
                tongHop.setDay13(result);
                tongHop.setError13(isError);
                break;
            case 14:
                tongHop.setDay14(result);
                tongHop.setError14(isError);
                break;
            case 15:
                tongHop.setDay15(result);
                tongHop.setError15(isError);
                break;
            case 16:
                tongHop.setDay16(result);
                tongHop.setError16(isError);
                break;
            case 17:
                tongHop.setDay17(result);
                tongHop.setError17(isError);
                break;
            case 18:
                tongHop.setDay18(result);
                tongHop.setError18(isError);
                break;
            case 19:
                tongHop.setDay19(result);
                tongHop.setError19(isError);
                break;
            case 20:
                tongHop.setDay20(result);
                tongHop.setError20(isError);
                break;
            case 21:
                tongHop.setDay21(result);
                tongHop.setError21(isError);
                break;
            case 22:
                tongHop.setDay22(result);
                tongHop.setError22(isError);
                break;
            case 23:
                tongHop.setDay23(result);
                tongHop.setError23(isError);
                break;
            case 24:
                tongHop.setDay24(result);
                tongHop.setError24(isError);
                break;
            case 25:
                tongHop.setDay25(result);
                tongHop.setError25(isError);
                break;
            case 26:
                tongHop.setDay26(result);
                tongHop.setError26(isError);
                break;
            case 27:
                tongHop.setDay27(result);
                tongHop.setError27(isError);
                break;
            case 28:
                tongHop.setDay28(result);
                tongHop.setError28(isError);
                break;
            case 29:
                tongHop.setDay29(result);
                tongHop.setError29(isError);
                break;
            case 30:
                tongHop.setDay30(result);
                tongHop.setError30(isError);
                break;
            case 31:
                tongHop.setDay31(result);
                tongHop.setError31(isError);
                break;
            default:
                throw new Exception("Ngày không hợp lệ: " + ngay);
        }
    }

}
