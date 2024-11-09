package com.example.demo.service.thongTinChiTietLoi;

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
public class TinhLoiChamCong {
    private DateTimeFormatter formatter = DateTimeFormatter.ofPattern("MM/dd/yyyy HH:mm");

    public String processChamCongForDay(List<ReadExcelDTO> thoiGianChamCongList) throws Exception {
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
        String ketQuaSang = processMorning(gioSangSomNhat);
        String ketQuaChieu = processAfternoon(gioChieuMuonNhat, gioSangSomNhat);

        return ketQuaSang + " + " + ketQuaChieu;
    }

    private String processMorning(LocalDateTime gioSang) {
        if (gioSang == null) {
            return "Nghỉ sáng";
        }
        int phut = (gioSang.getHour() * 60) + gioSang.getMinute();
        if (phut >= 6 * 60 && phut <= (8 * 60) + 5) {
            return "Sáng đi làm: 0 lỗi";
        } else if (phut <= 10 * 60) {
            return "Sáng đi làm: 1 lỗi";
        } else if (phut <= (13 * 60) + 30) {
            return "Nghỉ sáng";
        }
        return "Không tính chấm công";
    }

    private String processAfternoon(LocalDateTime gioChieu, LocalDateTime gioSang) {
        if (gioChieu == null) {
            return "Nghỉ chiều";
        }
        int phut = (gioChieu.getHour() * 60) + gioChieu.getMinute();

        if (gioSang == null && phut > (13 * 60) + 30) {
            return "Không tính chấm công";
        }

        if (phut >= (17 * 60) + 30 && phut <= (23 * 60) + 59) {
            return "Chiều: 0 lỗi";
        } else if (phut >= (15 * 60) + 30 && phut <= (17 * 60) + 29) {
            return "Chiều: 1 lỗi";
        }
        return "Không tính chấm công chiều";
    }

    public List<TongHopLoiChamCong> processChamCong(Map<String, Map<String, List<ReadExcelDTO>>> chamCongData) throws Exception {
        List<TongHopLoiChamCong> tongHopList = new ArrayList<>();

        for (String mcc : chamCongData.keySet()) {
            TongHopLoiChamCong tongHop = new TongHopLoiChamCong();
            tongHop.setMCC(mcc);

            Map<String, List<ReadExcelDTO>> chamCongNgay = chamCongData.get(mcc);

            for (String ngay : chamCongNgay.keySet()) {
                if (!CommonService.isWeekend(ngay)) {
                    List<ReadExcelDTO> thoiGianChamCongList = chamCongNgay.get(ngay);
                    String ketQua = processChamCongForDay(thoiGianChamCongList);

                    updateTongHop(tongHop, ngay, ketQua);
                }
            }
            tongHopList.add(tongHop);
        }

        return tongHopList;
    }

    private void updateTongHop(TongHopLoiChamCong tongHop, String ngay, String ketQua) throws Exception {
        String[] parts = ngay.split("/");
        int day = Integer.parseInt(parts[0]);

        switch (day) {
            case 1:
                tongHop.setDay1(ketQua);
                break;
            case 2:
                tongHop.setDay2(ketQua);
                break;
            case 3:
                tongHop.setDay3(ketQua);
                break;
            case 4:
                tongHop.setDay4(ketQua);
                break;
            case 5:
                tongHop.setDay5(ketQua);
                break;
            case 6:
                tongHop.setDay6(ketQua);
                break;
            case 7:
                tongHop.setDay7(ketQua);
                break;
            case 8:
                tongHop.setDay8(ketQua);
                break;
            case 9:
                tongHop.setDay9(ketQua);
                break;
            case 10:
                tongHop.setDay10(ketQua);
                break;
            case 11:
                tongHop.setDay11(ketQua);
                break;
            case 12:
                tongHop.setDay12(ketQua);
                break;
            case 13:
                tongHop.setDay13(ketQua);
                break;
            case 14:
                tongHop.setDay14(ketQua);
                break;
            case 15:
                tongHop.setDay15(ketQua);
                break;
            case 16:
                tongHop.setDay16(ketQua);
                break;
            case 17:
                tongHop.setDay17(ketQua);
                break;
            case 18:
                tongHop.setDay18(ketQua);
                break;
            case 19:
                tongHop.setDay19(ketQua);
                break;
            case 20:
                tongHop.setDay20(ketQua);
                break;
            case 21:
                tongHop.setDay21(ketQua);
                break;
            case 22:
                tongHop.setDay22(ketQua);
                break;
            case 23:
                tongHop.setDay23(ketQua);
                break;
            case 24:
                tongHop.setDay24(ketQua);
                break;
            case 25:
                tongHop.setDay25(ketQua);
                break;
            case 26:
                tongHop.setDay26(ketQua);
                break;
            case 27:
                tongHop.setDay27(ketQua);
                break;
            case 28:
                tongHop.setDay28(ketQua);
                break;
            case 29:
                tongHop.setDay29(ketQua);
                break;
            case 30:
                tongHop.setDay30(ketQua);
                break;
            case 31:
                tongHop.setDay31(ketQua);
                break;
            default:
                throw new Exception("Ngày không hợp lệ: " + ngay);
        }
    }
}

