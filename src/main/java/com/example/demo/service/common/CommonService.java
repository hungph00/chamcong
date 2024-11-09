package com.example.demo.service.common;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

public class CommonService {
    public static boolean isWeekend(String dateStr) {
        SimpleDateFormat sdf = new SimpleDateFormat("dd/MM/yyyy");
        try {
            Date date = sdf.parse(dateStr);
            Calendar calendar = Calendar.getInstance();
            calendar.setTime(date);

            int dayOfWeek = calendar.get(Calendar.DAY_OF_WEEK);
            return dayOfWeek == Calendar.SATURDAY || dayOfWeek == Calendar.SUNDAY;
        } catch (ParseException e) {
            e.printStackTrace();
            return false;
        }
    }
}
