package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookDaysOfWeek implements Constant
{
    SUNDAY(1),
    MONDAY(2),
    TUESDAY(4),
    WEDNESDAY(8),
    THURSDAY(16),
    FRIDAY(32),
    SATURDAY(64),
    ;

    private final int value;

    OutlookDaysOfWeek(int value) {this.value = value;}

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookDaysOfWeek findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookDaysOfWeek.class, value);
    }
}
