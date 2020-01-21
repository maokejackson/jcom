package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookRecurrenceType implements Constant
{
    DAILY(0),
    WEEKLY(1),
    MONTHLY(2),
    MONTH_NTH(3),
    YEARLY(5),
    YEAR_NTH(6),
    ;

    private final int value;

    OutlookRecurrenceType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookRecurrenceType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookRecurrenceType.class, value);
    }
}
