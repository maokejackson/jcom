package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookItemType implements Constant
{
    MAIL(0),
    APPOINTMENT(1),
    CONTACT(2),
    TASK(3),
    JOURNAL(4),
    NOTE(5),
    POST(6),
    DISTRIBUTION_LIST(7),
    ;

    private final int value;

    OutlookItemType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookItemType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookItemType.class, value);
    }
}
