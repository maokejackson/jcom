package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookMarkInterval implements Constant
{
    TODAY(0),
    TOMORROW(1),
    THIS_WEEK(2),
    NEXT_WEEK(3),
    NO_DATE(4),
    COMPLETE(5),
    ;

    private final int value;

    OutlookMarkInterval(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookMarkInterval findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookMarkInterval.class, value);
    }
}
