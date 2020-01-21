package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookViewType implements Constant
{
    TABLE(0),
    CARD(1),
    CALENDAR(2),
    ICON(3),
    TIMELINE(4),
    ;

    private final int value;

    OutlookViewType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookViewType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookViewType.class, value);
    }
}
