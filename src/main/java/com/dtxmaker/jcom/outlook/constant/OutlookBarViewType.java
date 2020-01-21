package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookBarViewType implements Constant
{
    LARGE_ICON(0),
    SMALL_ICON(1),
    ;

    private final int value;

    OutlookBarViewType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookBarViewType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookBarViewType.class, value);
    }
}
