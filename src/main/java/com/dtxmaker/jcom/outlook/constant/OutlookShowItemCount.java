package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookShowItemCount implements Constant
{
    NO(0),
    SHOW_UNREAD(1),
    SHOW_TOTAL(2),
    ;

    private final int value;

    OutlookShowItemCount(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookShowItemCount findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookShowItemCount.class, value);
    }
}
