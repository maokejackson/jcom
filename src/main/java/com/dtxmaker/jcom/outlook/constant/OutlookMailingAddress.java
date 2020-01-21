package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookMailingAddress implements Constant
{
    NONE(0),
    HOME(1),
    BUSINESS(2),
    OTHER(3),
    ;

    private final int value;

    OutlookMailingAddress(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookMailingAddress findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookMailingAddress.class, value);
    }
}
