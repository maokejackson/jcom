package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookImportance implements Constant
{
    LOW(0),
    NORMAL(1),
    HIGH(2),
    ;

    private final int value;

    OutlookImportance(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookImportance findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookImportance.class, value);
    }
}
