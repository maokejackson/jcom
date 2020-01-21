package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookSensitivity implements Constant
{
    NORMAL(0),
    PERSONAL(1),
    PRIVATE(2),
    CONFIDENTIAL(3),
    ;

    private final int value;

    OutlookSensitivity(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookSensitivity findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookSensitivity.class, value);
    }
}
