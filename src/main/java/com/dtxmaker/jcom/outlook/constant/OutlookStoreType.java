package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookStoreType implements Constant
{
    DEFAULT(1),
    UNICODE(2),
    ANSI(3),
    ;

    private final int value;

    OutlookStoreType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookStoreType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookStoreType.class, value);
    }
}
