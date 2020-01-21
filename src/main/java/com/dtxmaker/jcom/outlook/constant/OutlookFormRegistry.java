package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookFormRegistry implements Constant
{
    DEFAULT(0),
    PERSONAL(2),
    FOLDER(3),
    ORGANIZATION(4),
    ;

    private final int value;

    OutlookFormRegistry(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookFormRegistry findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookFormRegistry.class, value);
    }
}
