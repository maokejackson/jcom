package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookAutoDiscoverConnectionMode implements Constant
{
    EXTERNAL(1),
    INTERNAL(2),
    INTERNAL_DOMAIN(3),
    UNKNOWN(0),
    ;

    private final int value;

    OutlookAutoDiscoverConnectionMode(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookAutoDiscoverConnectionMode findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookAutoDiscoverConnectionMode.class, value);
    }
}
