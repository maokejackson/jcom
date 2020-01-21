package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookConnectionMode implements Constant
{
    OFFLINE(100),
    LOW_BANDWIDTH(200),
    ONLINE(300),
    ;

    private final int value;

    OutlookConnectionMode(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookConnectionMode findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookConnectionMode.class, value);
    }
}
