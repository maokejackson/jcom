package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookBusinessCardType implements Constant
{
    OUTLOOK(0),
    INTER_CONNECT(1),
    ;

    private final int value;

    OutlookBusinessCardType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookBusinessCardType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookBusinessCardType.class, value);
    }
}
