package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookActionResponseStyle implements Constant
{
    OPEN(0),
    SEND(1),
    PROMPT(2),
    ;

    private final int value;

    OutlookActionResponseStyle(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookActionResponseStyle findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookActionResponseStyle.class, value);
    }
}
