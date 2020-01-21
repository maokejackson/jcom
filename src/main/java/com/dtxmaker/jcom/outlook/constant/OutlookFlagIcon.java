package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookFlagIcon implements Constant
{
    NO(0),
    PURPLE(1),
    ORANGE(2),
    GREEN(3),
    YELLOW(4),
    BLUE(5),
    RED(6),
    ;

    private final int value;

    OutlookFlagIcon(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookFlagIcon findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookFlagIcon.class, value);
    }
}
