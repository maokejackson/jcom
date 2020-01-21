package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookBodyFormat implements Constant
{
    UNSPECIFIED(0),
    PLAIN(1),
    HTML(2),
    RICH_TEXT(3),
    ;

    private final int value;

    OutlookBodyFormat(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookBodyFormat findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookBodyFormat.class, value);
    }
}
