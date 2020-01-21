package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookFlagStatus implements Constant
{
    NO_FLAG(0),
    FLAG_COMPLETE(1),
    FLAG_MARKED(2),
    ;

    private final int value;

    OutlookFlagStatus(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookFlagStatus findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookFlagStatus.class, value);
    }
}
