package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookTaskResponse implements Constant
{
    SIMPLE(0),
    ASSIGN(1),
    ACCEPT(2),
    DECLINE(3),
    ;

    private final int value;

    OutlookTaskResponse(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookTaskResponse findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookTaskResponse.class, value);
    }
}
