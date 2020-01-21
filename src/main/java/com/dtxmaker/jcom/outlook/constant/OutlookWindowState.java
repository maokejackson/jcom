package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookWindowState implements Constant
{
    MAXIMIZED(0),
    MINIMIZED(1),
    NORMAL_WINDOW(2),
    ;

    private final int value;

    OutlookWindowState(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookWindowState findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookWindowState.class, value);
    }
}
