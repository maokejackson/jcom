package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookActionShowOn implements Constant
{
    DONT_SHOW(0),
    MENU(1),
    MENU_AND_TOOLBAR(2),
    ;

    private final int value;

    OutlookActionShowOn(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookActionShowOn findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookActionShowOn.class, value);
    }
}
