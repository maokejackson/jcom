package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookPermission implements Constant
{
    UNRESTRICTED(0),
    DO_NOT_FORWARD(1),
    PERMISSION_TEMPLATE(2),
    ;

    private final int value;

    OutlookPermission(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookPermission findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookPermission.class, value);
    }
}
