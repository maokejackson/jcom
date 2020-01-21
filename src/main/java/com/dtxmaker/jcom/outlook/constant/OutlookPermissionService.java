package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookPermissionService implements Constant
{
    UNKNOWN(0),
    WINDOWS(1),
    PASSPORT(2),
    ;

    private final int value;

    OutlookPermissionService(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookPermissionService findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookPermissionService.class, value);
    }
}
