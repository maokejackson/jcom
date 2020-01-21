package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookSyncState implements Constant
{
    STOPPED(0),
    STARTED(1),
    ;

    private final int value;

    OutlookSyncState(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookSyncState findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookSyncState.class, value);
    }
}
