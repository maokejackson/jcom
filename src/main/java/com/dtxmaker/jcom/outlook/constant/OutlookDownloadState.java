package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookDownloadState implements Constant
{
    HEADER_ONLY(0),
    FULL_ITEM(1),
    ;

    private final int value;

    OutlookDownloadState(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookDownloadState findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookDownloadState.class, value);
    }
}
