package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookBusyStatus implements Constant
{
    FREE(0),
    TENTATIVE(1),
    BUSY(2),
    OUT_OF_OFFICE(3),
    ;

    private final int value;

    OutlookBusyStatus(int value) {this.value = value;}

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookBusyStatus findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookBusyStatus.class, value);
    }
}
