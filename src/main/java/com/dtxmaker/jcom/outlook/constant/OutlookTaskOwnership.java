package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookTaskOwnership implements Constant
{
    NEW_TASK(0),
    DELEGATED_TASK(1),
    OWN_TASK(2),
    ;

    private final int value;

    OutlookTaskOwnership(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookTaskOwnership findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookTaskOwnership.class, value);
    }
}
