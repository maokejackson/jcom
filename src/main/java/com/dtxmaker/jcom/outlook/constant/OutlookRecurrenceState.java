package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookRecurrenceState implements Constant
{
    NOT_RECURRING(0),
    MASTER(1),
    OCCURRENCE(2),
    EXCEPTION(3),
    ;

    private final int value;

    OutlookRecurrenceState(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookRecurrenceState findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookRecurrenceState.class, value);
    }
}
