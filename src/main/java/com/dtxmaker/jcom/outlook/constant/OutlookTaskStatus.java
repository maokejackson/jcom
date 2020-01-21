package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookTaskStatus implements Constant
{
    NOT_STARTED(0),
    IN_PROGRESS(1),
    COMPLETE(2),
    WAITING(3),
    DEFERRED(4),
    ;

    private final int value;

    OutlookTaskStatus(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookTaskStatus findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookTaskStatus.class, value);
    }
}
