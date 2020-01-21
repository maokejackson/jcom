package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookResponseStatus implements Constant
{
    NONE(0),
    ORGANIZED(1),
    TENTATIVE(2),
    ACCEPTED(3),
    DECLINED(4),
    NOT_RESPONDED(5),
    ;

    private final int value;

    OutlookResponseStatus(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookResponseStatus findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookResponseStatus.class, value);
    }
}
