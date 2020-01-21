package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookMeetingResponse implements Constant
{
    TENTATIVE(2),
    ACCEPTED(3),
    DECLINED(4),
    ;

    private final int value;

    OutlookMeetingResponse(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookMeetingResponse findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookMeetingResponse.class, value);
    }
}
