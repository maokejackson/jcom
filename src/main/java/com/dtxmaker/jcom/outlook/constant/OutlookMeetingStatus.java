package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookMeetingStatus implements Constant
{
    NON_MEETING(0),
    MEETING(1),
    MEETING_RECEIVED(3),
    MEETING_CANCELED(5),
    ;

    private final int value;

    OutlookMeetingStatus(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookMeetingStatus findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookMeetingStatus.class, value);
    }
}
