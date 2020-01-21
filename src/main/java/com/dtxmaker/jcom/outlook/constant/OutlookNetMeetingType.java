package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookNetMeetingType implements Constant
{
    NET_MEETING(0),
    NET_SHOW(1),
    EXCHANGE_CONFERENCING(2),
    ;

    private final int value;

    OutlookNetMeetingType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookNetMeetingType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookNetMeetingType.class, value);
    }
}
