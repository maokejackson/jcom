package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookMeetingRecipientType implements Constant
{
    ORGANIZER(0),
    REQUIRED(1),
    OPTIONAL(2),
    RESOURCE(3),
    ;

    private final int value;

    OutlookMeetingRecipientType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookMeetingRecipientType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookMeetingRecipientType.class, value);
    }
}
