package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookActionCopyLike implements Constant
{
    REPLY(0),
    REPLY_ALL(1),
    FORWARD(2),
    REPLY_FOLDER(3),
    RESPOND(4),
    ;

    private final int value;

    OutlookActionCopyLike(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookActionCopyLike findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookActionCopyLike.class, value);
    }
}
