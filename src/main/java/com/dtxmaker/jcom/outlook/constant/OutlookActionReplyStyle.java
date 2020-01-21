package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookActionReplyStyle implements Constant
{
    OMIT_ORIGINAL_TEXT(0),
    EMBED_ORIGINAL_ITEM(1),
    INCLUDE_ORIGINAL_TEXT(2),
    INDENT_ORIGINAL_TEXT(3),
    LINK_ORIGINAL_ITEM(4),
    USER_PREFERENCE(5),
    REPLY_TICK_ORIGINAL_TEXT(1000),
    ;

    private final int value;

    OutlookActionReplyStyle(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookActionReplyStyle findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookActionReplyStyle.class, value);
    }
}
