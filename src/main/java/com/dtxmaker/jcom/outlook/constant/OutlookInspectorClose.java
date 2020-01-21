package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookInspectorClose implements Constant
{
    SAVE(0),
    DISCARD(1),
    PROMPT_FOR_SAVE(2),
    ;

    private final int value;

    OutlookInspectorClose(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookInspectorClose findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookInspectorClose.class, value);
    }
}
