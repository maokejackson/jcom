package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;

public enum OutlookAttachmentBlockLevel implements Constant
{
    NONE(0),
    OPEN(1),
    ;

    private final int value;

    OutlookAttachmentBlockLevel(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }
}
