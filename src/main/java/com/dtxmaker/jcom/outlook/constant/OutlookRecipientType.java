package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookRecipientType implements Constant
{
    ORIGINATOR(0),
    TO(1),
    CC(2),
    BCC(3),
    ;

    private final int value;

    OutlookRecipientType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookRecipientType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookRecipientType.class, value);
    }
}
