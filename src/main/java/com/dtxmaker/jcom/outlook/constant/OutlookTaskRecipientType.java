package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookTaskRecipientType implements Constant
{
    UPDATE(2),
    FINAL_STATUS(3),
    ;

    private final int value;

    OutlookTaskRecipientType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookTaskRecipientType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookTaskRecipientType.class, value);
    }
}
