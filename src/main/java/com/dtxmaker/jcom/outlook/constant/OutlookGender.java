package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookGender implements Constant
{
    UNSPECIFIED(0),
    FEMALE(1),
    MALE(2),
    ;

    private final int value;

    OutlookGender(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookGender findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookGender.class, value);
    }
}
