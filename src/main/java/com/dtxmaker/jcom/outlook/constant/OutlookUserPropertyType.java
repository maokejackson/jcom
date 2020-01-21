package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookUserPropertyType implements Constant
{
    OUTLOOK_INTERNAL(0),
    TEXT(1),
    NUMBER(3),
    DATE_TIME(5),
    YES_NO(6),
    DURATION(7),
    KEYWORDS(11),
    PERCENT(12),
    CURRENCY(14),
    FORMULA(18),
    COMBINATION(19),
    ;

    private final int value;

    OutlookUserPropertyType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookUserPropertyType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookUserPropertyType.class, value);
    }
}
