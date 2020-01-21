package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookDisplayType implements Constant
{
    USER(0),
    DIST_LIST(1),
    FORUM(2),
    AGENT(3),
    ORGANIZATION(4),
    PRIVATE_DIST_LIST(5),
    REMOTE_USER(6),
    ;

    private final int value;

    OutlookDisplayType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookDisplayType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookDisplayType.class, value);
    }
}
