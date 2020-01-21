package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookSortOrder implements Constant
{
    SORT_NONE(0),
    ASCENDING(1),
    DESCENDING(2),
    ;

    private final int value;

    OutlookSortOrder(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookSortOrder findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookSortOrder.class, value);
    }
}
