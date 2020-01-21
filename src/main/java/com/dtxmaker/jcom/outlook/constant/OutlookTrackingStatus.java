package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookTrackingStatus implements Constant
{
    NONE(0),
    DELIVERED(1),
    NOT_DELIVERED(2),
    NOT_READ(3),
    RECALL_FAILURE(4),
    RECALL_SUCCESS(5),
    READ(6),
    REPLIED(7),
    ;

    private final int value;

    OutlookTrackingStatus(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookTrackingStatus findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookTrackingStatus.class, value);
    }
}
