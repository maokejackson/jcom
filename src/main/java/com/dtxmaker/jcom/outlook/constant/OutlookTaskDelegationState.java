package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookTaskDelegationState implements Constant
{
    NOT_DELEGATED(0),
    DELEGATION_UNKNOWN(1),
    DELEGATION_ACCEPTED(2),
    DELEGATION_DECLINED(3),
    ;

    private final int value;

    OutlookTaskDelegationState(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookTaskDelegationState findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookTaskDelegationState.class, value);
    }
}
