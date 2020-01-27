package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookAppointmentCopyOption implements Constant
{
    PROMPT_USER(0),
    CREATE_APPOINTMENT(1),
    COPY_AS_ACCEPT(2),
    ;

    private final int value;

    OutlookAppointmentCopyOption(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookAppointmentCopyOption findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookAppointmentCopyOption.class, value);
    }
}
