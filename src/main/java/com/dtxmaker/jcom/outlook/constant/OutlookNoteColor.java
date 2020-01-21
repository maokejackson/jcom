package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookNoteColor implements Constant
{
    BLUE(0),
    GREEN(1),
    PINK(2),
    YELLOW(3),
    WHITE(4),
    ;

    private final int value;

    OutlookNoteColor(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookNoteColor findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookNoteColor.class, value);
    }
}
