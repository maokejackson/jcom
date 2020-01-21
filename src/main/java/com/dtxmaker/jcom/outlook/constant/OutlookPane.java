package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookPane implements Constant
{
    OUTLOOK_BAR(1),
    FOLDER_LIST(2),
    PREVIEW(3),
    NAVIGATION_PANE(4),
    ;

    private final int value;

    OutlookPane(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookPane findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookPane.class, value);
    }
}
