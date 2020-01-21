package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookFolderTypeMode implements Constant
{
    NORMAL(0),
    FOLDER_ONLY(1),
    NO_NAVIGATION(2),
    ;

    private final int value;

    OutlookFolderTypeMode(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookFolderTypeMode findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookFolderTypeMode.class, value);
    }
}
