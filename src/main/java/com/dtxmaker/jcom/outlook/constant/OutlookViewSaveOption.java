package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookViewSaveOption implements Constant
{
    THIS_FOLDER_EVERYONE(0),
    THIS_FOLDER_ONLY_ME(1),
    ALL_FOLDERS_OF_TYPE(2),
    ;

    private final int value;

    OutlookViewSaveOption(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookViewSaveOption findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookViewSaveOption.class, value);
    }
}
