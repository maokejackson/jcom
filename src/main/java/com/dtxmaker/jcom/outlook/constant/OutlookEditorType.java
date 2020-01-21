package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookEditorType implements Constant
{
    TEXT(1),
    HTML(2),
    RTF(3),
    WORD(4),
    ;

    private final int value;

    OutlookEditorType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookEditorType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookEditorType.class, value);
    }
}
