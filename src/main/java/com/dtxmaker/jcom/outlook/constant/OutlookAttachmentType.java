package com.dtxmaker.jcom.outlook.constant;

import com.dtxmaker.jcom.Constant;
import com.dtxmaker.jcom.util.EnumUtils;

public enum OutlookAttachmentType implements Constant
{
    BY_VALUE(1),
    BY_REFERENCE(4),
    EMBEDDED_ITEM(5),
    OLE(6),
    ;

    private final int value;

    OutlookAttachmentType(int value)
    {
        this.value = value;
    }

    @Override
    public int getValue()
    {
        return value;
    }

    public static OutlookAttachmentType findByValue(int value)
    {
        return EnumUtils.findByValue(OutlookAttachmentType.class, value);
    }
}
